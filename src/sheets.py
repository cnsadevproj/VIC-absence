"""
Google Sheets 연동 모듈
"""
import json
import re
from datetime import datetime
from pathlib import Path

import gspread
from google.oauth2.service_account import Credentials

from scraper import AbsenceRecord


# 스프레드시트 ID (URL에서 추출)
# https://docs.google.com/spreadsheets/d/1Gi2Qu_5nTba-pHfdRy6S_iyxXnhEiogxoxNFz7n37DY/edit
SPREADSHEET_ID = "1Gi2Qu_5nTba-pHfdRy6S_iyxXnhEiogxoxNFz7n37DY"

# 스코프 설정
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]


def load_student_data() -> tuple[dict, dict]:
    """학생 명렬 데이터 로드"""
    base_path = Path(__file__).parent.parent / "data"

    with open(base_path / "students_grade1.json", "r", encoding="utf-8") as f:
        grade1 = json.load(f)

    with open(base_path / "students_grade2.json", "r", encoding="utf-8") as f:
        grade2 = json.load(f)

    return grade1, grade2


def get_client(credentials_json: str | dict) -> gspread.Client:
    """Google Sheets 클라이언트 생성

    Args:
        credentials_json: JSON 파일 경로 또는 딕셔너리
    """
    if isinstance(credentials_json, str):
        # 파일 경로인 경우
        if credentials_json.startswith('{'):
            # JSON 문자열인 경우
            creds_dict = json.loads(credentials_json)
            creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        else:
            # 파일 경로인 경우
            creds = Credentials.from_service_account_file(credentials_json, scopes=SCOPES)
    else:
        # 딕셔너리인 경우
        creds = Credentials.from_service_account_info(credentials_json, scopes=SCOPES)

    return gspread.authorize(creds)


def get_today_existing_data(worksheet, today_date: str) -> dict:
    """오늘 날짜의 기존 데이터 조회

    Returns:
        {
            'grade1': {학번: {'row': 행번호, 'periods': "1,2교시"}},
            'grade2': {학번: {'row': 행번호, 'periods': "5,6교시"}},
            'max_seq': 최대 순번,
            'last_row': 마지막 데이터 행 번호
        }
    """
    all_values = worksheet.get_all_values()

    result = {
        'grade1': {},  # C열 학번 -> 행 정보
        'grade2': {},  # G열 학번 -> 행 정보
        'max_seq': 0,
        'last_row': 1  # 헤더 기본값
    }

    for idx, row in enumerate(all_values):
        row_num = idx + 1  # 1-based row number

        # 헤더 제외 (첫 행)
        if idx == 0:
            continue

        # 데이터가 있는 행 추적
        if len(row) >= 1 and row[0]:
            result['last_row'] = row_num

        # 오늘 날짜 데이터만 처리
        if len(row) >= 2 and row[0] == today_date:
            # 순번 최대값 추적
            try:
                seq = int(row[1])
                result['max_seq'] = max(result['max_seq'], seq)
            except ValueError:
                pass

            # 1학년 데이터 (C열=학번, F열=교시)
            if len(row) >= 6 and row[2]:
                student_id = row[2]
                periods = row[5] if len(row) > 5 else ""
                result['grade1'][student_id] = {
                    'row': row_num,
                    'periods': periods
                }

            # 2학년 데이터 (G열=학번, J열=교시)
            if len(row) >= 10 and row[6]:
                student_id = row[6]
                periods = row[9] if len(row) > 9 else ""
                result['grade2'][student_id] = {
                    'row': row_num,
                    'periods': periods
                }

    return result


def merge_periods(existing_periods: str, new_periods: list[int]) -> str:
    """기존 교시와 새 교시 병합

    예: "1,2교시" + [5, 6] -> "1,2,5,6교시"
    """
    # 기존 교시 파싱
    existing = set()
    if existing_periods:
        # "1,2교시" -> [1, 2]
        nums = re.findall(r'\d+', existing_periods.split('교시')[0])
        existing = set(int(n) for n in nums)

    # 새 교시 추가
    existing.update(new_periods)

    # 정렬 후 문자열로
    sorted_periods = sorted(existing)
    return ",".join(str(p) for p in sorted_periods) + "교시"


def get_cancelled_students(spreadsheet) -> set[str]:
    """취소자 명단 시트에서 학번 목록 가져오기

    Returns:
        취소자 학번 set
    """
    try:
        # "취소자 명단" 시트 찾기
        cancel_sheet = spreadsheet.worksheet("취소자 명단")
        # A열에서 2행부터 학번 가져오기
        student_ids = cancel_sheet.col_values(1)[1:]  # 헤더 제외
        return set(sid for sid in student_ids if sid.strip())
    except Exception as e:
        print(f"취소자 명단 시트 조회 실패: {e}")
        return set()


WEEKDAYS = ["월", "화", "수", "목", "금", "토", "일"]


def update_notification_message(
    spreadsheet,
    student_count: int,
    time_slot: str,
    target_date: str = None
):
    """알림 문구 시트의 B4 셀에 메시지 작성 (B3은 템플릿)

    Args:
        spreadsheet: gspread 스프레드시트 객체
        student_count: 결석 학생 수
        time_slot: "morning" 또는 "afternoon"
        target_date: 날짜 (YYYY-MM-DD), None이면 오늘
    """
    try:
        # 날짜 파싱
        if target_date:
            date_obj = datetime.strptime(target_date, "%Y-%m-%d")
        else:
            date_obj = datetime.now()

        month = date_obj.month
        day = date_obj.day
        weekday = WEEKDAYS[date_obj.weekday()]
        time_label = "오전" if time_slot == "morning" else "오후"

        # 메시지 작성
        message = f"""안녕하세요, 이현경 부장님.
{month}월 {day}일({weekday}) 겨울방학 방과후 출결결과 보내드립니다.

총 {student_count}명의 학생 및 학부모님께 {time_label} 알림 발송 완료했습니다.

[VIC 강의 출결 현황 스프레드시트] https://docs.google.com/spreadsheets/d/1Gi2Qu_5nTba-pHfdRy6S_iyxXnhEiogxoxNFz7n37DY/edit?usp=sharing"""

        # "알림 문구" 시트 찾기
        notify_sheet = spreadsheet.worksheet("알림 문구")
        # B4 셀에 메시지 작성 (B3은 템플릿으로 유지)
        notify_sheet.update(range_name="B4", values=[[message]], value_input_option="RAW")
        print(f"알림 문구 업데이트 완료 ({time_label}, {student_count}명)")

    except Exception as e:
        print(f"알림 문구 업데이트 실패: {e}")


def format_periods(periods: list[int]) -> str:
    """교시 리스트를 문자열로 변환

    예: [1, 2] -> "1,2교시"
        [1, 2, 3, 4] -> "1,2,3,4교시"
    """
    if not periods:
        return ""
    return ",".join(str(p) for p in periods) + "교시"


def write_absence_records(
    credentials_json: str | dict,
    records: list[AbsenceRecord],
    spreadsheet_id: str = SPREADSHEET_ID,
    target_date: str = None,
    start_row: int = None,
    time_slot: str = None
) -> int:
    """결석 기록을 스프레드시트에 작성

    - 오늘 날짜에 이미 있는 학생은 교시만 병합 (중복 행 방지)
    - 새 학생만 마지막 행 아래에 추가
    - 순번은 당일 기준 1부터 시작

    Args:
        credentials_json: 서비스 계정 JSON
        records: 결석 기록 리스트
        spreadsheet_id: 스프레드시트 ID
        target_date: 기록할 날짜 (YYYY-MM-DD), None이면 오늘
        start_row: 시작 행 번호, None이면 자동 추가 (deprecated)
        time_slot: "morning" 또는 "afternoon" (알림 문구용)

    Returns:
        작성된 행 수
    """
    if not records:
        print("작성할 결석 기록이 없습니다.")
        return 0

    # 클라이언트 및 스프레드시트 열기
    client = get_client(credentials_json)
    spreadsheet = client.open_by_key(spreadsheet_id)

    # 첫 번째 시트 사용 (또는 gid=626054561에 해당하는 시트)
    try:
        worksheet = spreadsheet.get_worksheet_by_id(626054561)
    except Exception:
        worksheet = spreadsheet.sheet1

    # 취소자 명단 가져오기
    cancelled_students = get_cancelled_students(spreadsheet)
    if cancelled_students:
        print(f"취소자 {len(cancelled_students)}명 제외")

    # 취소자 필터링
    filtered_records = [r for r in records if r.student_id not in cancelled_students]

    if len(filtered_records) < len(records):
        excluded = len(records) - len(filtered_records)
        print(f"  (취소자로 인해 {excluded}명 제외됨)")

    if not filtered_records:
        print("취소자 제외 후 작성할 결석 기록이 없습니다.")
        return 0

    # 학생 명렬 데이터 로드
    grade1_data, grade2_data = load_student_data()

    # 날짜 설정
    date_str = target_date if target_date else datetime.now().strftime("%Y-%m-%d")

    # 오늘 날짜의 기존 데이터 조회
    existing_data = get_today_existing_data(worksheet, date_str)
    print(f"기존 데이터: 1학년 {len(existing_data['grade1'])}명, 2학년 {len(existing_data['grade2'])}명")

    # 1학년, 2학년 분리
    grade1_records = [r for r in filtered_records if r.grade == 1]
    grade2_records = [r for r in filtered_records if r.grade == 2]

    # 중복 학생 분리 (기존에 있으면 업데이트, 없으면 신규)
    grade1_update = []  # 기존 학생 - 교시 업데이트
    grade1_new = []     # 신규 학생
    grade2_update = []
    grade2_new = []

    for r in grade1_records:
        if r.student_id in existing_data['grade1']:
            grade1_update.append(r)
        else:
            grade1_new.append(r)

    for r in grade2_records:
        if r.student_id in existing_data['grade2']:
            grade2_update.append(r)
        else:
            grade2_new.append(r)

    updated_count = 0
    new_rows_count = 0

    # 1. 기존 학생 교시 업데이트 (개별 셀 업데이트)
    for r in grade1_update:
        existing = existing_data['grade1'][r.student_id]
        merged = merge_periods(existing['periods'], r.periods)
        cell = f"F{existing['row']}"  # F열 = 1학년 교시
        worksheet.update(range_name=cell, values=[[merged]], value_input_option="USER_ENTERED")
        print(f"  1학년 {r.name} 교시 업데이트: {existing['periods']} -> {merged}")
        updated_count += 1

    for r in grade2_update:
        existing = existing_data['grade2'][r.student_id]
        merged = merge_periods(existing['periods'], r.periods)
        cell = f"J{existing['row']}"  # J열 = 2학년 교시
        worksheet.update(range_name=cell, values=[[merged]], value_input_option="USER_ENTERED")
        print(f"  2학년 {r.name} 교시 업데이트: {existing['periods']} -> {merged}")
        updated_count += 1

    # 2. 신규 학생 추가 (새 행)
    max_new_rows = max(len(grade1_new), len(grade2_new))

    if max_new_rows > 0:
        # 다음 순번 (당일 기준)
        next_seq = existing_data['max_seq'] + 1

        # 마지막 행 다음에 추가
        next_row = existing_data['last_row'] + 1

        rows_to_write = []
        for i in range(max_new_rows):
            row = [""] * 10  # A~J 열

            # A열: 날짜, B열: 순번
            row[0] = date_str
            row[1] = next_seq + i

            # 1학년 데이터 (C~F)
            if i < len(grade1_new):
                r = grade1_new[i]
                student_info = grade1_data.get(r.student_id, {})
                row[2] = r.student_id  # C: 학번
                row[3] = r.name  # D: 이름
                row[4] = student_info.get("type", "")  # E: 기숙/통학
                row[5] = format_periods(r.periods)  # F: 교시

            # 2학년 데이터 (G~J)
            if i < len(grade2_new):
                r = grade2_new[i]
                student_info = grade2_data.get(r.student_id, {})
                row[6] = r.student_id  # G: 학번
                row[7] = r.name  # H: 이름
                row[8] = student_info.get("type", "")  # I: 기숙/통학
                row[9] = format_periods(r.periods)  # J: 교시

            rows_to_write.append(row)

        # 데이터 작성
        end_row = next_row + len(rows_to_write) - 1
        range_str = f"A{next_row}:J{end_row}"
        worksheet.update(range_name=range_str, values=rows_to_write, value_input_option="USER_ENTERED")
        print(f"{next_row}행부터 {len(rows_to_write)}개의 신규 행 추가됨")
        new_rows_count = len(rows_to_write)

    # 알림 문구 업데이트 (time_slot이 지정된 경우)
    if time_slot:
        # 실제 결석 학생 수 (기존 + 신규 모두 포함)
        total_students = len(grade1_records) + len(grade2_records)
        update_notification_message(spreadsheet, total_students, time_slot, date_str)

    print(f"완료: 업데이트 {updated_count}명, 신규 추가 {new_rows_count}행")
    return new_rows_count


if __name__ == "__main__":
    # 테스트용
    import os

    creds_path = os.environ.get(
        "GOOGLE_CREDENTIALS_PATH",
        r"C:\Users\User\Downloads\vic-attendance-sms-6ec00f8eed28.json"
    )

    # 테스트 데이터
    test_records = [
        AbsenceRecord("21202", "김민서", 2, [1, 2]),
        AbsenceRecord("21220", "최진성", 2, [1, 2]),
        AbsenceRecord("10911", "김지몽", 1, [1, 2]),
    ]

    write_absence_records(creds_path, test_records)
