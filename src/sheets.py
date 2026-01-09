"""
Google Sheets 연동 모듈
"""
import json
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


def get_next_row_number(worksheet, today_date: str) -> int:
    """오늘 날짜의 다음 순번 계산

    A열에서 오늘 날짜가 있는 행들 중 B열의 최대값 + 1 반환
    """
    all_values = worksheet.get_all_values()

    max_num = 0
    for row in all_values[1:]:  # 헤더 제외
        if len(row) >= 2 and row[0] == today_date:
            try:
                num = int(row[1])
                max_num = max(max_num, num)
            except ValueError:
                pass

    return max_num + 1


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
    """알림 문구 시트의 B3 셀에 메시지 작성

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
        # B3 셀에 메시지 작성
        notify_sheet.update("B3", message, value_input_option="RAW")
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

    Args:
        credentials_json: 서비스 계정 JSON
        records: 결석 기록 리스트
        spreadsheet_id: 스프레드시트 ID
        target_date: 기록할 날짜 (YYYY-MM-DD), None이면 오늘
        start_row: 시작 행 번호, None이면 자동 추가
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

    # 1학년, 2학년 분리
    grade1_records = [r for r in filtered_records if r.grade == 1]
    grade2_records = [r for r in filtered_records if r.grade == 2]

    # 더 많은 쪽 기준으로 행 수 결정
    max_rows = max(len(grade1_records), len(grade2_records))

    if max_rows == 0:
        print("작성할 결석 기록이 없습니다.")
        return 0

    # 다음 순번 가져오기
    next_num = get_next_row_number(worksheet, date_str)

    # 작성할 데이터 준비
    rows_to_write = []

    for i in range(max_rows):
        row = [""] * 10  # A~J 열

        # A열: 날짜, B열: 순번
        row[0] = date_str
        row[1] = next_num + i

        # 1학년 데이터 (C~F)
        if i < len(grade1_records):
            r = grade1_records[i]
            student_info = grade1_data.get(r.student_id, {})
            row[2] = r.student_id  # C: 학번
            row[3] = r.name  # D: 이름
            row[4] = student_info.get("type", "")  # E: 기숙/통학
            row[5] = format_periods(r.periods)  # F: 교시

        # 2학년 데이터 (G~J)
        if i < len(grade2_records):
            r = grade2_records[i]
            student_info = grade2_data.get(r.student_id, {})
            row[6] = r.student_id  # G: 학번
            row[7] = r.name  # H: 이름
            row[8] = student_info.get("type", "")  # I: 기숙/통학
            row[9] = format_periods(r.periods)  # J: 교시

        rows_to_write.append(row)

    # 데이터 작성
    if start_row:
        # 특정 행부터 작성
        end_row = start_row + len(rows_to_write) - 1
        range_str = f"A{start_row}:J{end_row}"
        worksheet.update(range_str, rows_to_write, value_input_option="USER_ENTERED")
        print(f"{start_row}행부터 {len(rows_to_write)}개의 행이 작성되었습니다.")
    else:
        # 자동 추가
        worksheet.append_rows(rows_to_write, value_input_option="USER_ENTERED")
        print(f"{len(rows_to_write)}개의 행이 추가되었습니다.")

    # 알림 문구 업데이트 (time_slot이 지정된 경우)
    if time_slot:
        # 실제 결석 학생 수 (1학년 + 2학년)
        total_students = len(grade1_records) + len(grade2_records)
        update_notification_message(spreadsheet, total_students, time_slot, date_str)

    return len(rows_to_write)


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
