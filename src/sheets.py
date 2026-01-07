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
    spreadsheet_id: str = SPREADSHEET_ID
) -> int:
    """결석 기록을 스프레드시트에 작성

    Args:
        credentials_json: 서비스 계정 JSON
        records: 결석 기록 리스트
        spreadsheet_id: 스프레드시트 ID

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

    # 학생 명렬 데이터 로드
    grade1_data, grade2_data = load_student_data()

    # 오늘 날짜
    today = datetime.now().strftime("%Y-%m-%d")

    # 1학년, 2학년 분리
    grade1_records = [r for r in records if r.grade == 1]
    grade2_records = [r for r in records if r.grade == 2]

    # 더 많은 쪽 기준으로 행 수 결정
    max_rows = max(len(grade1_records), len(grade2_records))

    if max_rows == 0:
        print("작성할 결석 기록이 없습니다.")
        return 0

    # 다음 순번 가져오기
    next_num = get_next_row_number(worksheet, today)

    # 작성할 데이터 준비
    rows_to_append = []

    for i in range(max_rows):
        row = [""] * 10  # A~J 열

        # A열: 날짜, B열: 순번
        row[0] = today
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

        rows_to_append.append(row)

    # 데이터 추가
    worksheet.append_rows(rows_to_append, value_input_option="USER_ENTERED")

    print(f"{len(rows_to_append)}개의 행이 추가되었습니다.")
    return len(rows_to_append)


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
