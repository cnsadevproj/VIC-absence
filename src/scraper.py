"""
리로스쿨 결석 데이터 스크래핑 모듈
"""
import re
from dataclasses import dataclass
from playwright.sync_api import sync_playwright, Page


@dataclass
class AbsenceRecord:
    """결석 기록 데이터 클래스"""
    student_id: str
    name: str
    grade: int
    periods: list[int]


def login(page: Page, user_id: str, password: str) -> bool:
    """리로스쿨 로그인"""
    page.goto("https://cnsa.riroschool.kr")

    # 로그인 폼 입력
    page.fill('input[name="id"]', user_id)
    page.fill('input[name="pw"]', password)
    page.click('button[type="submit"], input[type="submit"], .btn_login, #login_btn')

    # 로그인 성공 확인 (페이지 이동 대기)
    page.wait_for_load_state("networkidle")

    # 로그인 성공 여부 확인
    return "로그아웃" in page.content() or "logout" in page.content().lower()


def parse_periods(period_text: str) -> list[int]:
    """교시 텍스트에서 교시 숫자 추출

    예: "1,2교시" -> [1, 2]
        "1,2,3,4교시" -> [1, 2, 3, 4]
    """
    # 숫자만 추출
    numbers = re.findall(r'\d+', period_text.split('교시')[0])
    return [int(n) for n in numbers]


def parse_grade(course_info: str) -> int:
    """과목 정보에서 학년 추출

    예: "[2학년 국제교육부 수목 1,2교시]" -> 2
        "[1학년 국어과 매일 1,2교시]" -> 1
        "[1+2학년 ...]" -> 0 (1,2학년 공통)
    """
    match = re.search(r'\[(\d)\+?(\d)?학년', course_info)
    if match:
        if match.group(2):  # 1+2학년 같은 경우
            return 0  # 공통 과목
        return int(match.group(1))
    return 0


def parse_period_from_course(course_info: str) -> list[int]:
    """과목 정보에서 교시 추출

    예: "[2학년 국제교육부 수목 1,2교시]" -> [1, 2]
    """
    match = re.search(r'(\d+(?:,\d+)*)교시\]', course_info)
    if match:
        return [int(n) for n in match.group(1).split(',')]
    return []


def parse_students(student_text: str) -> list[tuple[str, str]]:
    """결석생 텍스트 파싱

    예: "김민서(21202), 최진성(21220)" -> [("21202", "김민서"), ("21220", "최진성")]
    """
    students = []
    # 이름(학번) 패턴 매칭
    pattern = r'([가-힣]+)\((\d{5})\)'
    matches = re.findall(pattern, student_text)
    for name, student_id in matches:
        students.append((student_id, name))
    return students


def scrape_absence_data(page: Page, time_slot: str = "morning") -> list[AbsenceRecord]:
    """결석 데이터 스크래핑

    Args:
        page: Playwright 페이지 객체
        time_slot: "morning" (1-4교시) 또는 "afternoon" (5-8교시)

    Returns:
        AbsenceRecord 리스트
    """
    # 출결내역 조회 페이지로 이동
    page.goto("https://cnsa.riroschool.kr/lecture.php?db=1703&cate=34&action=stat")
    page.wait_for_load_state("networkidle")

    # 결석 라디오 버튼 클릭 (이미 선택되어 있을 수 있음)
    abs_radio = page.locator('input[value="abs"]')
    if abs_radio.count() > 0:
        abs_radio.first.click()
        page.wait_for_load_state("networkidle")

    # 테이블에서 데이터 추출
    rows = page.locator('table tbody tr').all()

    # 학생별 교시 집계를 위한 딕셔너리
    student_periods: dict[str, dict] = {}

    # 교시 필터링 기준
    if time_slot == "morning":
        valid_periods = {1, 2, 3, 4}
    else:  # afternoon
        valid_periods = {5, 6, 7, 8}

    for row in rows:
        cells = row.locator('td').all()
        if len(cells) < 6:
            continue

        # 과목 정보 (4번째 td)
        course_info = cells[3].inner_text().strip()

        # 결석생 명단 (6번째 td)
        student_text = cells[5].inner_text().strip()

        if not course_info or not student_text:
            continue

        # 학년 추출
        grade = parse_grade(course_info)

        # 교시 추출
        periods = parse_period_from_course(course_info)

        # 해당 시간대 교시만 필터링
        filtered_periods = [p for p in periods if p in valid_periods]
        if not filtered_periods:
            continue

        # 학생 파싱
        students = parse_students(student_text)

        for student_id, name in students:
            # 학번으로 학년 결정 (1학년: 1xxxx, 2학년: 2xxxx)
            actual_grade = 1 if student_id.startswith('1') else 2

            # 공통 과목(0)이거나 학년이 일치하는 경우만 처리
            if grade != 0 and grade != actual_grade:
                continue

            if student_id not in student_periods:
                student_periods[student_id] = {
                    'name': name,
                    'grade': actual_grade,
                    'periods': set()
                }

            student_periods[student_id]['periods'].update(filtered_periods)

    # AbsenceRecord 리스트로 변환
    records = []
    for student_id, data in student_periods.items():
        records.append(AbsenceRecord(
            student_id=student_id,
            name=data['name'],
            grade=data['grade'],
            periods=sorted(list(data['periods']))
        ))

    return records


def run_scraper(user_id: str, password: str, time_slot: str = "morning") -> list[AbsenceRecord]:
    """스크래퍼 실행 메인 함수"""
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()
        page = context.new_page()

        try:
            # 로그인
            if not login(page, user_id, password):
                raise Exception("로그인 실패")

            # 데이터 스크래핑
            records = scrape_absence_data(page, time_slot)

            return records
        finally:
            browser.close()


if __name__ == "__main__":
    # 테스트용
    import os
    user_id = os.environ.get("RIRO_USER_ID", "민수정")
    password = os.environ.get("RIRO_PASSWORD", "abcd123!")

    records = run_scraper(user_id, password, "morning")
    for r in records:
        print(f"{r.grade}학년 {r.name}({r.student_id}): {r.periods}교시")
