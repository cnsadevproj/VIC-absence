"""
메인 실행 스크립트

환경변수:
- RIRO_USER_ID: 리로스쿨 아이디
- RIRO_PASSWORD: 리로스쿨 비밀번호
- GOOGLE_CREDENTIALS: 서비스 계정 JSON 문자열 (GitHub Actions용)
- GOOGLE_CREDENTIALS_PATH: 서비스 계정 JSON 파일 경로 (로컬 테스트용)
- TIME_SLOT: "morning" (오전, 1-4교시) 또는 "afternoon" (오후, 5-8교시)
"""
import os
import sys
import argparse

from scraper import run_scraper
from sheets import write_absence_records


def main():
    # 인자 파싱
    parser = argparse.ArgumentParser(description="리로스쿨 결석 데이터 수집 및 스프레드시트 기록")
    parser.add_argument(
        "--time-slot",
        choices=["morning", "afternoon"],
        default=os.environ.get("TIME_SLOT", "morning"),
        help="시간대 선택: morning (1-4교시) 또는 afternoon (5-8교시)"
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="스프레드시트에 기록하지 않고 데이터만 출력"
    )
    args = parser.parse_args()

    # 환경변수에서 인증 정보 가져오기
    riro_user_id = os.environ.get("RIRO_USER_ID")
    riro_password = os.environ.get("RIRO_PASSWORD")

    if not riro_user_id or not riro_password:
        print("오류: RIRO_USER_ID, RIRO_PASSWORD 환경변수를 설정해주세요.")
        sys.exit(1)

    # Google 인증 정보
    google_creds = os.environ.get("GOOGLE_CREDENTIALS")  # JSON 문자열
    google_creds_path = os.environ.get("GOOGLE_CREDENTIALS_PATH")  # 파일 경로

    if not google_creds and not google_creds_path:
        print("오류: GOOGLE_CREDENTIALS 또는 GOOGLE_CREDENTIALS_PATH 환경변수를 설정해주세요.")
        sys.exit(1)

    credentials = google_creds if google_creds else google_creds_path

    # 시간대 출력
    time_label = "오전 (1-4교시)" if args.time_slot == "morning" else "오후 (5-8교시)"
    print(f"시간대: {time_label}")
    print("-" * 50)

    # 스크래핑 실행
    print("리로스쿨에서 결석 데이터 수집 중...")
    try:
        records = run_scraper(riro_user_id, riro_password, args.time_slot)
    except Exception as e:
        print(f"스크래핑 오류: {e}")
        sys.exit(1)

    print(f"수집된 결석 학생 수: {len(records)}명")
    print()

    # 결과 출력
    grade1_records = [r for r in records if r.grade == 1]
    grade2_records = [r for r in records if r.grade == 2]

    if grade1_records:
        print("[ 1학년 결석 ]")
        for r in grade1_records:
            periods_str = ",".join(str(p) for p in r.periods)
            print(f"  {r.name}({r.student_id}) - {periods_str}교시")
        print()

    if grade2_records:
        print("[ 2학년 결석 ]")
        for r in grade2_records:
            periods_str = ",".join(str(p) for p in r.periods)
            print(f"  {r.name}({r.student_id}) - {periods_str}교시")
        print()

    if not records:
        print("결석 학생이 없습니다.")
        return

    # 스프레드시트 기록
    if args.dry_run:
        print("(dry-run 모드: 스프레드시트에 기록하지 않음)")
    else:
        print("스프레드시트에 기록 중...")
        try:
            rows_written = write_absence_records(credentials, records, time_slot=args.time_slot)
            print(f"완료! {rows_written}개 행 추가됨")
        except Exception as e:
            print(f"스프레드시트 기록 오류: {e}")
            sys.exit(1)


if __name__ == "__main__":
    main()
