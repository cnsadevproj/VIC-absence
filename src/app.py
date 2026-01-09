"""
Cloud Run용 HTTP 앱

Cloud Scheduler에서 HTTP 요청을 받아 크롤링 실행
"""
import os
from datetime import datetime
from flask import Flask, request, jsonify

from scraper import run_scraper
from sheets import write_absence_records

app = Flask(__name__)

# 실행 허용 날짜 (MMDD 형식)
ALLOWED_DATES = {
    # 1월
    "0109", "0112", "0113", "0114", "0115", "0116",
    "0119", "0120", "0121", "0122", "0123",
    "0126", "0127", "0128", "0129", "0130",
    # 2월
    "0202", "0203"
}


def get_credentials():
    """환경변수에서 인증 정보 가져오기"""
    riro_user_id = os.environ.get("RIRO_USER_ID")
    riro_password = os.environ.get("RIRO_PASSWORD")
    google_creds = os.environ.get("GOOGLE_CREDENTIALS")

    if not riro_user_id or not riro_password:
        raise ValueError("RIRO_USER_ID, RIRO_PASSWORD 환경변수 필요")

    if not google_creds:
        raise ValueError("GOOGLE_CREDENTIALS 환경변수 필요")

    return riro_user_id, riro_password, google_creds


def is_allowed_date() -> bool:
    """오늘이 실행 허용 날짜인지 확인"""
    today = datetime.now().strftime("%m%d")
    return today in ALLOWED_DATES


def run_crawl(time_slot: str) -> dict:
    """크롤링 실행 및 스프레드시트 기록"""
    time_label = "오전 (1-4교시)" if time_slot == "morning" else "오후 (5-8교시)"
    now = datetime.now()
    now_str = now.strftime("%Y-%m-%d %H:%M:%S")
    today_mmdd = now.strftime("%m%d")

    # 허용 날짜 체크
    if not is_allowed_date():
        print(f"[{now_str}] 오늘({today_mmdd})은 실행 허용 날짜가 아닙니다. 스킵.")
        return {
            "skipped": True,
            "reason": f"오늘({today_mmdd})은 실행 허용 날짜가 아님",
            "timestamp": now_str
        }

    print(f"[{now_str}] {time_label} 크롤링 시작")

    riro_user_id, riro_password, credentials = get_credentials()

    # 스크래핑 실행 (오늘 날짜)
    print("리로스쿨에서 결석 데이터 수집 중...")
    records = run_scraper(riro_user_id, riro_password, time_slot)

    print(f"수집된 결석 학생 수: {len(records)}명")

    result = {
        "skipped": False,
        "time_slot": time_slot,
        "time_label": time_label,
        "timestamp": now_str,
        "total_count": len(records),
        "grade1_count": len([r for r in records if r.grade == 1]),
        "grade2_count": len([r for r in records if r.grade == 2]),
        "rows_written": 0
    }

    if records:
        # 스프레드시트 기록 (자동 append, start_row 없음)
        print("스프레드시트에 기록 중...")
        rows_written = write_absence_records(credentials, records, time_slot=time_slot)
        result["rows_written"] = rows_written
        print(f"완료! {rows_written}개 행 추가됨")
    else:
        print("결석 학생이 없습니다.")

    return result


@app.route("/", methods=["GET"])
def health():
    """헬스 체크"""
    return jsonify({"status": "ok"})


@app.route("/crawl/morning", methods=["POST", "GET"])
def crawl_morning():
    """오전 크롤링 (1-4교시) - 12:30 실행"""
    try:
        result = run_crawl("morning")
        return jsonify({"success": True, **result})
    except Exception as e:
        print(f"오류: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"success": False, "error": str(e)}), 500


@app.route("/crawl/afternoon", methods=["POST", "GET"])
def crawl_afternoon():
    """오후 크롤링 (5-8교시) - 16:30 실행"""
    try:
        result = run_crawl("afternoon")
        return jsonify({"success": True, **result})
    except Exception as e:
        print(f"오류: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"success": False, "error": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)
