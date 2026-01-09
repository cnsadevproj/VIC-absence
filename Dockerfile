FROM mcr.microsoft.com/playwright/python:v1.40.0-jammy

WORKDIR /app

# 의존성 설치
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 소스 코드 복사
COPY src/ ./src/
COPY data/ ./data/

# Playwright 브라우저 설치
RUN playwright install chromium

# 포트 설정
ENV PORT=8080
EXPOSE 8080

# 앱 실행
CMD ["python", "src/app.py"]
