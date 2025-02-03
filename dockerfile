# 1️⃣ Python이 설치된 공식 이미지를 가져옴
FROM python:3.9-slim

# 2️⃣ 작업 디렉토리를 설정 (Vercel에서 사용할 폴더)
WORKDIR /app

# 3️⃣ 필요한 파일 복사 (Python 스크립트 및 패키지 설치 파일)
COPY decrypt_excel.py .
COPY requirements.txt .

# 4️⃣ 필요한 패키지 설치
RUN pip install -r requirements.txt

# 5️⃣ 서버 실행 (여기서는 Flask 사용 예시)
CMD ["python", "decrypt_excel.py"]