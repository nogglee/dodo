# 기본 이미지: Python이 포함된 Node.js 환경 사용
FROM node:18-bullseye

# Python 및 필요한 패키지 설치
RUN apt-get update && apt-get install -y python3 python3-pip

# 작업 디렉토리 설정
WORKDIR /app

# 의존성 파일 복사
COPY package*.json ./
RUN npm install

# 모든 코드 복사
COPY . .

# 실행 권한 부여
RUN chmod +x decrypt_excel.py

# 기본 실행 명령
CMD ["npm", "run", "start"]