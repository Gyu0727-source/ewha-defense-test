#!/bin/bash
# 이화방어기제검사 자동화 시스템 설치 스크립트

set -e

echo "=========================================="
echo "이화방어기제검사 자동화 시스템 설치"
echo "=========================================="
echo ""

# Python 버전 확인
echo "1. Python 버전 확인..."
python3 --version

# 가상환경 생성
echo ""
echo "2. 가상환경 생성..."
if [ ! -d "venv" ]; then
    python3 -m venv venv
    echo "✅ 가상환경 생성 완료"
else
    echo "⚠️  가상환경이 이미 존재합니다"
fi

# 가상환경 활성화
echo ""
echo "3. 가상환경 활성화..."
source venv/bin/activate

# 라이브러리 설치
echo ""
echo "4. 필수 라이브러리 설치..."
pip install --upgrade pip
pip install -r requirements.txt

echo ""
echo "✅ 설치 완료!"
echo ""
echo "다음 단계:"
echo "1. Google Cloud Console에서 서비스 계정 생성"
echo "2. credentials.json 파일 다운로드 및 프로젝트 루트에 저장"
echo "3. .env.example을 복사하여 .env 파일 생성"
echo "4. .env 파일에 실제 값 입력"
echo ""
echo "사용 방법:"
echo "  source venv/bin/activate"
echo "  python3 src/edmt_automation.py --help"
echo ""
