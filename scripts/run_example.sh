#!/bin/bash
# 이화방어기제검사 자동화 예제 실행 스크립트

set -e

# .env 파일 로드
if [ -f .env ]; then
    export $(cat .env | grep -v '^#' | xargs)
else
    echo "❌ .env 파일이 없습니다. .env.example을 복사하여 .env 파일을 생성하세요."
    exit 1
fi

# 가상환경 활성화 확인
if [ -z "$VIRTUAL_ENV" ]; then
    echo "⚠️  가상환경이 활성화되지 않았습니다. 활성화 중..."
    source venv/bin/activate
fi

echo "=========================================="
echo "이화방어기제검사 자동화 실행"
echo "=========================================="
echo ""
echo "Google Sheets URL: ${GOOGLE_SHEETS_URL:0:50}..."
echo "템플릿 파일: $TEMPLATE_PATH"
echo "출력 디렉토리: $OUTPUT_DIR"
echo ""

# 스크립트 실행
python3 src/edmt_automation.py \
    --sheet-url "$GOOGLE_SHEETS_URL" \
    --template "$TEMPLATE_PATH" \
    --output-dir "$OUTPUT_DIR" \
    --credentials "$CREDENTIALS_FILE" \
    --log-level "$LOG_LEVEL"

echo ""
echo "✅ 처리 완료!"
echo "결과 파일 위치: $OUTPUT_DIR"
