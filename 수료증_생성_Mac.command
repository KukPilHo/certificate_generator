#!/bin/bash
# 더블클릭하면 수료증을 일괄 생성합니다. (macOS)
# 이 파일은 프로젝트 폴더 안에 그대로 두세요.

# 1) 이 스크립트가 있는 폴더로 이동
cd "$(dirname "$0")" || exit 1
echo "작업 폴더: $(pwd)"
echo "----------------------------------------"

# 2) Python 인터프리터 선택 (venv가 실제로 실행되면 우선 사용)
if [ -x "venv/bin/python" ] && venv/bin/python -c "" 2>/dev/null; then
  PY="venv/bin/python"
else
  PY="python3"
fi
echo "사용 Python: $PY"

# 3) 필수 패키지 확인 후 없으면 설치
if ! "$PY" -c "import openpyxl, pptx" 2>/dev/null; then
  echo "필수 패키지를 설치합니다..."
  "$PY" -m pip install -r requirements.txt || \
    "$PY" -m pip install --break-system-packages -r requirements.txt
fi

# 4) 입력 파일 자동 탐색 (폴더 안 첫 .xlsx / .pptx) — 파일명·한글 정규화 무관
EXCEL=$(ls "certificate excel/"*.xlsx 2>/dev/null | head -n 1)
TEMPLATE=$(ls "certificate ppt/"*.pptx 2>/dev/null | head -n 1)
if [ -z "$EXCEL" ] || [ -z "$TEMPLATE" ]; then
  echo "⚠ 'certificate excel' 폴더의 .xlsx 또는 'certificate ppt' 폴더의 .pptx를 찾지 못했습니다."
  read -n 1 -s -r -p "아무 키나 누르면 닫힙니다..."
  exit 1
fi
echo "엑셀:   $EXCEL"
echo "템플릿: $TEMPLATE"

# 5) 수료증 생성 실행
echo "----------------------------------------"
"$PY" generate_certificates.py --excel "$EXCEL" --template "$TEMPLATE"
STATUS=$?
echo "----------------------------------------"
if [ $STATUS -eq 0 ]; then
  echo "✅ 완료! 결과: 'output ppt' / 'output pdf' 폴더를 확인하세요."
else
  echo "⚠ 오류가 발생했습니다(코드 $STATUS). 위 메시지를 확인하세요."
fi

# 6) 창이 바로 닫히지 않도록 대기
echo ""
read -n 1 -s -r -p "아무 키나 누르면 이 창이 닫힙니다..."
echo ""
