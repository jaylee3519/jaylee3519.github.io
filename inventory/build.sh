#!/bin/sh
# apps-script.gs (구글 시트에 붙여넣는 파일) = server.gs + core.js
# server.gs 나 core.js 를 고친 뒤 이 스크립트를 실행하세요.
cd "$(dirname "$0")"
{
  cat server.gs
  printf '\n// ===== 아래는 core.js (앱과 함께 쓰는 재고 처리 로직) - 직접 고치지 말고 core.js를 고친 뒤 build.sh 실행 =====\n\n'
  cat core.js
} > apps-script.gs
echo "apps-script.gs 생성 완료"
