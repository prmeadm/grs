#!/bin/bash
# Проверки перед выкаткой: синтаксис фронтенда и бэкенда, отсутствие секретов в репозитории.
set -e
cd "$(git rev-parse --show-toplevel)"
TMP=$(mktemp -d); trap 'rm -rf "$TMP"' EXIT

cp apps-script/Code.gs "$TMP/code.js"
node --check "$TMP/code.js" && echo "✓ Code.gs: синтаксис"

python3 - "$TMP/front.js" <<'P'
import re, sys
h = open('index.html').read()
open(sys.argv[1], 'w').write('\n'.join(re.findall(r'<script>(.*?)</script>', h, re.S)))
P
node --check "$TMP/front.js" && echo "✓ index.html: синтаксис скрипта"

grep -q "const API_URL = 'https://script.google.com/macros/s/AKfycb" index.html \
  && echo "✓ API_URL на месте" || { echo "✗ API_URL пустой или изменён"; exit 1; }

# Токен бота (123456:ABC...) и ID супергруппы (-100...) не должны попасть в публичный репозиторий
if git ls-files -z | xargs -0 grep -nE "[0-9]{8,10}:[A-Za-z0-9_-]{30,}|-100[0-9]{6,}" 2>/dev/null; then
  echo "✗ Похоже на секрет в файлах выше — убери перед коммитом"; exit 1
fi
echo "✓ Секретов в репозитории нет"

echo "APP_VERSION в коде: $(grep -oE 'const APP_VERSION = [0-9]+' apps-script/Code.gs | grep -oE '[0-9]+$')"
