@echo off
chcp 65001 > nul
echo 栄養食事指導支援システム (3日分対応) を起動しています...
start http://localhost:8000
python -m http.server 8000
pause
