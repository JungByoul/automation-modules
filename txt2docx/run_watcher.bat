@echo off
chcp 65001 > nul
title txt→docx 자동 변환 감시

:: ANTHROPIC_API_KEY는 Windows 사용자 환경변수로 등록되어 있음
python "%~dp0doc_watcher.py"
pause
