@echo off
chcp 65001 >nul
set PYTHONUTF8=1
echo ============================================
echo   Outlook mail -^> MOC ledger auto-collect
echo ============================================
echo.
python "%~dp0collect_moc_mail.py"
echo.
echo --------------------------------------------
echo  Done. Press any key to close.
pause >nul
