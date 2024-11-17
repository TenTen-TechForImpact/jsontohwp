@echo off
:: 관리자 권한으로 실행 확인
setlocal

:: 관리자 권한으로 실행
echo Running script with Administrator privileges...
runas /user:Administrator "python install_modules.py"

endlocal
