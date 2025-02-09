@echo off

REM 가상 환경 활성화
call %~dp0\.venv\Scripts\activate.bat

REM 스크립트 실행
python %~dp0\standalone_patch_auto_install.py AMD64

REM 가상 환경 비활성화
deactivate

pause
