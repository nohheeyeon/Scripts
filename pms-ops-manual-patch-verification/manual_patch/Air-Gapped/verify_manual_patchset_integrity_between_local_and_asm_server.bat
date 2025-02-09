@echo off
cd /d %~dp0
python verify_manual_patchset_integrity_between_local_and_asm_server.py
pause
