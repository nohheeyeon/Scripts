@echo off
:: BatchGotAdmin
:-------------------------------------
REM  --> Check for permissions
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"

REM --> If error flag set, we do not have admin.
if '%errorlevel%' NEQ '0' (
    echo Requesting administrative privileges...
    goto UACPrompt
) else ( goto gotAdmin )

:UACPrompt
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    set params = %*:"=""
    echo UAC.ShellExecute "cmd.exe", "/c %~s0 %params%", "", "runas", 1 >> "%temp%\getadmin.vbs"

    "%temp%\getadmin.vbs"
    del "%temp%\getadmin.vbs"
    exit /B

:gotAdmin
    pushd "%CD%"
    CD /D "%~dp0"
:--------------------------------------

Rem Edge ������Ʈ ���� ���� ����
reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\edgeupdate" /v Start /t REG_DWORD /d 4 /f
reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\edgeupdatem" /v Start /t REG_DWORD /d 4 /f
reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\MicrosoftEdgeElevationService" /v Start /t REG_DWORD /d 4 /f
sc stop edgeupdate
sc stop edgeupdatem
sc stop MicrosoftEdgeElevationService
taskkill /f /im MicrosoftEdgeUpdate.exe
Rem ȣ��Ʈ ���� ���
copy "%systemroot%\system32\drivers\etc\hosts" "%systemroot%\system32\drivers\etc\hosts_bak"
Rem Edge ������Ʈ ���� ȣ��Ʈ ���� �߰�
echo.>> "%systemroot%\system32\drivers\etc\hosts"
echo 127.0.0.1 msedge.api.cdp.microsoft.com>> "%systemroot%\system32\drivers\etc\hosts"
echo 127.0.0.1 www.msedge.api.cdp.microsoft.com>> "%systemroot%\system32\drivers\etc\hosts"
echo 127.0.0.1 *.dl.delivery.mp.microsoft.com>> "%systemroot%\system32\drivers\etc\hosts"
echo 127.0.0.1 www.*.dl.delivery.mp.microsoft.com>> "%systemroot%\system32\drivers\etc\hosts"

@echo === ���� ������ �ڵ� ������Ʈ ���� �Ϸ� ===
@pause