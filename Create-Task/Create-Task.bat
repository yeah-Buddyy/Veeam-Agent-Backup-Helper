@echo off
setlocal EnableDelayedExpansion

:: Run as Admin
FSUTIL DIRTY query %SYSTEMDRIVE% >nul || (
    PowerShell.exe "Start-Process -FilePath %COMSPEC% -Args '/C CHDIR /D %CD% & "%0"' -Verb RunAs"
    EXIT
)

IF EXIST "%~dp0Veeam-Agent-Backup-Helper.ps1" (
  echo Veeam-Agent-Backup-Helper.ps1 found in current path
  set "parent=%~dp0"
) ELSE (
  REM Get parent directory path
  echo Veeam-Agent-Backup-Helper.ps1 not found, trying to find it in parent path
  for %%? in ("%~dp0..") do (
    set "parent=%%~f?\"
  )
)

IF EXIST "!parent!Veeam-Agent-Backup-Helper.ps1" (
  echo Found Veeam-Agent-Backup-Helper.ps1
) ELSE (
  echo Could not find Veeam-Agent-Backup-Helper.ps1
  pause
  exit
)

%windir%\System32\schtasks.exe /create /tn "VeeamBackupExternDrive" /ru %USERNAME% /RL HIGHEST /Sc ONLOGON /tr "'%WINDIR%\System32\WindowsPowerShell\v1.0\powershell.exe' -noprofile -nologo -windowstyle hidden -ExecutionPolicy Bypass \"!parent!Veeam-Agent-Backup-Helper.ps1\"

endlocal

:: English Systems
FOR /F %%I IN ('%windir%\System32\schtasks.exe /QUERY /FO LIST /TN "VeeamBackupExternDrive" ^| FIND /C "Running"') DO (
    IF %%I == 0 (SET STATUS=Running) Else (SET Status=Ready)
    ECHO %%I
)
ECHO %STATUS%

if "%STATUS%" == "Ready" (
    "%windir%\System32\schtasks.exe" /run /tn "VeeamBackupExternDrive"
)

:: German Systems
FOR /F %%I IN ('%windir%\System32\schtasks.exe /QUERY /FO LIST /TN "VeeamBackupExternDrive" ^| FIND /C "Bereit"') DO (
    IF %%I == 0 (SET STATUS=Running) Else (SET Status=Ready)
    ECHO %%I
)
ECHO %STATUS%

if "%STATUS%" == "Ready" (
    "%windir%\System32\schtasks.exe" /run /tn "VeeamBackupExternDrive"
)

pause
