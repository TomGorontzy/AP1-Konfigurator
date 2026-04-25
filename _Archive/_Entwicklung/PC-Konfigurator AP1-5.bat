@echo off
SETLOCAL ENABLEEXTENSIONS

REM Skriptpfad
SET SCRIPT=%~dp0AP1-5.ps1
SET LOGDIR=%~dp0logs
IF NOT EXIST "%LOGDIR%" mkdir "%LOGDIR%"
SET LOG=%LOGDIR%\AP1_Run_%DATE:~-4%%DATE:~3,2%%DATE:~0,2%_%TIME:~0,2%%TIME:~3,2%.log

REM Standardpfade
SET PS32=C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe
SET PS64=C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe

REM Prüfe Office-Bitness
SET OFFICEBIT=64
REG QUERY "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v Platform >nul 2>&1
IF %ERRORLEVEL% EQU 0 (
    FOR /F "tokens=3" %%A IN ('REG QUERY "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v Platform') DO (
        SET OFFICEBIT=%%A
    )
)

REM Wähle PowerShell je nach Office-Bitness
IF /I "%OFFICEBIT%"=="x86" (
    REM Versuche über Sysnative zu starten
    IF EXIST "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" (
        SET PSPATH=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe
    ) ELSE (
        SET PSPATH=%PS32%
    )
    SET BITINFO=32-bit
) ELSE (
    SET PSPATH=%PS64%
    SET BITINFO=64-bit
)

echo Starte AP1-5.ps1 mit %BITINFO% PowerShell...
"%PSPATH%" -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -Proxy Skip > "%LOG%" 2>&1

echo.
echo Skript wurde ausgeführt. Logdatei:
echo %LOG%
pause
ENDLOCAL
