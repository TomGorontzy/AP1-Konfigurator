@echo off

REM Execution Policy für diesen Aufruf auf Bypass setzen
REM Proxy-Parameter anpassen wie benötigt

set SCRIPT=%~dp0AP1-Konfigurator.ps1

set /p PROXYMODE=Proxy aktivieren? (J/N): 
if /I "%PROXYMODE%"=="J" (
	set PROXYPARAM=-Proxy On -ProxyServer "192.168.0.1:8080" -ProxyBypass "*.office365.com; *.cloudappsecurity.com; *.onmicrosoft.com; *.office.net; *.office.com; *.microsoft.com; *.microsoftonline.com; *.live.com; *.azure.net; *.gfx.ms; *.onestore.ms; *.msecnd.net; *.outlookgroups.ms; *.linkedin.com; *.msocdn.com; *.live.net; ihk-aka.de"
) else (
	set PROXYPARAM=-Proxy Off
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" %PROXYPARAM% %*
pause
REM exit /b