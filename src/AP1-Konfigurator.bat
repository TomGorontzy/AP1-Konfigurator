@echo off

REM Nicht-interaktiver Start des AP1-Konfigurators.
REM Proxy-Modus wird über übergebene Argumente gesteuert (z. B. -Proxy On/-Proxy Off).
REM Ohne Argumente wird der Standardmodus aus dem Skript verwendet (Skip).

set SCRIPT=%~dp0AP1-Konfigurator.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -Proxy Skip -Quiet %*