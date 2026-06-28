@echo off
REM Proxy deaktivieren (Windows)

REM Setzt die Proxy-Einstellungen für den aktuellen Benutzer zurück
reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings" /v ProxyServer /f
reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings" /v ProxyEnable /f

REM Setzt Umgebungsvariablen zurück (nur für aktuelle Sitzung)
setx http_proxy "" >nul
setx https_proxy "" >nul
setx HTTP_PROXY "" >nul
setx HTTPS_PROXY "" >nul

REM Hinweis für den Benutzer
ECHO Proxy wurde deaktiviert.
PAUSE