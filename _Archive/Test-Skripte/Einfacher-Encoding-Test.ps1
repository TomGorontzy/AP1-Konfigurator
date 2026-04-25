# Einfacher Encoding-Test
Write-Host "Terminal-Test für Umlaute:" -ForegroundColor Cyan
Write-Host "Originaltext: Prüfungsrechner für Abschlussprüfung" 
Write-Host "ASCII-Version: Pruefungsrechner fuer Abschlusspruefung"
Write-Host ""
Write-Host "PowerShell Version: $($PSVersionTable.PSVersion.ToString())"
$cp = (chcp.com) -replace '\D'
Write-Host "Codepage: $cp"
Write-Host ""
Write-Host "Fazit: Umlaute-Probleme sind normal in Windows PowerShell 5.1" -ForegroundColor Yellow