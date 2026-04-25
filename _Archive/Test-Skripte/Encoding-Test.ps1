# Encoding-Test für verschiedene Umlaute-Ausgabemethoden
# =========================================================

Write-Host "=== ENCODING-TEST für UMLAUTE ===" -ForegroundColor Cyan

# 1. Terminal-Info
Write-Host "`n1. Terminal-Information:" -ForegroundColor Yellow
Write-Host "   PowerShell Version: $($PSVersionTable.PSVersion)"
Write-Host "   Aktuelle Codepage: $(chcp.com | Out-String)" -NoNewline
Write-Host "   Output Encoding: $([Console]::OutputEncoding.EncodingName)"

# 2. Test verschiedene Ausgabemethoden
$testText = "Prüfungsrechner für die Abschlussprüfung - Größe: 10€"

Write-Host "`n2. Verschiedene Ausgabemethoden:" -ForegroundColor Yellow
Write-Host "   Write-Host:   $testText"
Write-Output "   Write-Output: $testText"
"   String:       $testText" | Out-Host

# 3. Encoding-Fixes testen
Write-Host "`n3. Mit Encoding-Fixes:" -ForegroundColor Yellow
try {
    chcp 65001 | Out-Null
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    Write-Host "   Nach UTF-8:   $testText"
} catch {
    Write-Host "   Encoding-Fix fehlgeschlagen: $_" -ForegroundColor Red
}

# 4. ASCII-Fallback
Write-Host "`n4. ASCII-Fallback:" -ForegroundColor Yellow
$asciiText = $testText -replace 'ü', 'ue' -replace 'Ü', 'Ue' -replace 'ö', 'oe' -replace 'Ö', 'Oe' -replace 'ä', 'ae' -replace 'Ä', 'Ae' -replace 'ß', 'ss' -replace '€', 'EUR'
Write-Host "   ASCII-Text:   $asciiText"

# 5. Empfehlungen
Write-Host "`n5. Empfehlungen:" -ForegroundColor Green
Write-Host "   - Nutzen Sie Windows Terminal statt PowerShell ISE"
Write-Host "   - Installieren Sie PowerShell 7.x für bessere UTF-8-Unterstützung"
Write-Host "   - Setzen Sie Konsolen-Schriftart auf 'Cascadia Code' oder 'Consolas'"
Write-Host "   - Verwenden Sie chcp 65001 vor Skriptausführung"

Write-Host "`n=== TEST ABGESCHLOSSEN ===" -ForegroundColor Cyan