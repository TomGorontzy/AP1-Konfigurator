# Test-Skript für Word USER-DOT-PATH Registry-Einstellung (persönliche Vorlagen)
$regPath = "HKCU:\Software\Microsoft\Office\16.0\Word\Options"
$nueraPath = "D:\OneDrive\Desktop\nuera2025_h"

Write-Host "Test: Word USER-DOT-PATH Registry-Einstellung (persönliche Vorlagen)" -ForegroundColor Cyan

# Prüfe aktueller Wert für persönliche Vorlagen
try {
    $currentValue = Get-ItemProperty -Path $regPath -Name "USER-DOT-PATH" -ErrorAction SilentlyContinue
    if ($currentValue) {
        Write-Host "Aktueller USER-DOT-PATH: $($currentValue.'USER-DOT-PATH')" -ForegroundColor Green
    } else {
        Write-Host "USER-DOT-PATH nicht gesetzt" -ForegroundColor Yellow
    }
} catch {
    Write-Host "Registry-Pfad nicht gefunden oder nicht lesbar" -ForegroundColor Red
}

# Prüfe, ob Nuera-Ordner existiert
if (Test-Path $nueraPath) {
    Write-Host "Nuera-Ordner gefunden: $nueraPath" -ForegroundColor Green
    $files = Get-ChildItem $nueraPath -Filter "*.dot*" | Select-Object -First 3
    if ($files) {
        Write-Host "Vorlagen-Dateien gefunden:"
        $files | ForEach-Object { Write-Host "  - $($_.Name)" }
    } else {
        Write-Host "Keine .dot/.dotx Dateien im Nuera-Ordner gefunden" -ForegroundColor Yellow
    }
} else {
    Write-Host "Nuera-Ordner nicht gefunden: $nueraPath" -ForegroundColor Red
}

Write-Host "Test abgeschlossen" -ForegroundColor Cyan