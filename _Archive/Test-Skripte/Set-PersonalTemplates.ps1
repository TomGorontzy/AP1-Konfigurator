# Direkter Test zum Setzen des PersonalTemplates Registry-Schlüssels
$regPath = "HKCU:\Software\Microsoft\Office\16.0\Word\Options"
$nueraTemplatesPath = "D:\OneDrive\Desktop\nuera2025_h\Vorlagendateien"

Write-Host "=== PersonalTemplates Registry Update ===" -ForegroundColor Cyan

# Aktuellen Wert anzeigen
try {
    $currentValue = Get-ItemProperty -Path $regPath -Name "PersonalTemplates" -ErrorAction SilentlyContinue
    Write-Host "Aktuell: $($currentValue.PersonalTemplates)" -ForegroundColor Yellow
} catch {
    Write-Host "PersonalTemplates nicht gesetzt" -ForegroundColor Yellow
}

# Neuen Wert setzen
try {
    New-ItemProperty -Path $regPath -Name "PersonalTemplates" -Value $nueraTemplatesPath -PropertyType String -Force | Out-Null
    Write-Host "Gesetzt auf: $nueraTemplatesPath" -ForegroundColor Green
    
    # Verifikation
    $newValue = Get-ItemProperty -Path $regPath -Name "PersonalTemplates" -ErrorAction Stop
    Write-Host "Verifiziert: $($newValue.PersonalTemplates)" -ForegroundColor Green
    
} catch {
    Write-Host "Fehler: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "=== Test abgeschlossen ===" -ForegroundColor Cyan
Write-Host "Bitte starten Sie Word neu, um die Aenderung zu sehen." -ForegroundColor Yellow