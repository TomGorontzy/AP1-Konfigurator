# Direkter Test zum Setzen der Word Standard-Schriftart Registry
$regPath = "HKCU:\Software\Microsoft\Office\16.0\Word\Options"

Write-Host "=== Word Standard-Schriftart Registry Update ===" -ForegroundColor Cyan

# Setze die Werte direkt
try {
    if (-not (Test-Path $regPath)) { 
        New-Item -Path $regPath -Force | Out-Null 
        Write-Host "Registry-Pfad erstellt: $regPath" -ForegroundColor Yellow
    }
    
    # Arial 11pt Einstellungen
    New-ItemProperty -Path $regPath -Name "DefaultFont" -Value "Arial" -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $regPath -Name "DefaultFontSize" -Value 22 -PropertyType DWord -Force | Out-Null  
    New-ItemProperty -Path $regPath -Name "Font" -Value "Arial" -PropertyType String -Force | Out-Null
    
    Write-Host "✅ Registry-Werte gesetzt:" -ForegroundColor Green
    Write-Host "  DefaultFont = Arial"
    Write-Host "  DefaultFontSize = 22 (11pt)"
    Write-Host "  Font = Arial"
    
    # Verifikation
    $props = Get-ItemProperty -Path $regPath -ErrorAction Stop
    $actualSize = $props.DefaultFontSize / 2
    Write-Host "`n✅ Verifiziert: $($props.DefaultFont) $actualSize pt" -ForegroundColor Green
    
} catch {
    Write-Host "❌ Fehler: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`n=== Update abgeschlossen ===" -ForegroundColor Cyan
Write-Host "Hinweis: Word muss neu gestartet werden fuer Aenderungen." -ForegroundColor Yellow