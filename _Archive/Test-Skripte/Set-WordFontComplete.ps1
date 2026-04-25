# Erweiterte Word Standard-Schriftart Registry-Einstellungen
$regWord = "HKCU:\Software\Microsoft\Office\16.0\Word\Options"

Write-Host "=== Erweiterte Word Schriftart-Einstellungen ===" -ForegroundColor Cyan

try {
    if (-not (Test-Path $regWord)) { 
        New-Item -Path $regWord -Force | Out-Null 
    }
    
    Write-Host "Setze umfassende Arial 11pt Einstellungen..." -ForegroundColor Yellow
    
    # Haupteinstellungen
    New-ItemProperty -Path $regWord -Name "DefaultFont" -Value "Arial" -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $regWord -Name "DefaultFontSize" -Value 22 -PropertyType DWord -Force | Out-Null  
    New-ItemProperty -Path $regWord -Name "Font" -Value "Arial" -PropertyType String -Force | Out-Null
    
    # Zusätzliche Schriftart-Schlüssel (falls von Word verwendet)
    New-ItemProperty -Path $regWord -Name "DefaultFontName" -Value "Arial" -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $regWord -Name "NormalFontName" -Value "Arial" -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $regWord -Name "NormalFontSize" -Value 22 -PropertyType DWord -Force | Out-Null
    
    # Font-Fallbacks für verschiedene Sprachen/Regions
    New-ItemProperty -Path $regWord -Name "DefaultFontLatin" -Value "Arial" -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $regWord -Name "FontSubstitutes" -Value "Aptos=Arial;Calibri=Arial" -PropertyType String -Force | Out-Null
    
    Write-Host "✅ Alle Schriftart-Registry-Werte gesetzt:" -ForegroundColor Green
    Write-Host "  DefaultFont = Arial"
    Write-Host "  DefaultFontSize = 22 (11pt)"  
    Write-Host "  Font = Arial"
    Write-Host "  DefaultFontName = Arial"
    Write-Host "  NormalFontName = Arial"
    Write-Host "  NormalFontSize = 22"
    Write-Host "  DefaultFontLatin = Arial"
    Write-Host "  FontSubstitutes = Aptos=Arial;Calibri=Arial"
    
} catch {
    Write-Host "❌ Fehler: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`n=== Umfassende Schriftart-Konfiguration abgeschlossen ===" -ForegroundColor Cyan
Write-Host "Word sollte jetzt definitiv Arial 11pt als Standard verwenden." -ForegroundColor Green