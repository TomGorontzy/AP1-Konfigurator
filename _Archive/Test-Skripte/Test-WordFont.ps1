# Test für Word Standard-Schriftart Registry-Einstellungen
$regPath = "HKCU:\Software\Microsoft\Office\16.0\Word\Options"

Write-Host "=== Word Standard-Schriftart Test ===" -ForegroundColor Cyan

try {
    $props = Get-ItemProperty -Path $regPath -ErrorAction Stop
    
    Write-Host "`nSchriftart Registry-Werte:" -ForegroundColor Yellow
    Write-Host "DefaultFont:     $($props.DefaultFont)"
    Write-Host "DefaultFontSize: $($props.DefaultFontSize) (in halben Punkten)"
    Write-Host "Font:            $($props.Font)"
    
    # Umrechnung Schriftgröße
    if ($props.DefaultFontSize) {
        $actualSize = $props.DefaultFontSize / 2
        Write-Host "Berechnete Größe: $actualSize pt" -ForegroundColor Green
    }
    
    Write-Host "`nErwartete Werte für Arial 11pt:" -ForegroundColor Yellow
    Write-Host "DefaultFont: Arial"
    Write-Host "DefaultFontSize: 22 (11pt * 2)"
    Write-Host "Font: Arial"
    
    # Vergleich
    $isCorrect = ($props.DefaultFont -eq 'Arial') -and ($props.DefaultFontSize -eq 22) -and ($props.Font -eq 'Arial')
    if ($isCorrect) {
        Write-Host "`n✅ Alle Schriftart-Einstellungen korrekt!" -ForegroundColor Green
    } else {
        Write-Host "`n⚠️ Schriftart-Einstellungen unvollständig oder fehlerhaft" -ForegroundColor Yellow
    }
    
} catch {
    Write-Host "❌ Fehler beim Lesen der Registry: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`n=== Test abgeschlossen ===" -ForegroundColor Cyan