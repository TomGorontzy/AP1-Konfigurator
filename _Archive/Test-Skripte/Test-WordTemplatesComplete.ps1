# Umfassender Test für Word Template Registry-Einstellungen
$regPath = "HKCU:\Software\Microsoft\Office\16.0\Word\Options"

Write-Host "=== Word Template Registry Test ===" -ForegroundColor Cyan

try {
    $props = Get-ItemProperty -Path $regPath -ErrorAction Stop
    
    Write-Host "`nRegistry-Werte:" -ForegroundColor Yellow
    Write-Host "PersonalTemplates:  $($props.'PersonalTemplates')"
    Write-Host "USER-DOT-PATH:      $($props.'USER-DOT-PATH')"
    Write-Host "DOT-PATH:           $($props.'DOT-PATH')"  
    Write-Host "WORKGROUP-DOT-PATH: $($props.'WORKGROUP-DOT-PATH')"
    
    # Prüfe Pfad-Existenz
    $paths = @($props.'PersonalTemplates', $props.'USER-DOT-PATH', $props.'DOT-PATH', $props.'WORKGROUP-DOT-PATH') | Where-Object { $_ }
    
    Write-Host "`nPfad-Validierung:" -ForegroundColor Yellow
    foreach ($path in ($paths | Select-Object -Unique)) {
        if ($path -and (Test-Path $path)) {
            $fileCount = (Get-ChildItem $path -File -ErrorAction SilentlyContinue | Measure-Object).Count
            Write-Host "✅ $path ($fileCount Dateien)" -ForegroundColor Green
        } elseif ($path) {
            Write-Host "❌ $path (nicht gefunden)" -ForegroundColor Red
        }
    }
    
    # Word-Priorität erklären
    Write-Host "`nWord Template-Priorität:" -ForegroundColor Yellow
    Write-Host "1. PersonalTemplates (moderner Office-Schlüssel - höchste Priorität)"
    Write-Host "2. WORKGROUP-DOT-PATH (Firmenvorlagen)"
    Write-Host "3. USER-DOT-PATH (Legacy-Schlüssel für persönliche Vorlagen)"  
    Write-Host "4. Standard-Pfad (niedrigste Priorität)"
    Write-Host "`nWord verwendet den ersten verfügbaren Pfad in dieser Reihenfolge."
    
} catch {
    Write-Host "❌ Fehler beim Lesen der Registry: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`n=== Test abgeschlossen ===" -ForegroundColor Cyan