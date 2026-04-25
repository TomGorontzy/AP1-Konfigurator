# Direkter Test für Word-Formatierung (Normal.dotm über COM)
Write-Host "=== Word Normal.dotm Formatierungs-Test ===" -ForegroundColor Cyan

try {
    # Word starten und Normal.dotm konfigurieren
    Write-Host "Starte Word und konfiguriere Normal.dotm..." -ForegroundColor Yellow
    
    # Alle Word-Prozesse beenden
    Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    
    $word = New-Object -ComObject Word.Application -ErrorAction Stop
    $word.Visible = $false
    $word.DisplayAlerts = 0
    
    Write-Host "Word gestartet, konfiguriere Normal-Style..." -ForegroundColor Yellow
    
    # Temporäres Dokument für Normal-Style Zugriff
    $doc = $word.Documents.Add()
    $normalStyle = $doc.Styles.Item('Normal')
    
    # Aktuelle Werte anzeigen
    Write-Host "`nVorher:" -ForegroundColor Red
    Write-Host "  Schriftart: $($normalStyle.Font.Name)"
    Write-Host "  Schriftgröße: $($normalStyle.Font.Size)"
    Write-Host "  Zeilenabstand: $($normalStyle.ParagraphFormat.LineSpacing)"
    Write-Host "  Abstand nach: $($normalStyle.ParagraphFormat.SpaceAfter)"
    
    # Neue Werte setzen
    $normalStyle.Font.Name = 'Arial'
    $normalStyle.Font.Size = 11
    $normalStyle.ParagraphFormat.LineSpacing = 12  # Einfacher Zeilenabstand
    $normalStyle.ParagraphFormat.SpaceAfter = 0   # Kein Abstand nach Absatz
    $normalStyle.ParagraphFormat.SpaceBefore = 0  # Kein Abstand vor Absatz
    
    # Normal.dotm speichern
    $normalTemplate = $word.NormalTemplate
    $normalTemplate.Save()
    
    Write-Host "`nNachher:" -ForegroundColor Green
    Write-Host "  Schriftart: $($normalStyle.Font.Name)"
    Write-Host "  Schriftgröße: $($normalStyle.Font.Size)"
    Write-Host "  Zeilenabstand: $($normalStyle.ParagraphFormat.LineSpacing)"
    Write-Host "  Abstand nach: $($normalStyle.ParagraphFormat.SpaceAfter)"
    
    # Dokument ohne Speichern schließen
    $doc.Close([ref]$false)
    
    Write-Host "`n✅ Normal.dotm erfolgreich konfiguriert!" -ForegroundColor Green
    
} catch {
    Write-Host "❌ Fehler: $($_.Exception.Message)" -ForegroundColor Red
} finally {
    if ($word) {
        try { $word.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host "`n=== Test abgeschlossen ===" -ForegroundColor Cyan
Write-Host "Starten Sie Word neu, um die Änderungen zu sehen." -ForegroundColor Yellow