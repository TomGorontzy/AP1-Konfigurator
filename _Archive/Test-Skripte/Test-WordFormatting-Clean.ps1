# Test-Skript für Word-Formatierung
# Direkte Konfiguration der Normal.dotm via COM

Write-Host "Teste Word-Formatierung über COM..." -ForegroundColor Green

try {
    # Word-Anwendung erstellen
    Write-Host "Starte Word COM-Objekt..."
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    
    # Normal-Template öffnen
    Write-Host "Zugriff auf Normal-Template..."
    $normalTemplate = $word.NormalTemplate
    $normalStyle = $normalTemplate.Styles.Item("Normal")
    
    # Aktuelle Einstellungen anzeigen
    Write-Host "`nAktuelle Einstellungen:" -ForegroundColor Cyan
    Write-Host "Schriftart: $($normalStyle.Font.Name)"
    Write-Host "Schriftgröße: $($normalStyle.Font.Size)"
    Write-Host "Zeilenabstand: $($normalStyle.ParagraphFormat.LineSpacing)"
    Write-Host "Absatzabstand vor: $($normalStyle.ParagraphFormat.SpaceBefore)"
    Write-Host "Absatzabstand nach: $($normalStyle.ParagraphFormat.SpaceAfter)"
    
    # Neue Einstellungen setzen
    Write-Host "`nSetze neue Einstellungen..." -ForegroundColor Yellow
    $normalStyle.Font.Name = "Arial"
    $normalStyle.Font.Size = 11
    $normalStyle.ParagraphFormat.LineSpacing = 12  # Einfacher Zeilenabstand
    $normalStyle.ParagraphFormat.SpaceBefore = 0
    $normalStyle.ParagraphFormat.SpaceAfter = 0
    
    # Template speichern
    Write-Host "Speichere Normal.dotm..."
    $normalTemplate.Save()
    
    # Neue Einstellungen bestätigen
    Write-Host "`nNeue Einstellungen:" -ForegroundColor Green
    Write-Host "Schriftart: $($normalStyle.Font.Name)"
    Write-Host "Schriftgröße: $($normalStyle.Font.Size)"
    Write-Host "Zeilenabstand: $($normalStyle.ParagraphFormat.LineSpacing)"
    Write-Host "Absatzabstand vor: $($normalStyle.ParagraphFormat.SpaceBefore)"
    Write-Host "Absatzabstand nach: $($normalStyle.ParagraphFormat.SpaceAfter)"
    
    Write-Host "`nERFOLGREICH: Normal.dotm wurde aktualisiert!" -ForegroundColor Green
    Write-Host "Starten Sie Word neu, um die Änderungen zu sehen." -ForegroundColor Yellow
    
} catch {
    Write-Error "Fehler beim Konfigurieren von Word: $($_.Exception.Message)"
} finally {
    # COM-Objekt freigeben
    if ($word) {
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host "`nTest abgeschlossen." -ForegroundColor Cyan