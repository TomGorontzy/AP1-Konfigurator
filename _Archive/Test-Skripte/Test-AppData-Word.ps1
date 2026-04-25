# Word-Formatierung mit korrektem AppData Template-Pfad
Write-Host "Testing Word formatting with AppData template path..." -ForegroundColor Green

try {
    # Standard Template-Pfad ermitteln
    $templatePath = "$env:APPDATA\Microsoft\Templates\Normal.dotm"
    Write-Host "Expected template path: $templatePath"
    Write-Host "Template exists: $(Test-Path $templatePath)"
    
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    
    # Aktuellen Template-Pfad von Word abrufen
    $actualTemplatePath = $word.NormalTemplate.FullName
    Write-Host "Actual template path: $actualTemplatePath"
    
    # Dokument erstellen
    $doc = $word.Documents.Add()
    
    # Standard-Stil konfigurieren
    Write-Host "`nConfiguring Standard style..."
    $standardStyle = $doc.Styles.Item("Standard")
    
    Write-Host "`nCurrent settings:" -ForegroundColor Cyan
    Write-Host "Font: $($standardStyle.Font.Name) $($standardStyle.Font.Size)pt"
    Write-Host "Line Spacing: $($standardStyle.ParagraphFormat.LineSpacing)"
    Write-Host "Space After: $($standardStyle.ParagraphFormat.SpaceAfter)"
    
    # Neue Einstellungen
    $standardStyle.Font.Name = "Arial"
    $standardStyle.Font.Size = 11
    $standardStyle.ParagraphFormat.LineSpacing = 12
    $standardStyle.ParagraphFormat.SpaceBefore = 0
    $standardStyle.ParagraphFormat.SpaceAfter = 0
    
    Write-Host "`nNew settings:" -ForegroundColor Green
    Write-Host "Font: $($standardStyle.Font.Name) $($standardStyle.Font.Size)pt"
    Write-Host "Line Spacing: $($standardStyle.ParagraphFormat.LineSpacing)"
    Write-Host "Space After: $($standardStyle.ParagraphFormat.SpaceAfter)"
    
    # Dokument schließen
    $doc.Close($false)
    
    # Template direkt bearbeiten
    Write-Host "`nOpening Normal.dotm template directly..."
    if (Test-Path $actualTemplatePath) {
        # Template öffnen
        $template = $word.Documents.Open($actualTemplatePath)
        
        # Standard-Stil im Template ändern
        $templateStyle = $template.Styles.Item("Standard")
        $templateStyle.Font.Name = "Arial"
        $templateStyle.Font.Size = 11
        $templateStyle.ParagraphFormat.LineSpacing = 12
        $templateStyle.ParagraphFormat.SpaceBefore = 0
        $templateStyle.ParagraphFormat.SpaceAfter = 0
        
        Write-Host "Modified template Standard style to Arial 11pt"
        
        # Template speichern
        $template.Save()
        Write-Host "Template saved successfully!" -ForegroundColor Green
        
        $template.Close()
    } else {
        Write-Host "Template not found at expected location" -ForegroundColor Yellow
        
        # Alternatives Verzeichnis prüfen
        $altPath = "$env:USERPROFILE\AppData\Roaming\Microsoft\Templates\Normal.dotm"
        Write-Host "Checking alternative path: $altPath"
        if (Test-Path $altPath) {
            Write-Host "Found at alternative location!"
        }
    }
    
    Write-Host "`nSUCCESS: Word formatting updated!" -ForegroundColor Green
    Write-Host "Please restart Word to see changes." -ForegroundColor Yellow
    
} catch {
    Write-Error "Error: $($_.Exception.Message)"
    Write-Host "Full error: $($_.Exception)" -ForegroundColor Red
} finally {
    if ($template) { try { $template.Close() } catch { } }
    if ($doc) { try { $doc.Close($false) } catch { } }
    if ($word) {
        try { 
            $word.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        } catch { }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host "`nTest completed." -ForegroundColor Cyan