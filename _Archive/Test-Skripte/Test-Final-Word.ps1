# Finale Word-Formatierung mit korrektem Template-Pfad
Write-Host "Final Word formatting test..." -ForegroundColor Green

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $doc = $word.Documents.Add()
    
    # Standard-Stil konfigurieren
    Write-Host "Configuring Standard style..."
    $standardStyle = $doc.Styles.Item("Standard")
    
    Write-Host "`nBefore changes:" -ForegroundColor Cyan
    Write-Host "Font: $($standardStyle.Font.Name) $($standardStyle.Font.Size)pt"
    Write-Host "Line Spacing: $($standardStyle.ParagraphFormat.LineSpacing)"
    Write-Host "Space After: $($standardStyle.ParagraphFormat.SpaceAfter)"
    
    # Änderungen vornehmen
    $standardStyle.Font.Name = "Arial"
    $standardStyle.Font.Size = 11
    $standardStyle.ParagraphFormat.LineSpacing = 12
    $standardStyle.ParagraphFormat.SpaceBefore = 0
    $standardStyle.ParagraphFormat.SpaceAfter = 0
    
    Write-Host "`nAfter changes:" -ForegroundColor Green
    Write-Host "Font: $($standardStyle.Font.Name) $($standardStyle.Font.Size)pt"
    Write-Host "Line Spacing: $($standardStyle.ParagraphFormat.LineSpacing)"
    Write-Host "Space After: $($standardStyle.ParagraphFormat.SpaceAfter)"
    
    # Template-Pfad anzeigen und versuchen zu speichern
    $templatePath = $word.NormalTemplate.FullName
    Write-Host "`nNormal template path: $templatePath"
    
    try {
        # Erst das Dokument mit den neuen Einstellungen schließen
        $doc.Close($false)
        
        # Template direkt öffnen, ändern und speichern
        Write-Host "Opening Normal.dotm directly..."
        $template = $word.Documents.Open($templatePath)
        
        # Standard-Stil im Template ändern
        $templateStandardStyle = $template.Styles.Item("Standard")
        $templateStandardStyle.Font.Name = "Arial"
        $templateStandardStyle.Font.Size = 11
        $templateStandardStyle.ParagraphFormat.LineSpacing = 12
        $templateStandardStyle.ParagraphFormat.SpaceBefore = 0
        $templateStandardStyle.ParagraphFormat.SpaceAfter = 0
        
        # Template speichern
        $template.Save()
        $template.Close()
        
        Write-Host "SUCCESS: Normal.dotm updated directly!" -ForegroundColor Green
        
    } catch {
        Write-Host "Direct template save failed: $($_.Exception.Message)" -ForegroundColor Yellow
        
        # Alternative: Neues Dokument erstellen und als Template speichern
        Write-Host "Trying alternative method..."
        $newDoc = $word.Documents.Add()
        $newStandardStyle = $newDoc.Styles.Item("Standard")
        $newStandardStyle.Font.Name = "Arial"
        $newStandardStyle.Font.Size = 11
        $newStandardStyle.ParagraphFormat.LineSpacing = 12
        $newStandardStyle.ParagraphFormat.SpaceBefore = 0
        $newStandardStyle.ParagraphFormat.SpaceAfter = 0
        
        # Als Template speichern
        $templateDir = Split-Path $templatePath -Parent
        $newTemplatePath = Join-Path $templateDir "Normal_New.dotm"
        $newDoc.SaveAs2($newTemplatePath, 15) # 15 = wdFormatXMLTemplate
        $newDoc.Close()
        
        Write-Host "Saved as: $newTemplatePath" -ForegroundColor Green
    }
    
} catch {
    Write-Error "Error: $($_.Exception.Message)"
} finally {
    if ($template) { try { $template.Close() } catch { } }
    if ($newDoc) { try { $newDoc.Close($false) } catch { } }
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