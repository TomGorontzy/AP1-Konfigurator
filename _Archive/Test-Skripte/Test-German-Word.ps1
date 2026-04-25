# Korrekter Test für deutsche Word-Formatierung
Write-Host "Configuring German Word Standard style..." -ForegroundColor Green

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $doc = $word.Documents.Add()
    
    # Zugriff auf "Standard" Stil (deutsche Version)
    Write-Host "Accessing 'Standard' style..."
    $standardStyle = $doc.Styles.Item("Standard")
    
    Write-Host "`nCurrent Standard style settings:" -ForegroundColor Cyan
    Write-Host "Font Name: $($standardStyle.Font.Name)"
    Write-Host "Font Size: $($standardStyle.Font.Size)"
    Write-Host "Line Spacing: $($standardStyle.ParagraphFormat.LineSpacing)"
    Write-Host "Space Before: $($standardStyle.ParagraphFormat.SpaceBefore)"
    Write-Host "Space After: $($standardStyle.ParagraphFormat.SpaceAfter)"
    
    # Neue Formatierung setzen
    Write-Host "`nSetting new formatting..." -ForegroundColor Yellow
    $standardStyle.Font.Name = "Arial"
    $standardStyle.Font.Size = 11
    $standardStyle.ParagraphFormat.LineSpacing = 12  # Einfacher Zeilenabstand
    $standardStyle.ParagraphFormat.SpaceBefore = 0
    $standardStyle.ParagraphFormat.SpaceAfter = 0
    
    # Template speichern
    Write-Host "Saving changes to Normal template..."
    $word.NormalTemplate.Save()
    
    Write-Host "`nNew Standard style settings:" -ForegroundColor Green
    Write-Host "Font Name: $($standardStyle.Font.Name)"
    Write-Host "Font Size: $($standardStyle.Font.Size)"
    Write-Host "Line Spacing: $($standardStyle.ParagraphFormat.LineSpacing)"
    Write-Host "Space Before: $($standardStyle.ParagraphFormat.SpaceBefore)"
    Write-Host "Space After: $($standardStyle.ParagraphFormat.SpaceAfter)"
    
    $doc.Close($false)
    
    Write-Host "`nSUCCESS: Standard style updated to Arial 11pt!" -ForegroundColor Green
    Write-Host "Restart Word to see changes." -ForegroundColor Yellow
    
} catch {
    Write-Error "Error: $($_.Exception.Message)"
} finally {
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