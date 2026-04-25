# Word COM Test with proper error handling
Write-Host "Testing Word COM access..." -ForegroundColor Green

try {
    # Start Word
    Write-Host "Creating Word COM object..."
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    
    # Create a new document to initialize templates
    Write-Host "Creating document to initialize templates..."
    $doc = $word.Documents.Add()
    
    # Access Normal style through document
    Write-Host "Accessing Normal style..."
    $normalStyle = $doc.Styles.Item("Normal")
    
    Write-Host "`nCurrent Normal style settings:" -ForegroundColor Cyan
    Write-Host "Font Name: $($normalStyle.Font.Name)"
    Write-Host "Font Size: $($normalStyle.Font.Size)"
    Write-Host "Line Spacing: $($normalStyle.ParagraphFormat.LineSpacing)"
    Write-Host "Space Before: $($normalStyle.ParagraphFormat.SpaceBefore)"
    Write-Host "Space After: $($normalStyle.ParagraphFormat.SpaceAfter)"
    
    # Modify Normal style
    Write-Host "`nModifying Normal style..." -ForegroundColor Yellow
    $normalStyle.Font.Name = "Arial"
    $normalStyle.Font.Size = 11
    $normalStyle.ParagraphFormat.LineSpacing = 12
    $normalStyle.ParagraphFormat.SpaceBefore = 0
    $normalStyle.ParagraphFormat.SpaceAfter = 0
    
    # Save to template
    Write-Host "Saving changes to Normal template..."
    $word.NormalTemplate.Save()
    
    Write-Host "`nNew Normal style settings:" -ForegroundColor Green
    Write-Host "Font Name: $($normalStyle.Font.Name)"
    Write-Host "Font Size: $($normalStyle.Font.Size)"
    Write-Host "Line Spacing: $($normalStyle.ParagraphFormat.LineSpacing)"
    Write-Host "Space Before: $($normalStyle.ParagraphFormat.SpaceBefore)"
    Write-Host "Space After: $($normalStyle.ParagraphFormat.SpaceAfter)"
    
    # Close document without saving
    $doc.Close($false)
    
    Write-Host "`nSUCCESS: Normal style updated in template!" -ForegroundColor Green
    
} catch {
    Write-Error "Error: $($_.Exception.Message)"
    Write-Host "Full error details: $($_.Exception)" -ForegroundColor Red
} finally {
    if ($doc) { 
        try { $doc.Close($false) } catch { }
    }
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