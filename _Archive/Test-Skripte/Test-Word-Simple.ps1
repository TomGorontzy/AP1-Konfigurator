# Test Script for Word Formatting via COM
Write-Host "Testing Word formatting via COM..." -ForegroundColor Green

try {
    Write-Host "Starting Word COM object..."
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    
    Write-Host "Accessing Normal template..."
    $normalTemplate = $word.NormalTemplate
    $normalStyle = $normalTemplate.Styles.Item("Normal")
    
    Write-Host "`nCurrent settings:" -ForegroundColor Cyan
    Write-Host "Font: $($normalStyle.Font.Name)"
    Write-Host "Size: $($normalStyle.Font.Size)"
    Write-Host "Line spacing: $($normalStyle.ParagraphFormat.LineSpacing)"
    Write-Host "Space before: $($normalStyle.ParagraphFormat.SpaceBefore)"
    Write-Host "Space after: $($normalStyle.ParagraphFormat.SpaceAfter)"
    
    Write-Host "`nSetting new values..." -ForegroundColor Yellow
    $normalStyle.Font.Name = "Arial"
    $normalStyle.Font.Size = 11
    $normalStyle.ParagraphFormat.LineSpacing = 12
    $normalStyle.ParagraphFormat.SpaceBefore = 0
    $normalStyle.ParagraphFormat.SpaceAfter = 0
    
    Write-Host "Saving Normal.dotm..."
    $normalTemplate.Save()
    
    Write-Host "`nNew settings:" -ForegroundColor Green
    Write-Host "Font: $($normalStyle.Font.Name)"
    Write-Host "Size: $($normalStyle.Font.Size)"
    Write-Host "Line spacing: $($normalStyle.ParagraphFormat.LineSpacing)"
    Write-Host "Space before: $($normalStyle.ParagraphFormat.SpaceBefore)"
    Write-Host "Space after: $($normalStyle.ParagraphFormat.SpaceAfter)"
    
    Write-Host "`nSUCCESS: Normal.dotm updated!" -ForegroundColor Green
    Write-Host "Restart Word to see changes." -ForegroundColor Yellow
    
} catch {
    Write-Error "Error configuring Word: $($_.Exception.Message)"
} finally {
    if ($word) {
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host "`nTest completed." -ForegroundColor Cyan