# Umfassende Aptos-zu-Arial Korrektur
Write-Host "=== Comprehensive Aptos to Arial Fix ===" -ForegroundColor Green

# Word vollständig beenden
Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 3

try {
    Write-Host "`n1. Registry Font Substitution..." -ForegroundColor Yellow
    
    # Font Substitution Registry-Einträge
    $fontSubstPath = "HKCU:\Software\Microsoft\Office\16.0\Word\Options"
    $globalFontSubst = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\FontSubstitutes"
    
    # Word-spezifische Font-Substitution
    if (!(Test-Path $fontSubstPath)) {
        New-Item -Path $fontSubstPath -Force | Out-Null
    }
    
    # Aptos explizit durch Arial ersetzen
    Set-ItemProperty -Path $fontSubstPath -Name "FontSubstitutes" -Value "Aptos=Arial" -Force
    Write-Host "Set Word FontSubstitutes: Aptos=Arial"
    
    # Alle möglichen Aptos-Varianten ersetzen
    $aptosVariants = @(
        "Aptos", "Aptos Light", "Aptos Black", "Aptos ExtraLight", 
        "Aptos Medium", "Aptos SemiBold", "Aptos Bold"
    )
    
    foreach ($variant in $aptosVariants) {
        try {
            Set-ItemProperty -Path $fontSubstPath -Name $variant -Value "Arial" -Force
            Write-Host "Substituted: $variant -> Arial"
        } catch {
            Write-Host "Could not substitute: $variant"
        }
    }
    
    Write-Host "`n2. Word COM - Comprehensive Font Fix..." -ForegroundColor Yellow
    
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    
    # Aktuelles Template
    $templatePath = $word.NormalTemplate.FullName
    Write-Host "Template path: $templatePath"
    
    # Template direkt bearbeiten
    $template = $word.Documents.Open($templatePath)
    
    # Alle Styles durchgehen, nicht nur Standard
    Write-Host "`nUpdating all styles in template..."
    $stylesCount = $template.Styles.Count
    Write-Host "Total styles: $stylesCount"
    
    $updatedStyles = 0
    for ($i = 1; $i -le $stylesCount; $i++) {
        try {
            $style = $template.Styles.Item($i)
            $styleName = $style.NameLocal
            
            # Prüfen ob Style eine Schriftart hat
            if ($style.Font.Name -eq "Aptos" -or $style.Font.Name -like "Aptos*") {
                Write-Host "Updating style '$styleName': $($style.Font.Name) -> Arial"
                $style.Font.Name = "Arial"
                $updatedStyles++
            }
            
            # Standard-Style besonders behandeln
            if ($styleName -eq "Standard" -or $styleName -eq "Normal") {
                $style.Font.Name = "Arial"
                $style.Font.Size = 11
                $style.ParagraphFormat.LineSpacing = 12
                $style.ParagraphFormat.SpaceAfter = 0
                $style.ParagraphFormat.SpaceBefore = 0
                Write-Host "Force-updated Standard/Normal style to Arial 11pt"
            }
        } catch {
            # Styles ohne Font-Eigenschaften überspringen
        }
    }
    
    Write-Host "Updated $updatedStyles styles total"
    
    # Template speichern
    Write-Host "`nSaving template..."
    $template.Save()
    $template.Close()
    
    Write-Host "`n3. Testing new document..." -ForegroundColor Yellow
    
    # Neues Dokument erstellen und testen
    $testDoc = $word.Documents.Add()
    $testStyle = $testDoc.Styles.Item("Standard")
    
    Write-Host "New document Standard style:"
    Write-Host "Font Name: '$($testStyle.Font.Name)'"
    Write-Host "Font Size: $($testStyle.Font.Size)"
    Write-Host "Line Spacing: $($testStyle.ParagraphFormat.LineSpacing)"
    
    # Text eingeben und Formatierung prüfen
    $range = $testDoc.Range()
    $range.Text = "Test text with Arial font"
    $range.Font.Name = "Arial"  # Explizit setzen
    
    Write-Host "Test text font: '$($range.Font.Name)'"
    
    # Ohne Speichern schließen
    $testDoc.Close($false)
    
    Write-Host "`n4. Additional Registry Settings..." -ForegroundColor Yellow
    
    # Weitere Registry-Einstellungen für Office
    $officeCommonPath = "HKCU:\Software\Microsoft\Office\16.0\Common\LanguageResources"
    if (!(Test-Path $officeCommonPath)) {
        New-Item -Path $officeCommonPath -Force | Out-Null
    }
    
    # Default UI Font überschreiben
    Set-ItemProperty -Path $fontSubstPath -Name "DefaultUIFont" -Value "Arial" -Force
    Set-ItemProperty -Path $fontSubstPath -Name "BodyFont" -Value "Arial" -Force
    Set-ItemProperty -Path $fontSubstPath -Name "HeadingFont" -Value "Arial" -Force
    
    Write-Host "Set additional font overrides"
    
    Write-Host "`nSUCCESS: Comprehensive font fix applied!" -ForegroundColor Green
    
} catch {
    Write-Error "Error: $($_.Exception.Message)"
} finally {
    if ($testDoc) { try { $testDoc.Close($false) } catch { } }
    if ($template) { try { $template.Close() } catch { } }
    if ($word) {
        try {
            $word.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        } catch { }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host "`n=== FINAL INSTRUCTIONS ===" -ForegroundColor Cyan
Write-Host "1. Close ALL Word windows completely" -ForegroundColor Yellow
Write-Host "2. Wait 10 seconds" -ForegroundColor Yellow  
Write-Host "3. Open Word fresh" -ForegroundColor Yellow
Write-Host "4. Create new blank document" -ForegroundColor Yellow
Write-Host "5. Check font name in Home ribbon" -ForegroundColor Yellow
Write-Host "6. If still Aptos, manually select Arial from font dropdown" -ForegroundColor Yellow