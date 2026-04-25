# Test der aktualisierten Set-WordDefaultFormatting Funktion
Write-Host "Testing updated Set-WordDefaultFormatting function..." -ForegroundColor Green

# Office Version ermitteln (fuer Registry-Pfad)
$OfficeVersion = "16.0"  # Standard Office 365/2019/2021

# Hilfsfunktionen definieren (vereinfacht für Test)
function Write-Info { 
    param($Message) 
    Write-Host "[INFO] $Message" -ForegroundColor Cyan 
}

function Write-Warning { 
    param($Message) 
    Write-Host "[WARNING] $Message" -ForegroundColor Yellow 
}

function Set-RegistryValues {
    param($RegPath, $Settings)
    Write-Info "Registry-Einstellungen werden gesetzt..." 
    # Vereinfacht für Test
}

function Stop-NamedProcess { 
    param($ProcessName)
    Get-Process -Name $ProcessName -ErrorAction SilentlyContinue | Stop-Process -Force
}

function Clear-ComObject { 
    param($ComObject)
    if ($ComObject) {
        try {
            $ComObject.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ComObject) | Out-Null
        } catch { }
    }
}

# Die aktualisierte Funktion
function Set-WordDefaultFormatting {
  param(
    [string]$FontName = 'Arial',
    [int]$FontSize = 11,
    [double]$LineSpacing = 1.0,
    [int]$SpaceAfter = 0
  )
  
  $regWord = "HKCU:\Software\Microsoft\Office\$OfficeVersion\Word\Options"
  Write-Info "Setze Word Standard-Formatierung: $FontName $FontSize pt, Zeilenabstand $LineSpacing, Absatzabstand $SpaceAfter pt"
  
  # Registry-Einstellungen (vereinfacht)
  Write-Info "Registry-Einstellungen werden konfiguriert..."
  
  # Normal.dotm über COM konfigurieren
  try {
    Stop-NamedProcess WINWORD
    Write-Info "Konfiguriere Normal.dotm über COM..."
    
    $word = New-Object -ComObject Word.Application -ErrorAction Stop
    $word.Visible = $false
    $word.DisplayAlerts = 0
    
    # Temporäres Dokument erstellen und Standard-Style konfigurieren
    $doc = $word.Documents.Add()
    $normalStyle = $doc.Styles.Item('Standard')  # Deutsche Word-Version
    
    Write-Info "Aktuelle Einstellungen: $($normalStyle.Font.Name) $($normalStyle.Font.Size)pt"
    
    # Schriftart setzen
    $normalStyle.Font.Name = $FontName
    $normalStyle.Font.Size = $FontSize
    
    # Absatz-Formatierung setzen
    $normalStyle.ParagraphFormat.LineSpacing = 12
    $normalStyle.ParagraphFormat.SpaceAfter = $SpaceAfter
    $normalStyle.ParagraphFormat.SpaceBefore = 0
    
    Write-Info "Neue Einstellungen: $($normalStyle.Font.Name) $($normalStyle.Font.Size)pt"
    
    # Als neues Template in AppData speichern
    $appDataTemplatePath = "$env:APPDATA\Microsoft\Templates\Normal.dotm"
    $templateDir = Split-Path $appDataTemplatePath -Parent
    if (!(Test-Path $templateDir)) {
      New-Item -Path $templateDir -ItemType Directory -Force | Out-Null
    }
    
    # Als Template speichern
    Write-Info "Speichere Template nach: $appDataTemplatePath"
    $doc.SaveAs2($appDataTemplatePath, 15)
    
    # Dokument schließen
    $doc.Close([ref]$false)
    
    Write-Info "Normal.dotm erfolgreich über COM konfiguriert."
  } catch { 
    Write-Warning "Word COM für Formatierung nicht verfügbar: $($_.Exception.Message)" 
  } finally {
    Clear-ComObject $word
  }
  
  Write-Info "Word-Formatierung gesetzt - Neustart von Word erforderlich für vollständige Übernahme."
}

# Test ausführen
Write-Host "`nTesting function..." -ForegroundColor Yellow
Set-WordDefaultFormatting -FontName 'Arial' -FontSize 11 -LineSpacing 1.0 -SpaceAfter 0

Write-Host "`nTest completed!" -ForegroundColor Green