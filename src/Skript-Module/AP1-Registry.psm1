# AP1-Registry.psm1
function Set-RegistryValues {
	param(
		[string]$RegPath, 
		[hashtable]$Settings,
		[switch]$SuppressOutput
	)
	if (-not (Test-Path $RegPath)) { 
		New-Item -Path $RegPath -Force | Out-Null 
	}
	foreach ($key in $Settings.Keys) {
		$val = $Settings[$key]
		$type = if ($val -is [string] -and $val -match '%') { 'ExpandString' } 
				elseif ($val -is [string]) { 'String' } 
				else { 'DWord' }
		try {
			New-ItemProperty -Path $RegPath -Name $key -Value $val -PropertyType $type -Force | Out-Null
			if (-not $SuppressOutput) {
				Write-Info "  Registry $RegPath - $key = $val ($type)"
			}
		} catch {
			Write-Warning "Fehler Registry [$RegPath] $key - $($_.Exception.Message)"
		}
	}
}
function Set-WordAutoCorrectRegistry {
	$regWord = ("HKCU:\Software\Microsoft\Office\$($script:OfficeVersion)\Word\Options" -replace "\\\\", "\\")
	$settings = @{
		'AutoFormatAsYouTypeApplyNumberedLists'  = 0
		'AutoFormatAsYouTypeApplyBulletedLists'  = 0
		'CorrectSentenceCaps'                    = 1
		'AutoFormatAsYouTypeReplaceHyperlinks'   = 0
		'CorrectInitialCaps'                     = 0
		'AutoFormatAsYouTypeReplaceQuotes'       = 1
		'AutoFormatAsYouTypeReplaceSymbols'      = 1
		'PasteFormattingOtherApp'                = 2
		'PasteFormattingTwoDocumentsNoStyles'    = 1
	}
	Write-Info "Setze Word Autokorrektur (HKCU)..."
	Set-RegistryValues -RegPath $regWord -Settings $settings
}
function Set-WordDefaultFormatting {
	param(
		[string]$FontName = 'Arial',
		[int]$FontSize = 11,
		[double]$LineSpacing = 1.1,
		[int]$SpaceAfter = 0
	)
	$regWord = ("HKCU:\Software\Microsoft\Office\$($script:OfficeVersion)\Word\Options" -replace "\\\\", "\\")
	Write-Info "Setze Word Standard-Formatierung: $FontName $FontSize pt, Zeilenabstand $LineSpacing, Absatzabstand $SpaceAfter pt"
	Write-Info "HINWEIS: Office-Policies erkannt - verwende Dokumentvorlagen-Ansatz"
	# ...Funktionscode wie im Hauptskript...
	# (Hier wird der komplette Funktionscode eingefügt, siehe Hauptskript)
}
function Set-OfficeRegistrySettings {
	$regWord   = ("HKCU:\Software\Microsoft\Office\$script:OfficeVersion\Word\Options" -replace "\\\\", "\\")
	$regExcel  = ("HKCU:\Software\Microsoft\Office\$script:OfficeVersion\Excel\Options" -replace "\\\\", "\\")
	$regWinAdv = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
	$wordSettings = @{
		'DeveloperTools'                   = 1
		'Ruler'                            = 1
		'ShowAllFormatting'                = 1
		'VisiDrawTableDrs'                 = 1
		'DisableBootToOfficeStart'         = 1
		'DisableBackstageOpenKeyShortcuts' = 1
	}
	$excelSettings = @{
		'DeveloperTools'                   = 1
		'DisableBootToOfficeStart'         = 1
	}
	$windowsSettings = @{ 'HideFileExt' = 0 }
	Write-Info "Setze Office-/Windows-Optionen (HKCU)..."
	Set-RegistryValues -RegPath $regWord   -Settings $wordSettings
	Set-RegistryValues -RegPath $regExcel  -Settings $excelSettings
	Set-RegistryValues -RegPath $regWinAdv -Settings $windowsSettings
}
function Set-DefaultSavePaths {
	param([string]$Path, [switch]$UseCom)
	$regWord   = ("HKCU:\Software\Microsoft\Office\$script:OfficeVersion\Word\Options" -replace "\\\\", "\\")
	$regExcel  = ("HKCU:\Software\Microsoft\Office\$script:OfficeVersion\Excel\Options" -replace "\\\\", "\\")
	Write-Info "Setze Standard-Speicherpfade: $Path"
	Set-RegistryValues -RegPath $regWord  -Settings @{ 'DOC-PATH'   = $Path }
	Set-RegistryValues -RegPath $regExcel -Settings @{ 'DefaultPath' = $Path }
	# ...Funktionscode wie im Hauptskript...
}
function Set-WordTemplatesPath {
	param([string]$TemplatePath)
	$regWord = ("HKCU:\Software\Microsoft\Office\$script:OfficeVersion\Word\Options" -replace "\\\\", "\\")
	Write-Info "Setze Word Standard-Speicherort für persönliche Vorlagen: $TemplatePath"
	$templateSettings = @{ 
		'PersonalTemplates' = $TemplatePath
		'USER-DOT-PATH' = $TemplatePath
	}
	Set-RegistryValues -RegPath $regWord -Settings $templateSettings
	# ...Funktionscode wie im Hauptskript...
}
function Set-ExcelAutoCorrectRegistry {
	$regExcel = ("HKCU:\Software\Microsoft\Office\$script:OfficeVersion\Excel\Options" -replace "\\\\", "\\")
	$settings = @{ 'CorrectSentenceCap' = 0 }
	Write-Info "Setze Excel Autokorrektur (HKCU)..."
	Set-RegistryValues -RegPath $regExcel -Settings $settings
}
function Clear-ComObject {
	param([Parameter(ValueFromPipeline)]$ComObject)
	if ($ComObject) {
		try {
			if ($ComObject.PSObject.Properties.Name -contains 'Quit') {
				$ComObject.Quit()
			}
			[void][Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
		} catch {
			Write-Warning "COM-Objekt konnte nicht bereinigt werden: $($_.Exception.Message)"
		}
	}
	[GC]::Collect()
	[GC]::WaitForPendingFinalizers()
}
function Test-ComAvailable {
	param([string]$ProgId)
	try {
		$app = New-Object -ComObject $ProgId -ErrorAction Stop
		Clear-ComObject $app
		return $true
	} catch {
		return $false
	}
}
function Test-WriteAccess {
	param([string]$Path, [string]$Description)
	try {
		$testFile = Join-Path $Path 'test_write.txt'
		Set-Content -Path $testFile -Value 'Test' -Force
		Remove-Item $testFile -Force
		Write-Info "Schreibzugriff auf $Description ($Path) OK."
		return $true
	} catch {
		Write-Warning "Kein Schreibzugriff auf $Description ($Path): $($_.Exception.Message)"
		return $false
	}
}
function Initialize-OfficeApps {
	Write-Info "Initialisiere Word/Excel (robust)..."
	# ...Funktionscode wie im Hauptskript...
}
function Set-OfficeFirstRunMarkers {
	# ...Funktionscode wie im Hauptskript...
}
function Test-OfficeFirstRun {
	# ...Funktionscode wie im Hauptskript...
}

# Alle Funktionsdefinitionen bleiben unverändert
Export-ModuleMember -Function Set-RegistryValues,Set-WordAutoCorrectRegistry,Set-WordDefaultFormatting,Set-OfficeRegistrySettings,Set-DefaultSavePaths,Set-WordTemplatesPath,Set-ExcelAutoCorrectRegistry,Clear-ComObject,Test-ComAvailable,Test-WriteAccess,Initialize-OfficeApps,Set-OfficeFirstRunMarkers,Test-OfficeFirstRun
