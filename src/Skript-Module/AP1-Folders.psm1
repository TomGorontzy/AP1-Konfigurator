# AP1-Folders.psm1
function New-CandidateFoldersFromExcel {
	param([Parameter(Mandatory)] [string]$WorkbookPath, [int]$MaxRows = 500)
	$rootPath = Join-Path $script:ScriptRoot '2. Bei Bedarf anpassen\Ordner'
	New-EnsuredPath $rootPath
	$excel = $null; $wb = $null
	try {
		$excel = New-Object -ComObject Excel.Application -ErrorAction Stop
		$excel.Visible = $false; $excel.DisplayAlerts = $false
		$wb = $excel.Workbooks.Open($WorkbookPath)
		$sheet = $wb.Sheets.Item(1)
		$used = $sheet.UsedRange
		$rowCount = [Math]::Min($used.Rows.Count, $MaxRows)
		Write-Info "Erzeuge Ordnerstruktur aus Excel (max. $rowCount Zeilen)..."
		for ($r = 1; $r -le $rowCount; $r++) {
			$a = $sheet.Cells.Item($r,1).Text
			$b = $sheet.Cells.Item($r,2).Text
			if ([string]::IsNullOrWhiteSpace($a) -or [string]::IsNullOrWhiteSpace($b)) { continue }
			$safeA = ($a -replace '[\\/:*?"<>|]', '_').Trim()
			$safeB = ($b -replace '[\\/:*?"<>|]', '_').Trim()
			if ([string]::IsNullOrWhiteSpace($safeA) -or [string]::IsNullOrWhiteSpace($safeB)) { continue }
			$path = Join-Path (Join-Path $rootPath $safeA) $safeB
			New-EnsuredPath $path
		}
		Write-Info "Ordner fuer Pruefungskandidaten wurden angelegt."
	} catch {
		throw
	} finally {
		if ($wb) { 
			try { $wb.Close($false) | Out-Null } catch {}
			Clear-ComObject $wb
		}
		if ($excel) { 
			try { $excel.DisplayAlerts = $true } catch {}
			Clear-ComObject $excel
		}
		Stop-NamedProcess EXCEL
	}
	return $rootPath
}
function New-CandidateFoldersFromCsv {
	param([Parameter(Mandatory)][string]$CsvPath, [int]$MaxRows=500)
	if (-not (Test-Path $CsvPath)) { throw "CSV nicht gefunden: $CsvPath" }
	$rootPath = Join-Path $script:ScriptRoot '2. Bei Bedarf anpassen\Ordner'
	New-EnsuredPath $rootPath
	$rows = Import-Csv -Path $CsvPath -Delimiter ';' -Header 'Account','Kandidat'
	$i = 0
	foreach ($row in $rows) {
		if ($i -ge $MaxRows) { break }
		$a = $row.Account; $b = $row.Kandidat
		if ([string]::IsNullOrWhiteSpace($a) -or [string]::IsNullOrWhiteSpace($b)) { continue }
		$safeA = ($a -replace '[\\/:*?"<>|]', '_').Trim()
		$safeB = ($b -replace '[\\/:*?"<>|]', '_').Trim()
		if ([string]::IsNullOrWhiteSpace($safeA) -or [string]::IsNullOrWhiteSpace($safeB)) { continue }
		$path = Join-Path (Join-Path $rootPath $safeA) $safeB
		New-EnsuredPath $path
		$i++
	}
	Write-Info "Ordner aus CSV angelegt (${i}) Zeilen."
	return $rootPath
}
function Copy-CandidateFolderToDesktop {
	param([string]$SourceRoot)
	$userName = $env:UserName
	$desktop  = Get-DesktopPath
	$source   = Join-Path $SourceRoot $userName
	if (-not (Test-Path $source)) { 
		Write-Warning "Kein Kandidaten-Ordner fuer aktuellen Nutzer: $source"
		$availableFolders = Get-ChildItem -Path $SourceRoot -Directory -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name
		if ($availableFolders) {
			Write-Info "Verfuegbare Benutzerordner: $($availableFolders -join ', ')"
			$similarFolder = $availableFolders | Where-Object { $_ -like "*$userName*" -or $userName -like "*$_*" }
			if ($similarFolder) {
				Write-Info "Aehnlicher Ordner gefunden: $similarFolder - verwende diesen stattdessen"
				$source = Join-Path $SourceRoot $similarFolder
			}
		}
		if (-not (Test-Path $source)) { return }
	}
	# Suche den ersten Unterordner (Kandidatenordner)
	$candidateFolder = Get-ChildItem -Path $source -Directory | Select-Object -First 1
	if (-not $candidateFolder) {
		Write-Warning "Kein Kandidatenordner im Benutzerordner gefunden: $source"
		return
	}
	$target = Join-Path $desktop $candidateFolder.Name
	try {
		if (Test-Path $target) {
			Remove-Item $target -Recurse -Force -ErrorAction Stop
		}
		Copy-Item -Path $candidateFolder.FullName -Destination $target -Recurse -Force
		Write-Info "Kandidaten-Ordner auf Desktop bereitgestellt: $target"
	} catch { 
		Write-Warning "Kandidaten-Ordner konnte nicht kopiert werden: $($_.Exception.Message)" 
	}
}
function New-EnsuredPath {
	param([string]$Path)
	$Path = $Path -replace 'NÃ¼', 'Nü' -replace 'Ã¤', 'ä' -replace 'Ã¶', 'ö' -replace 'ÃÖ', 'Ö' -replace 'ÃŸ', 'ß'
	if (-not (Test-Path $Path)) { New-Item -Path $Path -ItemType Directory -Force | Out-Null }
}
function Join-PathSafe {
	param([string]$Path, [string]$ChildPath)
	$result = Join-Path -Path $Path -ChildPath $ChildPath
	$result = $result -replace 'NÃ¼', 'Nü' -replace 'Ã¤', 'ä' -replace 'Ã¶', 'ö' -replace 'ÃÖ', 'Ö' -replace 'ÃŸ', 'ß'
	return $result
}
function Get-DesktopPath { [Environment]::GetFolderPath('Desktop') }

# Alle Funktionsdefinitionen bleiben unverändert
Export-ModuleMember -Function New-CandidateFoldersFromExcel,New-CandidateFoldersFromCsv,Copy-CandidateFolderToDesktop,New-EnsuredPath,Join-PathSafe,Get-DesktopPath
