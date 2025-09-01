# EDITABLE
$excelRelPath = "BEST PRACTICE INDEX.xlsx"
$idHeader = "ID"
$keywordsHeader = "KEYWORDS"
$hiddenHeader = "HIDDEN (REASON FOR HIDING FROM SEARCH)"
$htmlRelPath = "BPsearch.html"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

<#---------------
	 Excel Reading 
	---------------#>
	
$excelPath = Join-Path -Path $scriptDir -ChildPath $excelRelPath

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

$master = $excel.Workbooks.Open($excelPath, $false, $true)
$mainSheet = $master.Worksheets.Item(1)

$table = $mainSheet.ListObjects.Item("Table1")
$numRows = $table.Range.Rows.Count
$numCols = $table.Range.Columns.Count

# Finding columns based off column name
$idCol = 0
$keywordsCol = 0
$hiddenCol = 0
for ($i = 1; $i -le $numCols; $i++) {
	if ($idHeader -eq $table.Range.Cells.Item(1, $i).Text) {
		$idCol = $i
	}
	elseif ($keywordsHeader -eq $table.Range.Cells.Item(1, $i).Text) {
		$keywordsCol = $i
	}
	elseif ($hiddenHeader -eq $table.Range.Cells.Item(1, $i).Text) {
		$hiddenCol = $i
	}
}
# Error checking
if ($idCol -eq 0) {
	Write-Error "`nUnable to find column labeled $($idHeader), if the header was changed reflect the change in $($PSCommandPath), by correcting the idCol variable."
	exit
}
elseif ($keywordsCol -eq 0) {
	Write-Error "`nUnable to find column labeled $($keywordsHeader), if the header was changed reflect the change in $($PSCommandPath), by correcting the keywordsCol variable."
	exit
}
elseif ($hiddenCol -eq 0) {
	Write-Error "`nUnable to find column labeled $($hiddenHeader), if the header was changed reflect the change in $($PSCommandPath), by correcting the hiddenCol variable."
	exit
}

# Searching for row with a matching id
$hashTable = @{}
$hidden = @()
for ($i = 1; $i -le $numRows; $i++) {
	if (-not $table.Range.Cells.Item($i, $hiddenCol).Text) {
		$id = $table.Range.Cells.Item($i, $idCol).Text
		$keywords = $table.Range.Cells.Item($i, $keywordsCol).Text -replace "'", "\'" -Split "\r?\n"
		$hashTable[$id] = $keywords
	}
	else {
		$id = $table.Range.Cells.Item($i, $idCol).Text
		$hidden += $id
	}
}

$master.Close($false)
$excel.Quit()

# Writes the success message to the terminal
Write-Host "`tSuccessfully read $($excelPath)"

<#--------------
	 html Writing 
	--------------#>

$htmlPath = Join-Path -Path $scriptDir -ChildPath $htmlRelPath

$script = ""
$edit = 0
Get-Content -Path $htmlPath | ForEach-Object {
	if ($edit) {	# Inside data, looking to add/edit object
		if ($_ -eq "		];") {	# Found end of data, leaving edit mode
			$script = $script + "`n" + $_ + "`n"
			$edit = 0
		}
		else {	# Updating keywords
			$objects = $_ -Split ":"
			$id = $objects[1] -replace ", title", "" -replace "'", ""
			if ($hashTable[$id]) {	# Matching id, updating keywords
				$keywords = $hashTable[$id]
				$script = $script + "`n" + $objects[0] + ":" + $objects[1] + ":" + $objects[2] + ":["
				for ($i = 0; $i -lt $keywords.Length; $i++) {
					$script = $script + "'" + $keywords[$i].ToLower() + "'"
					if ($i -lt $keywords.Length - 1) {
						$script = $script + ", "
					}
				}
				$script = $script + "], content:"
				for ($i = 4; $i -lt $objects.Length; $i++) {
					$script = $script + $objects[$i]
					if ($i -lt $objects.Length - 1) {
						$script = $script + ":"
					}
				}
			}
			elseif ($hidden.Contains($id)) { # The id is marked as hidden in the index document, but exists in the html
				$script = $script + "`n" + $_
				Write-Warning "The ID: $($id) exists in the html, but is marked as hidden in the index document`n`t`tTo correct, run: ./updateBP $($id)"
			}
			else {	# The id is not in the index document, should not happen but copying data over
				$script = $script + "`n" + $_
				Write-Warning "The ID: $($id) exists in the html, but not in the index document`n`t`tTo correct, open the html and delete the id's line in the data json"
			}
			$hashTable.Remove("$($id)")
		}
	}
	else {	# Outside of data, copying html
		if ($_.TrimStart() -eq "const data = [") {	# Found start of data, entering edit mode
			$script = $script + $_
			$edit = 1
		}
		else {	# Continuing to copy html over
			$script = $script + $_ + "`n"
		}
	}
}
# The id is not in the html, but in the index document
$leftoverIDs = $hashTable.Keys
$leftoverIDS | ForEach-Object {
	Write-Warning "The ID: $($_) exists in the index document, but is not in the html`n`tTo correct, run: ./updateBP $($_)"
}#>

# Writes the script string to the html file, with thw ASCII encoding for the stylesheet linking
$script | Out-File -Encoding ASCII -FilePath $htmlPath -NoNewLine

# Writes the success message to the terminal
Write-Host "Successfully updated BPsearch.html"

# Release COM objects to free up memory
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($master) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
$master = $null
$excel = $null

# Perform garbage collection
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()