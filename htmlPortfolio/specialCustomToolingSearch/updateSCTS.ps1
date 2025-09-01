# EDITABLE: to change the names of the .html output, .csv input, or the header for the PN column
$htmlName = "Special, Custom Tooling Search.html"
$excelName = "SPECIAL CUSTOM TOOLING INVENTORY.xlsx"
$defaultCriteria = "PART NUMBERS ASSOCIATED WITH TOOL"
$headerBList = @("CATEGORY")
$PNheader = "PART NUMBERS ASSOCIATED WITH TOOL"

# Initializes the paths to the .html and .csv, assumes they're in the same directory as the script
$htmlRelPath = "$($htmlName)"
$excelRelPath = "$($excelName)"

function createDataEntry {
	param (
		[string[]]$headers,
		[string[]]$row
	)
	$line = "			{ "
	for ($i = 0; $i -lt $headers.Length; $i++) {
		$line = $line + "'$($headers[$i])':"
		if ($row[$i].Contains("<p>")) {
			$array = $row[$i] -split "<p>"
			$line = $line + "["
			for ($j = 0; $j -lt $array.Length; $j++) {
				$line = $line + "'$($array[$j])'"
				if ($j -lt $array.Length - 1) {
					$line = $line + ", "
				}
			}
			$line = $line + "]"
		}
		else {
			$line = $line + "'$($row[$i])'"
		}
		if ($i -lt $headers.Length - 1) {
			$line = $line + ", "
		}
	}
	$line = $line + " }"
	return $line
}

function createSelectEntry {
	param (
		[string]$header
	)
	if (-not ($header -in $headerBList)) {	# Checks that the column is not on the Header Black List
		$line = "						<option value='$($header)'"
		if ($header -eq $PNheader) {	# Shortens the Header for PN
			$disp = "PART NUMBER"
		}
		elseif ($header -match "<p>") {	# Headers that include a new line char, only use the first line
			$lines = $header -split '<p>'
			$disp = $lines[0]
		}
		else {	# The remaining headers
			$disp = $header
		}
		if ($header -eq $defaultCriteria) { # Formats the default criteria
			$line = $line + " selected>$($disp)</option>"
		}
		else {	# Formats the other criteria
			$line = $line + ">$($disp)</option>"
		}
		return $line
	}
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Formatting headerBList
for ($i = 0; $i -lt $headerBList.Length; $i++) {
	$headerBList[$i] = $headerBList[$i] -replace "`r?`n", "<p>"
}

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

# Finding table contents
$contents = @($null) * $numRows

# Body rows are handled here
for ($i = 0; $i -lt $numRows; $i++) {
	$row = @($null) * $numCols
	for ($j = 0; $j -lt $numCols; $j++) {
		$row[$j] = $table.Range.Cells.Item($i + 1, $j + 1).Text -replace "`r?`n", "<p>" -replace "'", "\'"
	}
	$contents[$i] = $row
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
	if ($edit -eq 0) { # Outside of data and select, copying html
		if ($_.TrimStart() -eq "const data = [") {	# Found start of data, swithcing to edit = 1
			$script = $script + $_ + "`n"
			$edit = 1
		}
		elseif ($_.TrimStart() -eq "<select id='searchSelect'>") {	# Found start of select, switching to edit = 2
			$script = $script + $_ + "`n"
			$edit = 2
		}
		else {	# Continuing to copy html over
			$script = $script + $_ + "`n"
		}
	}
	elseif ($edit -eq 1) {	# Inside data, looking to add/edit object
		if ($_.TrimStart() -eq "];") {	# Reached the end of data
			for ($i = 1; $i -lt $numRows; $i++) {
				$line = createDataEntry $contents[0] $contents[$i]
				if ($i -lt $($numRows - 1)) {
					$script = $script + $line + ",`n"
				}
				else {
					$script = $script + $line + "`n"
				}
			}			
			$script = $script + $_ + "`n"
			$edit = 0
		}
	}
	elseif ($edit -eq 2) {
		if ($_.TrimStart() -eq "</select>") {	# Reached the end of select
			for ($i = 0; $i -lt $numCols; $i++) {	# Reached the end of select
				$line = createSelectEntry $contents[0][$i]
				if ($line) {
					$script = $script + $line + "`n"
				}
			}
			$script = $script + $_ + "`n"
			$edit = 0
		}
	}
}

# Writes the script string to the html file, with thw ASCII encoding for the stylesheet linking
$script | Out-File -Encoding ASCII -FilePath $htmlPath -NoNewLine

# Writes the success message to the terminal
Write-Host "Successfully updated $($htmlName)"

# Release COM objects to free up memory
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($master) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
$master = $null
$excel = $null

# Perform garbage collection
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()