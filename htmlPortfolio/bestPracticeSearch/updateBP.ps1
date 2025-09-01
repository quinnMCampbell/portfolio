param (
	[string]$id
)
if (-not $id) {
	Write-Error "`nMissing argument:`n`tUsage:`t./updateBP BPI-BP-XXXX-####"
	exit
}

# EDITABLE
$excelRelPath = "BEST PRACTICE INDEX.xlsx"
$idHeader = "ID"
$titleHeader = "TITLE"
$keywordsHeader = "KEYWORDS"
$wordPathHeader = "WORD DOCUMENT PATH"
$hiddenHeader = "HIDDEN (REASON FOR HIDING FROM SEARCH)"
$htmlRelPath = "BPsearch.html"

function createDataEntry {
	param (
		[string]$id,
		[string]$title,
		[string[]]$keywords,
		[string]$content,
		[string]$pdfPath,
		[string]$description
	)
	$line = "			{ id:'" + $id + "', title:'" + $title + "', keywords:["
	for ($i = 0; $i -lt $keywords.Length; $i++) {
		$line = $line + "'" + $keywords[$i].ToLower() + "'"
		if ($i -lt $keywords.Length - 1) {
			$line = $line + ", "
		}
	}
	$line = $line + "], content:'" + $content + "', pdfPath:'" + $pdfPath + "', description:'" + $description + "', idMatch:0, titleMatches:0, keywordMatches:0, contentMatches:0, contentTypes:0}"
	return $line
}

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

# Finding columns based off column header name
$idCol = 0
$titleCol = 0
$keywordsCol = 0
$wordPathCol = 0
$hiddenCol = 0
for ($i = 1; $i -le $numCols; $i++) {
	$cell = $table.Range.Cells.Item(1, $i).Text
	if ($idHeader -eq $cell) {
		$idCol = $i
	}
	elseif ($titleHeader -eq $cell) {
		$titleCol = $i
	}
	elseif ($keywordsHeader -eq $cell) {
		$keywordsCol = $i
	}
	elseif ($wordPathHeader -eq $cell) {
		$wordPathCol = $i
	}
	elseif ($hiddenHeader -eq $cell) {
		$hiddenCol = $i
	}
}

# Error checking
if ($idCol -eq 0) {
	Write-Error "`nUnable to find column labeled $($idHeader), if the header was changed reflect the change in $($PSCommandPath), by correcting the idCol variable."
	exit
}
elseif ($titleCol -eq 0) {
	Write-Error "`nUnable to find column labeled $($titleHeader), if the header was changed reflect the change in $($PSCommandPath), by correcting the titleCol variable."
	exit
}
elseif ($keywordsCol -eq 0) {
	Write-Error "`nUnable to find column labeled $($keywordsHeader), if the header was changed reflect the change in $($PSCommandPath), by correcting the keywordsCol variable."
	exit
}
elseif ($wordPathCol -eq 0) {
	Write-Error "`nUnable to find column labeled $($wordPathHeader), if the header was changed reflect the change in $($PSCommandPath), by correcting the wordPathCol variable."
	exit
}
elseif ($hiddenCol -eq 0) {
	Write-Error "`nUnable to find column labeled $($hiddenHeader), if the header was changed reflect the change in $($PSCommandPath), by correcting the hiddenCol variable."
	exit
}

# Searching for row with a matching id
for ($i = 1; $i -le $numRows; $i++) {
	$cell = $table.Range.Cells.Item($i, $idCol).Text
	if ($cell -eq $id) {
		$title = $table.Range.Cells.Item($i, $titleCol).Text -replace "'", "\'"
		$keywords = $table.Range.Cells.Item($i, $keywordsCol).Text -replace "'", "\'" -Split "\r?\n" 
		$wordRelPath = $table.Range.Cells.Item($i, $wordPathCol).Text -replace "'", "\'" -replace '"', ''
		if ($table.Range.Cells.Item($i, $hiddenCol).Text) {
			$hidden = $true
		}
		else {
			$hidden = $false
		}
		break
	}
}
if ($i -gt $numRows) {
	Write-Error "`nUnable to find a row with ID: $($ID).`nEnsure that the document ID being updated is in the master document." 
	exit
}

$pdfPath = "./PDFs/" + $id + ".pdf"

$master.Close($false)
$excel.Quit()

# Writes the success message to the terminal
Write-Host "`tSuccessfully read $($excelPath)"

<#--------------
	 Word Reading
	--------------#>

$word = New-Object -ComObject Word.Application
$word.Visible = $false

$wordPath = Join-Path -Path $scriptDir -ChildPath $wordRelPath

$doc = $word.Documents.Open($wordPath, $false, $true) # Path, ConfirmConversions, ReadOnly
$paragraphs = $doc.Paragraphs

# Accessing purpose for a description
$description = $($paragraphs.Item(1).Range.Text -replace "Purpose:", "").Trim() -replace "'", "\'"

# Accessing content
$content = ""
foreach ($para in $paragraphs) {
	$text = $para.Range.Text
	# Process the text as needed, e.g., display it or store it in an array
	$content = $content + $text
}
$content = $content.ToLower() -replace "'", "\'" -replace "`r?`n", " " -replace "\s+", " "

$doc.Close($false)
$word.Quit()

# Writes the success message to the terminal
Write-Host "`tSuccessfully read $($wordPath)"

<#--------------
	 html Writing 
	--------------#>

$htmlPath = Join-Path -Path $scriptDir -ChildPath $htmlRelPath

$script = ""
$edit = 0
Get-Content -Path $htmlPath | ForEach-Object {
	if ($edit) {	# Inside data, looking to add/edit object
		$objects = $_ -Split ":"
		if ($objects[1]) {	# The line includes a colon
			$objects[1] = $objects[1] -replace ", title", "" -replace "'", ""
			if ($objects[1] -eq $id) {	# Matched id, editing object based on excel and word data
				if ($hidden) {	# Line is hidden and is not being added
					if ($objects[$objects.Length-1] -eq "0},") {	# Is not the last object in data
						$script = $script + "`n"
					}
					else {	# Is the last object in data
						$script = $script.TrimEnd(",`n") + "`n"
					}
				}
				else {	# Line is being added
					$line = createDataEntry $id $title $keywords $content $pdfPath $description
					if ($objects[$objects.Length-1] -eq "0},") {	# Is not the last object in data
						$script = $script + "`n" + $line + ",`n"
					}
					else {	# Is the last object in data
						$script = $script + "`n" + $line + "`n"
					}
				}
				$edit = 0
			}
			elseif ($id -lt $objects[1]) {	# No match, greater than id, creating object based on excel and word data
				if ($hidden) {	# Line is hidden and is not being added
					$script = $script + "`n" + $_ + "`n"
				}
				else {	# Line is being added
					$line = createDataEntry $id $title $keywords $content $pdfPath $description
					$script = $script + "`n" + $line + ",`n" + $_ + "`n"
				}
				$edit = 0
			}
			else {	# No match, less than id, copying other data object
				$script = $script + "`n" + $_
			}
		}
		elseif ($_ -eq "		];") {	# Reached the end of data with no matches, creating object based on excel and word data
			if (-not $hidden) {	# Line is being added
				$line = createDataEntry $id $title $keywords $content $pdfPath $description
				$script = $script + ",`n" + $line + "`n" + $_ + "`n"
			}
			$edit = 0
		}
	}
	else {	# Outside of data, copying html
		if ($_.TrimStart() -eq "const data = [") {	# Found start of data, swithcing to edit mode
			$script = $script + $_
			$edit = 1
		}
		else {	# Continuing to copy html over
			$script = $script + $_ + "`n"
		}
	}
}

# Writes the script string to the html file, with thw ASCII encoding for the stylesheet linking
$script | Out-File -Encoding ASCII -FilePath $htmlPath -NoNewLine

# Writes the success message to the terminal
Write-Host "Successfully updated $($htmlPath)"

# Release COM objects to free up memory
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($master) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
$master = $null
$excel = $null
$doc = $null
$word = $null

# Perform garbage collection
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()