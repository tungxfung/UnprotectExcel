param([string]$Excel="")

If ($Excel -eq "") {
	Write-Output "Failed"
	throw "No excel argument"
}

$ExcelFilePath = Split-Path -Path $excel
$ExcelName = Split-Path -Path $excel -Leaf
$ExcelTempDir = $ExcelFilePath + '\' + $ExcelName + "_temp"
$ExcelFileSaved = $ExcelFilePath + "\" + $ExcelName + "_unprotected.xlsx"

If (Test-Path $ExcelTempDir){
	Remove-Item $ExcelTempDir
}

If (Test-Path $ExcelFileSaved){
	Remove-Item $ExcelFileSaved
}

Add-Type -A System.IO.Compression.FileSystem
[IO.Compression.ZipFile]::ExtractToDirectory($Excel, $ExcelTempDir)

$Input = $ExcelTempDir + "\xl\worksheets\sheet1.xml"
$Output = $ExcelTempDir + "\xl\worksheets\sheet1.xml"

# Load the existing document
$Doc = [xml](Get-Content $Input)

# Specify tag names to delete and then find them
$DeleteNames = "sheetProtection"
($Doc.worksheet.ChildNodes |Where-Object { $DeleteNames -contains $_.Name }) | ForEach-Object {
	# Remove each node from its parent
	[void]$_.ParentNode.RemoveChild($_)
}

# Save the modified document
$Doc.Save($Output)

[System.IO.Compression.ZipFile]::CreateFromDirectory($ExcelTempDir, $ExcelFileSaved) ;

Write-Output "Success"

[Environment]::Exit(200)


