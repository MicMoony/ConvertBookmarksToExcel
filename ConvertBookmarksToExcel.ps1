<#
-----------------------------------------------------------------------------
Script Name:     ConvertBookmarksToExcel.ps1
Author:          MicMoony
Version:         1.0
Created:         2025-05-29
License:         MIT License — https://opensource.org/licenses/MIT

Description:
  This script parses a bookmarks HTML file exported from a web browser (e.g. Chrome, Firefox, Edge),
  extracts bookmark titles, URLs, and folder hierarchy, and writes the data to a tab-delimited CSV file.
  And (if Excel is found) converts the CSV to an Excel (.xlsx) file using COM automation.

  The script supports the standard Netscape Bookmark File format used by most major browsers.
  It handles nested folders, decodes HTML entities, and preserves folder paths for clearer organization.

Usage:
  .\ConvertBookmarksToExcel.ps1 [-InputFile "bookmarks.html"] [-OutputFile "bookmarks.xlsx"]

Parameters:
  InputFile   [string]  - Path to the bookmarks HTML file.
                          Default: "$PSScriptRoot\bookmarks.html"

  OutputFile  [string]  - Path to the output Excel (.xlsx) file.
                          Default: "$PSScriptRoot\Output\bookmarks.xlsx"
                          The CSV file will be created temporarily during processing.

Supported Browsers:
  - Google Chrome
  - Mozilla Firefox
  - Microsoft Edge
  - Opera
  - Safari (if no proprietary metadata is present)

Notes:
  - Requires Microsoft Excel for .xlsx export.
  - Does not require administrative privileges to run.
-----------------------------------------------------------------------------
#>

param (
    [string]$inputFile = "$PSScriptRoot\bookmarks.html",
    [string]$outputFile = "$PSScriptRoot\Output\bookmarks.xlsx"
)

# Create a temporary CSV file
$tempCsv = [System.IO.Path]::GetTempFileName()
$tempCsv = [System.IO.Path]::ChangeExtension($tempCsv, ".csv")

# Ensure the output directory exists
$outputDir = [System.IO.Path]::GetDirectoryName($outputFile)
if (-not (Test-Path $outputDir)) {
    New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
}

Add-Type -AssemblyName System.Web

# Initialize list to store bookmarks and stack to track folder hierarchy
[System.Collections.Generic.List[object]]$bookmarks = @()
$folderStack = @()

foreach ($line in Get-Content $inputFile) {
    $trimmed = $line.Trim()

    # Detect folder start (H3 tag), decode and push to stack
    if ($trimmed -match "<H3[^>]*>(.*?)</H3>") {
        $folderName = [System.Web.HttpUtility]::HtmlDecode($matches[1].Trim())
        $folderStack += $folderName
    }
    # Detect folder end (</DL> tag), pop last folder from stack
    elseif ($trimmed -match "</DL>") {
        if ($folderStack.Count -gt 0) {
            $folderStack = $folderStack[0..($folderStack.Count - 2)]
        }
    }
    # Detect bookmark entry, decode and add to list with full folder path
    elseif ($trimmed -match '<A[^>]*HREF="([^"]+)"[^>]*>(.*?)</A>') {
        $url = [System.Web.HttpUtility]::HtmlDecode($matches[1].Trim())
        $title = [System.Web.HttpUtility]::HtmlDecode($matches[2].Trim())
        $folderPath = $folderStack -join " / "

        $bookmarks.Add([PSCustomObject]@{
            Folder = $folderPath
            Title  = $title
            URL    = $url
        })
    }
}

# Export to tab-delimited text file (TSV for Excel)
# Prepare header and lines manually
$lines = @()
$lines += "Folder`tTitle`tURL"

foreach ($bookmark in $bookmarks) {
    $line = "$($bookmark.Folder)`t$($bookmark.Title)`t$($bookmark.URL)"
    $lines += $line
}

# Write to file with UTF-8 BOM
$utf8BomEncoding = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllLines($tempCsv, $lines, $utf8BomEncoding)

try {
    # Open in Excel and save as .xlsx
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false

    # Open the tab-delimited text file using Excel's OpenText method
    $excel.Workbooks.OpenText(
        $tempCsv,
        "65001",               # Origin = UTF-8 code page as string
        1,                     # StartRow
        1,                     # DataType (1 = xlDelimited)
        1,                     # TextQualifier (1 = xlTextQualifierDoubleQuote)
        $false,                # ConsecutiveDelimiter
        $true,                 # Tab
        $false,                # Semicolon
        $false,                # Comma
        $false,                # Space
        $false                 # Other
    )

    # Get the active workbook after OpenText
    $workbook = $excel.ActiveWorkbook

    # Save as .xlsx
    $excel.DisplayAlerts = $false
    $workbook.SaveAs($outputFile, 51)  # 51 = xlOpenXMLWorkbook
    $workbook.Close($false)
    $excel.Quit()

    Write-Host "Exported $($bookmarks.Count) bookmarks to $outputFile"
}
catch {
    Write-Warning "Excel export failed: $_"
}
finally {
    # Release Excel COM object and clean up to avoid lingering Excel.exe
    if ($excel) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }

    # Delete the temporary CSV
    if (Test-Path $tempCsv) {
        Remove-Item $tempCsv -Force
    }
}
