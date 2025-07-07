# ConvertBookmarksToExcel.ps1

A PowerShell script that extracts bookmarks from an HTML file exported by your web browser (e.g. Chrome, Firefox, Edge) and converts them to an Excel `.xlsx` file. The output includes folder hierarchy, title, and URL.

## Features

- Parses standard Netscape Bookmark File format
- Extracts nested folder paths
- Decodes HTML entities
- Outputs a tab-delimited `.csv` file
- Converts to Excel `.xlsx` via COM automation (optional)
- No admin rights required

## Supported Browsers

- Google Chrome  
- Mozilla Firefox  
- Microsoft Edge  
- Opera  
- Safari (if no proprietary metadata is present)

## Requirements

- PowerShell 5.1 or later  
- Microsoft Excel (for `.xlsx` export)

## Usage

```powershell
.\ConvertBookmarksToExcel.ps1 [-InputFile "bookmarks.html"] [-OutputXlsx "bookmarks.xlsx"]

## Parameters

| Parameter   | Type   | Description                                                                 |
|-------------|--------|-----------------------------------------------------------------------------|
| `InputFile` | string | Path to the HTML bookmarks file. Default: `.\bookmarks.html`                |
| `OutputXlsx`| string | Path to the resulting Excel file. Default: `.\Output\bookmarks.xlsx`        |

> **Note**: The `.csv` file will be saved automatically in the same folder as the `.xlsx` file, using the `.csv` extension.

## Example

```powershell
.\ConvertBookmarksToExcel.ps1 -InputFile "my_links.html" -OutputXlsx "C:\Exported\Links.xlsx"
