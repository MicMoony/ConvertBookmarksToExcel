# ConvertBookmarksToExcel.ps1

Ever find yourself collecting dozens, or even hundreds, of browser bookmarks over time, only to lose track of what’s where? You're not alone. Staying organized gets tricky fast. But that ends now.

This PowerShell script helps bring order to the chaos by extracting bookmarks from an HTML file exported by your web browser (like Chrome, Firefox, Edge) and converting them into a clean, searchable Excel `.xlsx` file. The output includes the full folder hierarchy, bookmark titles, and URLs, all neatly arranged for easy viewing.

## Features

- Parses standard Netscape Bookmark File format
- Extracts nested folder paths
- Decodes HTML entities
- Outputs a tab-delimited `.csv` file
- Converts to Excel `.xlsx` via COM automation
- No administrative privileges required

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
```

## Parameters

| Parameter   | Type   | Description                                                                 |
|-------------|--------|-----------------------------------------------------------------------------|
| `InputFile` | string | Path to the HTML bookmarks file. Default: `.\bookmarks.html`                |
| `OutputXlsx`| string | Path to the resulting Excel file. Default: `.\Output\bookmarks.xlsx`        |

> **Note**: The `.csv` file will be saved automatically in the same folder as the `.xlsx` file, using the `.csv` extension.

## Example

```powershell
.\ConvertBookmarksToExcel.ps1 -InputFile "MyEdgeBookmarks.html" -OutputXlsx "C:\Export\MyEdgeBookmarks_List.xlsx"
```
