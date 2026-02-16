# export-bookmarks
Exports browser bookmarks to CSV files for Azure domain users with OneDrive storage
A PowerShell utility that exports bookmarks from multiple browsers to CSV format for backup, migration, or analysis purposes.

## 🚀 Features

- **Multi-browser support**: Chrome, Edge, Firefox, and Internet Explorer
- **Automatic browser detection**: Scans for installed browsers on the system
- **Profile support**: Handles multiple browser profiles (Default, Profile 1, Profile 2, etc.)
- **Firefox support**: Uses sqlite3.exe to extract bookmarks from Firefox's places.sqlite database
- **Comprehensive logging**: Creates detailed log files for troubleshooting
- **CSV output**: Easy to import into Excel or other analysis tools

## 📋 Requirements

- Windows operating system
- PowerShell 5.1 or higher
- For Firefox: `sqlite3.exe` in the same folder as the script
- No administrative rights required (runs in user context)

## 🔧 Installation

1. Download all files to a local folder:
   - `ExportBookMark.bat` - Batch wrapper for easy execution
   - `ExportBookMarks.ps1` - Main PowerShell script
   - `sqlite3.exe` - Required for Firefox bookmark export

2. Ensure all files are in the **same directory**

## 🎯 Usage

### Method 1: Double-click (Easiest)
Simply double-click `ExportBookMark.bat` and the script will run automatically.

### Method 2: Command Line
powershell.exe -ExecutionPolicy Bypass -NoProfile -File .\ExportBookMarks.ps1 .

## 📁 Output

The script creates the following in `C:\Temp\bookmarkexport27\`:

| File | Description |
|------|-------------|
| `Chrome_Bookmarks_YYYYMMDD.csv` | Bookmarks from all Chrome profiles |
| `Edge_Bookmarks_YYYYMMDD.csv` | Bookmarks from all Edge profiles |
| `Firefox_Bookmarks_YYYYMMDD.csv` | Bookmarks from Firefox |
| `InternetExplorer_Bookmarks_YYYYMMDD.csv` | Internet Explorer favorites |
| `BookmarkExport_YYYYMMDD_HHmmss.log` | Detailed execution log |

CSV Format
Each CSV file contains:

Title: Bookmark name
URL: Web address
Folder: Bookmark folder structure
DateAdded: When the bookmark was created
Browser: Source browser and profile
ExportDate: When the export was performed .

## 🔍 Browser Support Details

| Browser | Support Level | Method |
|---------|--------------|--------|
| **Chrome** | ✅ Full Support | Reads Bookmarks JSON file |
| **Edge** | ✅ Full Support | Reads Bookmarks JSON file |
| **Firefox** | ✅ Full Support | SQLite query via sqlite3.exe |
| **Internet Explorer** | ✅ Full Support | Parses .url files in Favorites folder |
| *Brave* | ❌ Not Supported | Planned for future version |
| *Opera* | ❌ Not Supported | Planned for future version |
| *Vivaldi* | ❌ Not Supported | Planned for future version |

## Troubleshooting
Firefox not exporting?
Ensure sqlite3.exe is in the same folder as the script
Check if Firefox is installed and has been used at least once
Verify the profiles folder exists: %APPDATA%\Mozilla\Firefox\Profiles\

Check if you have any bookmarks saved
Review the log file for details

Permission errors?
The script runs in user context and doesn't require admin rights
Ensure you have write access to the output folder .

## 📝 Version History
v2.7 (2024-10-25)
Optimized RegEx compilation for faster Chrome/Edge extraction
Enhanced error handling
Improved logging
Multi-profile support

v2.6 (Previous)
Added browser detection
Firefox SQLite integration
CSV output formatting .
