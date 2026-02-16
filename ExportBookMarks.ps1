# ExportBookMarks.ps1 - 241025
# Version: v2.7 (Optimized: RegEx Compilation and Error Handling)Version 27
# SCOPE
#    Exports browser bookmarks to CSV files for Azure domain users with OneDrive storage. Results will be explored to a subfolder .\Results27
# DESCRIPTION
#    Detects installed browsers and exports their bookmarks to CSV files for aggregation.
#    Uses local sqlite3.exe for Firefox bookmark extraction.
#    Uses ultra-tolerant RegEx for Chrome/Edge bookmarks to bypass security/environment blocks.
# Easy Usage Proposla: powershell.exe -ExecutionPolicy Bypass -NoProfile -File .\ExportBookMarks.ps1

# --- Global Configuration ---
# Define output root path - modify this value as needed
$OutputRootPath = "C:\Temp\bookmarkexport"
$Script:SqlitePath = Join-Path $PSScriptRoot "sqlite3.exe"
# Cache the critical AppData paths at script start for stability (v2.6 fix)
$Script:LocalAppDataPath = $env:LOCALAPPDATA
$Script:AppDataPath = $env:APPDATA

# --- Initialization of Log and Output Path ---
$LogFile = $null 

function Initialize-Log {
    param([string]$RootPath)
    # Use Results27 to ensure a clean test run
    $script:OutputRootPath = "$($RootPath)27"
    if (-not (Test-Path $script:OutputRootPath)) {
        Write-Host "Creating output directory: $($script:OutputRootPath)"
        New-Item -ItemType Directory -Path $script:OutputRootPath -Force -ErrorAction Stop | Out-Null
    }
    
    # Define LogFile path after the directory exists
    $script:LogFile = Join-Path $script:OutputRootPath "BookmarkExport_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    Write-Log "Log file initialized." "DEBUG"
}

function Write-Log {
    param([string]$Message, [string]$Type = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Type] $Message"
    Write-Host $logEntry
    
    if ($script:LogFile) {
        Add-Content -Path $script:LogFile -Value $logEntry -ErrorAction SilentlyContinue
    }
}
# Call initialization before first Write-Log call
Initialize-Log -RootPath $OutputRootPath
Write-Log "Starting browser bookmarks export process (v2.7 - Optimized Final Version)"
Write-Log "Script directory: $PSScriptRoot"
Write-Log "Output path: $Script:OutputRootPath"


function Test-Sqlite {
    if (Test-Path $Script:SqlitePath) {
        Write-Log "sqlite3.exe found at: $Script:SqlitePath"
        return $true
    } else {
        Write-Log "sqlite3.exe not found in script directory: $PSScriptRoot" "ERROR"
        Write-Log "Please ensure sqlite3.exe is placed in the same folder as this script" "WARNING"
        return $false
    }
}

function Get-InstalledBrowsers {
    # This function uses native $env: variables for initial detection only
    $browsers = @()
    
    # Chrome detection
    $chromePaths = @(
        "$env:LOCALAPPDATA\Google\Chrome",
        "$env:PROGRAMFILES\Google\Chrome",
        "$env:PROGRAMFILES(X86)\Google\Chrome"
    )
    
    foreach ($path in $chromePaths) {
        if (Test-Path $path) {
            $browsers += "Chrome"
            Write-Log "Chrome detected at: $path"
            break
        }
    }
    
    # Edge detection (Chromium-based)
    $edgePaths = @(
        "$env:LOCALAPPDATA\Microsoft\Edge",
        "$env:PROGRAMFILES\Microsoft\Edge",
        "$env:PROGRAMFILES(X86)\Microsoft\Edge"
    )
    
    foreach ($path in $edgePaths) {
        if (Test-Path $path) {
            $browsers += "Edge"
            Write-Log "Edge detected at: $path"
            break
        }
    }
    
    # Firefox detection
    $firefoxPaths = @(
        "$env:APPDATA\Mozilla\Firefox",
        "$env:PROGRAMFILES\Mozilla Firefox",
        "$env:PROGRAMFILES(X86)\Mozilla Firefox"
    )
    
    foreach ($path in $firefoxPaths) {
        if (Test-Path $path) {
            $browsers += "Firefox"
            Write-Log "Firefox detected at: $path"
            break
        }
    }
    
    # Internet Explorer (always available on Windows)
    $browsers += "Internet Explorer"
    Write-Log "Internet Explorer detected (Windows default)"
    
    return $browsers | Sort-Object -Unique
}

# --- CHROME/EDGE BOOKMARK EXTRACTION (v2.7) ---
function Export-ChromeBookmarks {
    param([string]$BrowserName, [string]$BookmarksPath, [string]$OutputCSV)
    
    Write-Log "Exporting $BrowserName bookmarks from: $BookmarksPath using v2.7 (Optimized RegEx)"
    
    if (-not (Test-Path $BookmarksPath)) {
        Write-Log "$BrowserName bookmarks file not found: $BookmarksPath" "WARNING"
        return $false
    }
    
    $tempBookmarksPath = $null
    try {
        # Copy file to TEMP location to avoid locking issues
        $tempBookmarksPath = Join-Path $env:TEMP "Bookmarks_$([System.Guid]::NewGuid().ToString('N')).json"
        Write-Log "Attempting to copy $BookmarksPath to temporary file: $tempBookmarksPath" "DEBUG"
        Copy-Item $BookmarksPath $tempBookmarksPath -Force -ErrorAction Stop
        
        # Read the file content
        $fileContent = Get-Content $tempBookmarksPath -Raw -Encoding UTF8 -ErrorAction Stop
        
        # Optimized: Compile the RegEx pattern for faster repeated matching
        $RegExOptions = [System.Text.RegularExpressions.RegexOptions]::Singleline -bor [System.Text.RegularExpressions.RegexOptions]::Compiled
        $regex = [regex]::new('"name":\s*"(?<Name>[^"]*?)".*?"type":\s*"url".*?"url":\s*"(?<URL>[^"]*?)".*?"date_added":\s*"(?<DateAdded>[^"]*?)"', $RegExOptions)
        
        $matches = $regex.Matches($fileContent)
        $bookmarks = @()
        
        if ($matches.Count -eq 0) {
            Write-Log "RegEx failed to find any URLs in the $BrowserName bookmark file." "WARNING"
            return $false
        }
        
        Write-Log "RegEx matched $($matches.Count) potential URL entries." "DEBUG"
        
        foreach ($match in $matches) {
            $name = $match.Groups["Name"].Value
            $url = $match.Groups["URL"].Value
            $dateAddedRaw = $match.Groups["DateAdded"].Value
            
            $dateObj = $null
            if ($dateAddedRaw) { 
                try {
                    # Convert Chrome/Edge time (microseconds since 1601-01-01) to DateTime
                    $dateObj = [datetime]::FromFileTime(($dateAddedRaw -replace '[^0-9]','') -as [long] * 10)
                } catch {
                    $null
                }
            }
            
            $bookmarks += [PSCustomObject]@{
                Title = $name
                URL = $url
                Folder = "RegEx Extraction (Folder Unknown)" # Folder cannot be reliably determined via RegEx
                DateAdded = $dateObj
                Browser = $BrowserName
                ExportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
        }
        
        if ($bookmarks.Count -gt 0) {
            # Use string literal "UTF8" for PowerShell 5.1 compatibility
            $csvParams = @{
                Path            = $OutputCSV
                NoTypeInformation = $true
                Encoding        = "UTF8" 
            }
            # Append only if the file already exists (to handle multiple profiles)
            if (Test-Path $OutputCSV) {
                $csvParams.Add("Append", $true)
            }
            
            $bookmarks | Export-Csv @csvParams -ErrorAction Stop
            
            Write-Log "Successfully exported $($bookmarks.Count) ${BrowserName} bookmarks to: $OutputCSV"
            return $true
        } else {
            Write-Log "No bookmarks found in $BrowserName profile (0 URLs extracted after v2.7 RegEx parsing)." "WARNING"
            return $false
        }
    }
    catch {
        Write-Log "Critical Error exporting $BrowserName bookmarks. Error: $($_.Exception.Message)" "ERROR"
        return $false
    }
    finally {
        # Clean up temporary file
        if (Test-Path $tempBookmarksPath) {
            Remove-Item $tempBookmarksPath -Force -ErrorAction SilentlyContinue
        }
    }
}


function Export-FirefoxBookmarks {
    param([string]$OutputCSV)
    
    Write-Log "Exporting Firefox bookmarks"
    
    # Use cached AppData path for stability
    $firefoxProfilesPath = Join-Path $Script:AppDataPath "Mozilla\Firefox\Profiles"
    if (-not (Test-Path $firefoxProfilesPath)) {
        Write-Log "Firefox profiles path not found: $firefoxProfilesPath" "WARNING"
        return $false
    }
    
    if (-not (Test-Sqlite)) {
        Write-Log "Cannot export Firefox bookmarks without sqlite3.exe" "ERROR"
        return $false
    }
    
    try {
        $profiles = Get-ChildItem $firefoxProfilesPath -Directory -ErrorAction Stop | Where-Object { 
            Test-Path (Join-Path $_.FullName "places.sqlite") 
        }
        
        if ($profiles.Count -eq 0) {
            Write-Log "No Firefox profiles with bookmarks database found" "WARNING"
            return $false
        }
        
        $bookmarks = @()
        
        foreach ($profile in $profiles) {
            $placesDB = Join-Path $profile.FullName "places.sqlite"
            Write-Log "Processing Firefox profile: $($profile.Name)"
            
            # Copy database to avoid locking issues
            $tempDB = Join-Path $env:TEMP "places_$([System.Guid]::NewGuid().ToString('N')).sqlite"
            Copy-Item $placesDB $tempDB -Force -ErrorAction Stop
            
            # SQL query to extract bookmarks
            $query = @"
SELECT 
    b.title as Title,
    p.url as URL,
    f.title as Folder,
    datetime(b.dateAdded/1000000, 'unixepoch') as DateAdded
FROM moz_bookmarks b
LEFT JOIN moz_places p ON b.fk = p.id
LEFT JOIN moz_bookmarks f ON b.parent = f.id
WHERE b.type = 1 AND p.url IS NOT NULL AND p.url != ''
"@
            
            # Escape query for command line
            $escapedQuery = $query -replace '"', '\"'
            
            # Execute sqlite3 command (2>&1 redirects stderr to stdout)
            $result = & $Script:SqlitePath -header -csv $tempDB $escapedQuery 2>&1
            
            if ($LASTEXITCODE -eq 0 -and $result -and $result.Length -gt 1) {
                # Convert CSV output to objects
                $bookmarkData = $result | ConvertFrom-Csv
                foreach ($bookmark in $bookmarkData) {
                    $bookmarks += [PSCustomObject]@{
                        Title = $bookmark.Title
                        URL = $bookmark.URL
                        Folder = $bookmark.Folder
                        DateAdded = $bookmark.DateAdded
                        Browser = "Firefox ($($profile.Name))"
                        ExportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    }
                }
                Write-Log "Found $($bookmarkData.Count) bookmarks in profile $($profile.Name)"
            } else {
                Write-Log "No bookmarks found in profile $($profile.Name) or query failed" "WARNING"
            }
            
            # Clean up temporary file
            Remove-Item $tempDB -Force -ErrorAction SilentlyContinue
        }
        
        if ($bookmarks.Count -gt 0) {
            # Use string literal "UTF8" for PowerShell 5.1 compatibility
            $bookmarks | Export-Csv -Path $OutputCSV -NoTypeInformation -Encoding "UTF8" -ErrorAction Stop
            Write-Log "Successfully exported $($bookmarks.Count) Firefox bookmarks to: $OutputCSV"
            return $true
        } else {
            Write-Log "No Firefox bookmarks found in any profile" "WARNING"
            return $false
        }
    }
    catch {
        Write-Log "Error exporting Firefox bookmarks: $($_.Exception.Message)" "ERROR"
        return $false
    }
}


function Export-IEBookmarks {
    param([string]$OutputCSV)
    
    Write-Log "Exporting Internet Explorer Favorites"
    
    $ieFavoritesPath = [System.Environment]::GetFolderPath('Favorites')
    if (-not (Test-Path $ieFavoritesPath)) {
        Write-Log "IE Favorites folder not found: $ieFavoritesPath" "WARNING"
        return $false
    }
    
    try {
        $favorites = @()
        
        function Get-IEFavorites {
            param($folderPath, $currentFolder)
            
            $items = Get-ChildItem $folderPath -ErrorAction Stop
            
            foreach ($item in $items) {
                if ($item.Extension -eq '.url') {
                    # Parse .url files
                    $content = Get-Content $item.FullName -ErrorAction Stop
                    if ($content) {
                        $url = ($content | Where-Object { $_ -match '^URL=' } | Select-Object -First 1) -replace '^URL=',''
                        
                        $favorites += [PSCustomObject]@{
                            Title = $item.BaseName
                            URL = $url
                            Folder = $currentFolder
                            DateAdded = $item.CreationTime
                            Browser = "Internet Explorer"
                            ExportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                        }
                    }
                } elseif ($item.PSIsContainer) {
                    # Recursively process subfolders
                    Get-IEFavorites -folderPath $item.FullName -currentFolder "$currentFolder/$($item.Name)"
                }
            }
        }
        
        Get-IEFavorites -folderPath $ieFavoritesPath -currentFolder "Favorites"
        
        if ($favorites.Count -gt 0) {
            # Use string literal "UTF8" for PowerShell 5.1 compatibility
            $favorites | Export-Csv -Path $OutputCSV -NoTypeInformation -Encoding "UTF8" -ErrorAction Stop
            Write-Log "Successfully exported $($favorites.Count) IE Favorites to: $OutputCSV"
            return $true
        } else {
            Write-Log "No IE Favorites found" "WARNING"
            return $false
        }
    }
    catch {
        Write-Log "Error exporting IE Favorites: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# Main execution (starts after log initialization)
$installedBrowsers = Get-InstalledBrowsers
Write-Log "Detected browsers: $($installedBrowsers -join ', ')"

$exportResults = @()

foreach ($browser in $installedBrowsers) {
    # Define a single output file path for the browser
    $outputFile = Join-Path $Script:OutputRootPath "$($browser)_Bookmarks_$(Get-Date -Format 'yyyyMMdd').csv"
    $allSuccess = $false # Tracks if at least one profile/export succeeded

    switch ($browser) {
        "Chrome" {
            # Use cached AppData path for stability (v2.6 fix)
            $browserRoot = Join-Path $Script:LocalAppDataPath "Google\Chrome\User Data"
            
            if (Test-Path $browserRoot) {
                $profileFolders = Get-ChildItem $browserRoot -Directory -ErrorAction SilentlyContinue | Where-Object { 
                    $_.Name -match "^(Default|Profile \d+)$"
                }

                foreach ($profile in $profileFolders) {
                    $bookmarksPath = Join-Path $profile.FullName "Bookmarks"
                    $profileBrowserName = "Chrome ($($profile.Name))"
                    
                    if (Export-ChromeBookmarks -BrowserName $profileBrowserName -BookmarksPath $bookmarksPath -OutputCSV $outputFile) {
                        $allSuccess = $true
                    }
                }
            } else {
                Write-Log "Chrome User Data path not found: $browserRoot" "WARNING"
            }
            
            $exportResults += [PSCustomObject]@{
                Browser = "Chrome"
                Success = $allSuccess
                File = $outputFile
            }
        }
        
        "Edge" {
            # Use cached AppData path for stability (v2.6 fix)
            $browserRoot = Join-Path $Script:LocalAppDataPath "Microsoft\Edge\User Data"
            
            if (Test-Path $browserRoot) {
                $profileFolders = Get-ChildItem $browserRoot -Directory -ErrorAction SilentlyContinue | Where-Object { 
                    $_.Name -match "^(Default|Profile \d+|Guest Profile)$"
                }

                foreach ($profile in $profileFolders) {
                    $bookmarksPath = Join-Path $profile.FullName "Bookmarks"
                    $profileBrowserName = "Edge ($($profile.Name))"
                    
                    if (Export-ChromeBookmarks -BrowserName $profileBrowserName -BookmarksPath $bookmarksPath -OutputCSV $outputFile) {
                        $allSuccess = $true
                    }
                }
            } else {
                Write-Log "Edge User Data path not found: $browserRoot" "WARNING"
            }
            
            $exportResults += [PSCustomObject]@{
                Browser = "Edge"
                Success = $allSuccess
                File = $outputFile
            }
        }
        
        "Firefox" {
            $outputFile = Join-Path $Script:OutputRootPath "Firefox_Bookmarks_$(Get-Date -Format 'yyyyMMdd').csv"
            $success = Export-FirefoxBookmarks -OutputCSV $outputFile
            $exportResults += [PSCustomObject]@{
                Browser = "Firefox"
                Success = $success
                File = $outputFile
            }
        }
        
        "Internet Explorer" {
            $outputFile = Join-Path $Script:OutputRootPath "InternetExplorer_Bookmarks_$(Get-Date -Format 'yyyyMMdd').csv"
            $success = Export-IEBookmarks -OutputCSV $outputFile
            $exportResults += [PSCustomObject]@{
                Browser = "Internet Explorer"
                Success = $success
                File = $outputFile
            }
        }
    }
}

# Summary report
Write-Log "=== EXPORT SUMMARY ==="
$successfulExports = $exportResults | Where-Object { $_.Success }
$failedExports = $exportResults | Where-Object { -not $_.Success }

Write-Log "Successful exports: $($successfulExports.Count)"
foreach ($export in $successfulExports) {
    Write-Log "  - $($export.Browser): $($export.File)"
}

if ($failedExports.Count -gt 0) {
    Write-Log "Failed exports: $($failedExports.Count)" "WARNING"
    foreach ($export in $failedExports) {
        Write-Log "  - $($export.Browser)" "WARNING"
    }
}

Write-Log "Bookmark export process completed"
Write-Log "Log file: $LogFile"

# Display final message to user
if ($successfulExports.Count -gt 0) {
    Write-Host "`nBookmarks exported successfully! 👍" -ForegroundColor Green
    Write-Host "Output location: $($Script:OutputRootPath)" -ForegroundColor Yellow
    Write-Host "Files created:" -ForegroundColor Cyan
    foreach ($export in $successfulExports) {
        Write-Host "  - $($export.File)" -ForegroundColor White
    }
} else {
    Write-Host "`nNo bookmarks were exported. Check log file for details. 😞" -ForegroundColor Red
    Write-Host "Log file: $LogFile" -ForegroundColor Yellow
}

Write-Host "`nNote: To change the output location, modify the `$OutputRootPath variable in the script." -ForegroundColor Gray