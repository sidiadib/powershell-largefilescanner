# Function to convert bytes to human readable format
function Convert-Size {
    param([long]$Size)
    $sizes = 'Bytes,KB,MB,GB,TB'
    $sizes = $sizes.Split(',')
    $index = 0
    while ($Size -ge 1024 -and $index -lt ($sizes.Count - 1)) {
        $Size = $Size / 1024
        $index++
    }
    return "{0:N2} {1}" -f $Size, $sizes[$index]
}

# Function to validate path
function Test-ValidPath {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $false
    }
    if (-not (Test-Path -Path $Path)) {
        return $false
    }
    return $true
}

# Function to get sanitized directory name
function Get-SanitizedDirectoryName {
    param([string]$Path)
    return (Split-Path $Path -Leaf) -replace '[\\/:*?"<>|]', '_'
}

# Function to get age filter parameters
function Get-AgeFilterParams {
    Write-Host "`nFile Age Filter Options:" -ForegroundColor Cyan
    Write-Host "1. Scan all files (no age filter)"
    Write-Host "2. Scan files older than a specific date (MM/DD/YYYY)"
    Write-Host "3. Scan files older than X days"
    Write-Host "4. Scan files older than X years"
    
    do {
        $choice = Read-Host "`nSelect an option (1-4)"
    } while ($choice -notmatch '^[1-4]$')
    
    switch ($choice) {
        '1' { return $null }
        '2' {
            do {
                $date = Read-Host "Enter date (MM/DD/YYYY)"
                try {
                    $parsedDate = [datetime]::ParseExact($date, "MM/dd/yyyy", $null)
                    return $parsedDate
                }
                catch {
                    Write-Host "Invalid date format. Please use MM/DD/YYYY format" -ForegroundColor Red
                    continue
                }
            } while ($true)
        }
        '3' {
            do {
                $days = Read-Host "Enter number of days"
            } while (-not ([int]::TryParse($days, [ref]$null)))
            return (Get-Date).AddDays(-[int]$days)
        }
        '4' {
            do {
                $years = Read-Host "Enter number of years"
            } while (-not ([int]::TryParse($years, [ref]$null)))
            return (Get-Date).AddYears(-[int]$years)
        }
    }
}

# Main scanning function
function Start-DirectoryScanner {
    try {
        $ErrorActionPreference = 'Stop'
        $start = Get-Date
        
        Write-Host "`n=== Directory Scanner Tool ===" -ForegroundColor Cyan
        Write-Host "NOTE: For best results and to avoid permission issues, please run this script as Administrator" -ForegroundColor Yellow
        Write-Host "============================================================`n" -ForegroundColor Cyan
        
        # Get source path with validation
        do {
            $sourcePath = Read-Host "Enter source directory path"
            if (-not (Test-ValidPath $sourcePath)) {
                Write-Host "Invalid or empty path. Please enter a valid directory path." -ForegroundColor Red
            }
        } while (-not (Test-ValidPath $sourcePath))
        
        # Get age filter
        $dateFilter = Get-AgeFilterParams
        
        # Get output path with Desktop as default
        $defaultOutput = [System.Environment]::GetFolderPath('Desktop')
        $outputPath = Read-Host "Enter output directory path (press Enter for Desktop)"
        
        if ([string]::IsNullOrWhiteSpace($outputPath)) {
            $outputPath = $defaultOutput
            Write-Host "Using default output path: $outputPath" -ForegroundColor Yellow
        }
        else {
            # Validate custom output path
            if (-not (Test-ValidPath $outputPath)) {
                Write-Host "Invalid output path. Using default path: $defaultOutput" -ForegroundColor Yellow
                $outputPath = $defaultOutput
            }
        }
        
        do {
            $resultCount = Read-Host "Enter the number of largest files to display and export (e.g., 100)"
        } while (-not ($resultCount -match '^\d+$'))
        
        # Create output filename
        $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
        $dirName = Get-SanitizedDirectoryName $sourcePath
        $csvFileName = "{0}_{1}.csv" -f $dirName, $timestamp
        $csvPath = Join-Path $outputPath $csvFileName

        # Create access denied log filename
        $logFileName = "AccessDenied_{0}_{1}.txt" -f $dirName, $timestamp
        $logFilePath = Join-Path $defaultOutput $logFileName
        
        # Initialize progress variables
        $activity = "Scanning directory"
        $totalFiles = 0
        $accessDeniedPaths = @()
        
        Write-Host "Starting directory scan." -ForegroundColor Green
        
        # Create an ArrayList to store file information
        $filesList = [System.Collections.ArrayList]::new()
        
        # Scan directory and collect file information
        Get-ChildItem -Path $sourcePath -Recurse -File -ErrorAction SilentlyContinue -ErrorVariable accessErrors | ForEach-Object {
            $totalFiles++
            Write-Progress -Activity $activity -Status "Files processed: $totalFiles" -PercentComplete -1
            
            try {
                # Skip files that don't meet the age filter
                if ($dateFilter -and $_.LastWriteTime -gt $dateFilter) {
                    return
                }
                
                [void]$filesList.Add([PSCustomObject]@{
                    Path           = $_.FullName
                    SizeBytes     = $_.Length  # Used for sorting only
                    SizeReadable   = Convert-Size $_.Length
                    CreationTime   = $_.CreationTime.ToString("yyyy-MM-dd HH:mm:ss")
                    LastAccessTime = $_.LastAccessTime.ToString("yyyy-MM-dd HH:mm:ss")
                    LastWriteTime  = $_.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
                })
            }
            catch {
                Write-Warning "Error processing file $($_.FullName): $($_.Exception.Message)"
            }
        }

        # Process access errors
        foreach ($error in $accessErrors) {
            if ($error.Exception -is [System.UnauthorizedAccessException]) {
                $accessDeniedPaths += $error.TargetObject
            }
        }
        
        # Sort by actual size and select top files, excluding the SizeBytes property
        $selectedFiles = $filesList | Sort-Object SizeBytes -Descending | 
            Select-Object -First $resultCount |
            Select-Object Path, SizeReadable, CreationTime, LastAccessTime, LastWriteTime
        
        # Export to CSV
        $selectedFiles | Export-Csv -Path $csvPath -NoTypeInformation

        # Log access denied paths if any
        if ($accessDeniedPaths.Count -gt 0) {
            "Access Denied Paths - Scan performed on $(Get-Date)`n" | Out-File -FilePath $logFilePath
            "The following paths could not be scanned due to access restrictions:`n" | Out-File -FilePath $logFilePath -Append
            $accessDeniedPaths | ForEach-Object { $_ | Out-File -FilePath $logFilePath -Append }
        }
        
        # Calculate and display summary
        $end = Get-Date
        $duration = $end - $start
        
        Write-Host "`nScan Complete!" -ForegroundColor Green
        Write-Host "Summary:"
        Write-Host "- Total files scanned: $totalFiles"
        Write-Host "- Files exported: $($selectedFiles.Count)"
        Write-Host "- Report saved to: $csvPath"
        Write-Host "- Execution time: $($duration.TotalSeconds.ToString('F2')) seconds"

        if ($accessDeniedPaths.Count -gt 0) {
            Write-Host "`nWARNING: Some directories could not be accessed due to permissions" -ForegroundColor Yellow
            Write-Host "Number of inaccessible paths: $($accessDeniedPaths.Count)" -ForegroundColor Yellow
            Write-Host "Access denied log saved to: $logFilePath" -ForegroundColor Yellow
        }
        
        # Ask user for next action
        if ($accessDeniedPaths.Count -gt 0) {
            $openLog = Read-Host "`nWould you like to view the access denied log? (Y/N) [Y]"
            if ($openLog -eq '' -or $openLog -eq 'Y') {
                Invoke-Item $logFilePath
            }
        }

        $openFolder = Read-Host "`nWould you like to open the output folder? (Y/N) [Y]"
        if ($openFolder -eq '' -or $openFolder -eq 'Y') {
            explorer.exe /select, $csvPath
        }

        $continue = Read-Host "`nWould you like to scan another directory? (Y/N) [N]"
        if ($continue -eq 'Y') {
            Write-Host "`nStarting new scan...`n" -ForegroundColor Green
            Start-DirectoryScanner
            return
        }
        
        Write-Host "`nThank you for using the Directory Scanner Tool. Goodbye!" -ForegroundColor Green
    }
    catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    }
    finally {
        Write-Progress -Activity $activity -Completed
    }
}

# Start the scanner
Start-DirectoryScanner
