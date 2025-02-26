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

# Function to check if a path is contained within another path
function Test-IsChildPath {
  param(
    [string]$ChildPath,
    [string]$ParentPath
  )
  
  # Ensure paths end with backslash for proper comparison
  if (-not $ChildPath.EndsWith('\')) {
    $ChildPath = $ChildPath + '\'
  }
  
  if (-not $ParentPath.EndsWith('\')) {
    $ParentPath = $ParentPath + '\'
  }
  
  return $ChildPath.StartsWith($ParentPath, [StringComparison]::OrdinalIgnoreCase) -and $ChildPath -ne $ParentPath
}

# Function to check if a path is parent of another path
function Test-IsParentPath {
  param(
    [string]$PotentialParentPath,
    [string]$PotentialChildPath
  )
  
  return (Test-IsChildPath -ChildPath $PotentialChildPath -ParentPath $PotentialParentPath)
}

# Function to get scan type (directories or files)
function Get-ScanType {
  Write-Host "`nScan Type Options:" -ForegroundColor Cyan
  Write-Host "1. Scan for largest files"
  Write-Host "2. Scan for largest directories"
  
  do {
    $choice = Read-Host "`nSelect an option (1-2)"
  } while ($choice -notmatch '^[1-2]$')
  
  return [int]$choice
}

# Function to get age filter parameters
function Get-AgeFilterParams {
  Write-Host "`nFile Age Filter Options:" -ForegroundColor Cyan
  Write-Host "1. Scan all items (no age filter)"
  Write-Host "2. Scan items older than a specific date (MM/DD/YYYY)"
  Write-Host "3. Scan items older than X days"
  Write-Host "4. Scan items older than X years"
  
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
      return (Get-Date).AddDays( - [int]$days)
    }
    '4' {
      do {
        $years = Read-Host "Enter number of years"
      } while (-not ([int]::TryParse($years, [ref]$null)))
      return (Get-Date).AddYears( - [int]$years)
    }
  }
}

# Function to scan for largest files
function Get-LargestFiles {
  param(
    [string]$Path,
    [nullable[datetime]]$DateFilter = $null,
    [int]$ResultCount
  )
  
  $filesList = [System.Collections.ArrayList]::new()
  $totalItems = 0
  $accessDeniedPaths = @()
  
  Write-Host "Scanning for largest files..." -ForegroundColor Green
  
  Get-ChildItem -Path $Path -Recurse -File -ErrorAction SilentlyContinue -ErrorVariable accessErrors | ForEach-Object {
    $totalItems++
    Write-Progress -Activity "Scanning directory for files" -Status "Files processed: $totalItems" -PercentComplete -1
      
    try {
      # Skip files that don't meet the age filter
      if ($null -ne $DateFilter -and $_.LastWriteTime -gt $DateFilter) {
        return
      }
          
      [void]$filesList.Add([PSCustomObject]@{
          Path           = $_.FullName
          SizeBytes      = $_.Length  # Used for sorting only
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
  $selectedItems = $filesList | Sort-Object SizeBytes -Descending | 
  Select-Object -First $ResultCount |
  Select-Object Path, SizeReadable, CreationTime, LastAccessTime, LastWriteTime
  
  return @{
    SelectedItems     = $selectedItems
    TotalItems        = $totalItems
    AccessDeniedPaths = $accessDeniedPaths
  }
}

# Function to scan for largest directories
function Get-LargestDirectories {
  param(
    [string]$Path,
    [nullable[datetime]]$DateFilter = $null,
    [int]$ResultCount
  )
  
  $dirsList = [System.Collections.ArrayList]::new()
  $totalItems = 0
  $accessDeniedPaths = @()
  
  Write-Host "Scanning for largest directories..." -ForegroundColor Green
  
  # Get all directories
  $directories = Get-ChildItem -Path $Path -Recurse -Directory -ErrorAction SilentlyContinue -ErrorVariable accessErrors
  $totalDirs = $directories.Count
  $currentDir = 0
  
  foreach ($dir in $directories) {
    $currentDir++
    $totalItems++
    $percentComplete = [math]::Min(100, [math]::Round(($currentDir / $totalDirs) * 100))
    Write-Progress -Activity "Scanning directories" -Status "Directories processed: $currentDir of $totalDirs" -PercentComplete $percentComplete
      
    try {
      # Skip directories that don't meet the age filter
      if ($null -ne $DateFilter -and $dir.LastWriteTime -gt $DateFilter) {
        continue
      }
          
      # Calculate directory size
      $size = 0
      Get-ChildItem -Path $dir.FullName -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
        $size += $_.Length
      }
          
      [void]$dirsList.Add([PSCustomObject]@{
          Path           = $dir.FullName
          SizeBytes      = $size  # Used for sorting only
          SizeReadable   = Convert-Size $size
          CreationTime   = $dir.CreationTime.ToString("yyyy-MM-dd HH:mm:ss")
          LastAccessTime = $dir.LastAccessTime.ToString("yyyy-MM-dd HH:mm:ss")
          LastWriteTime  = $dir.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
        })
    }
    catch {
      Write-Warning "Error processing directory $($dir.FullName): $($_.Exception.Message)"
    }
  }
  
  # Process access errors
  foreach ($error in $accessErrors) {
    if ($error.Exception -is [System.UnauthorizedAccessException]) {
      $accessDeniedPaths += $error.TargetObject
    }
  }
  
  # Sort directories by size descending
  $sortedDirs = $dirsList | Sort-Object SizeBytes -Descending
  
  # New approach: Select top directories while filtering out parent directories
  $finalDirs = [System.Collections.ArrayList]::new()
  
  # First get the top N directories by size
  $topDirs = $sortedDirs | Select-Object -First ($ResultCount * 2)  # Get more than needed for filtering
  
  # Now filter out parent directories of already included directories
  foreach ($dir in $topDirs) {
    $shouldInclude = $true
      
    # Check if this directory is a parent of any already included directory
    foreach ($included in $finalDirs) {
      if (Test-IsParentPath -PotentialParentPath $dir.Path -PotentialChildPath $included.Path) {
        $shouldInclude = $false
        break
      }
    }
      
    # If it's not a parent of any included dir, also check if it's a child of any included dir to avoid duplicates
    if ($shouldInclude) {
      foreach ($included in $finalDirs) {
        if (Test-IsParentPath -PotentialParentPath $included.Path -PotentialChildPath $dir.Path) {
          $shouldInclude = $false
          break
        }
      }
    }
      
    # Include it if it passed all checks
    if ($shouldInclude) {
      [void]$finalDirs.Add($dir)
          
      # Break if we've reached our target count
      if ($finalDirs.Count -ge $ResultCount) {
        break
      }
    }
  }
  
  # Sort the final list by size again and select only the required properties
  $selectedItems = $finalDirs | Sort-Object SizeBytes -Descending | 
  Select-Object Path, SizeReadable, CreationTime, LastAccessTime, LastWriteTime
  
  return @{
    SelectedItems     = $selectedItems
    TotalItems        = $totalItems
    AccessDeniedPaths = $accessDeniedPaths
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
      
    # Get scan type
    $scanType = Get-ScanType
      
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
      $itemType = if ($scanType -eq 1) { "files" } else { "directories" }
      $resultCount = Read-Host "Enter the number of largest $itemType to display and export (e.g., 100)"
    } while (-not ($resultCount -match '^\d+$'))
      
    # Create output filename
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $dirName = Get-SanitizedDirectoryName $sourcePath
    $scanTypeStr = if ($scanType -eq 1) { "Files" } else { "Directories" }
    $csvFileName = "{0}_{1}_{2}.csv" -f $dirName, $scanTypeStr, $timestamp
    $csvPath = Join-Path $outputPath $csvFileName

    # Create access denied log filename
    $logFileName = "AccessDenied_{0}_{1}.txt" -f $dirName, $timestamp
    $logFilePath = Join-Path $defaultOutput $logFileName
      
    # Perform the scan based on scan type
    if ($scanType -eq 1) {
      $result = Get-LargestFiles -Path $sourcePath -DateFilter $dateFilter -ResultCount $resultCount
    }
    else {
      $result = Get-LargestDirectories -Path $sourcePath -DateFilter $dateFilter -ResultCount $resultCount
    }
      
    $selectedItems = $result.SelectedItems
    $totalItems = $result.TotalItems
    $accessDeniedPaths = $result.AccessDeniedPaths
      
    # Export to CSV
    $selectedItems | Export-Csv -Path $csvPath -NoTypeInformation

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
    $itemTypeStr = if ($scanType -eq 1) { "files" } else { "directories" }
    Write-Host "- Total $itemTypeStr scanned: $totalItems"
    Write-Host "- $($itemTypeStr.Substring(0,1).ToUpper() + $itemTypeStr.Substring(1)) exported: $($selectedItems.Count)"
    Write-Host "- Report saved to: $csvPath"
    Write-Host "- Execution time: $($duration.TotalSeconds.ToString('F2')) seconds"

    if ($accessDeniedPaths.Count -gt 0) {
      Write-Host "`nWARNING: Some locations could not be accessed due to permissions" -ForegroundColor Yellow
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

    # Prompt for another scan with type selection
    Write-Host "`nWhat would you like to do next?" -ForegroundColor Cyan
    Write-Host "1. Start a new scan for files"
    Write-Host "2. Start a new scan for directories" 
    Write-Host "3. Exit the program"
      
    do {
      $nextAction = Read-Host "`nSelect an option (1-3)"
    } while ($nextAction -notmatch '^[1-3]$')
      
    if ($nextAction -eq '1' -or $nextAction -eq '2') {
      # Store the selected scan type from the menu (1=files, 2=directories)
      $SCRIPT:PreSelectedScanType = [int]$nextAction
          
      Write-Host "`nStarting new scan...`n" -ForegroundColor Green
      Start-DirectoryScanner
      return
    }
    else {
      Write-Host "`nThank you for using the Directory Scanner Tool. Goodbye!" -ForegroundColor Green
    }
  }
  catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
  }
  finally {
    Write-Progress -Activity "Scanning" -Completed
  }
}

# Add pre-selection support to Get-ScanType
function Get-ScanType {
  # Check if we have a pre-selected scan type
  if ($null -ne $SCRIPT:PreSelectedScanType) {
    $choice = $SCRIPT:PreSelectedScanType
    # Clear it for future scans
    $SCRIPT:PreSelectedScanType = $null
    return $choice
  }
  
  # Normal selection if no pre-selection
  Write-Host "`nScan Type Options:" -ForegroundColor Cyan
  Write-Host "1. Scan for largest files"
  Write-Host "2. Scan for largest directories"
  
  do {
    $choice = Read-Host "`nSelect an option (1-2)"
  } while ($choice -notmatch '^[1-2]$')
  
  return [int]$choice
}

# Start the scanner
# Initialize the script-scope variable for scan type pre-selection
$SCRIPT:PreSelectedScanType = $null
Start-DirectoryScanner