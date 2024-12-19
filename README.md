# PowerShell Large File Scanner

A powerful and user-friendly PowerShell script for scanning directories and identifying large files with advanced filtering capabilities.

## Features

- **Large File Detection**: Identifies and reports the largest files in specified directories
- **Age-Based Filtering Options**:
  - No age filter (scan all files)
  - Files older than a specific date (MM/DD/YYYY)
  - Files older than X days
  - Files older than X years
- **Human-Readable Output**: Automatically converts file sizes to human-readable format (Bytes, KB, MB, GB, TB)
- **Detailed Timestamps**: Reports creation time, last access time, and last write time for each file
- **Access Denial Logging**: Tracks and reports directories that couldn't be accessed due to permissions
- **CSV Export**: Exports results in CSV format for easy analysis
- **Progress Tracking**: Real-time progress monitoring during scanning
- **Flexible Output Location**: Customizable output directory with desktop as default

## Usage

1. Run the script in PowerShell
2. Enter the source directory path to scan
3. Select an age filter option (if desired)
4. Specify the number of largest files to report
5. Choose output location (optional)

## Output Format

The script generates two potential files:
1. **Main Report (CSV)**:
   - File Path
   - Size (Human-Readable)
   - Creation Time
   - Last Access Time
   - Last Write Time

2. **Access Denied Log (TXT)** (if applicable):
   - List of paths that couldn't be accessed due to permissions

## System Requirements

- This script was developed and tested on PowerShell Core 7. It should work with Windows PowerShell 5.1 or later but some features might not work as intended. Features such as the naming of the output file or opening the wrong output file directory when requested to do so after the scan has completed.
- Administrative privileges recommended for full access to system directories

## Best Practices

- Run as Administrator for complete system access
- Use full paths when specifying directories
- Consider filtering options for large directory structures

## Error Handling

The script includes comprehensive error handling for:
- Invalid paths
- Permission issues
- Date format validation
- File access errors

## Performance Considerations

- Uses ArrayList for efficient memory management
- Implements progress tracking for large directories
- Optimized sorting for large file collections

## Sample Output

=== Directory Scanner Tool ===
Summary:
- Total files scanned: 1000
- Files exported: 100
- Report saved to: C:\Users\Username\Desktop\ScanResults_2024-01-01_12-00-00.csv
- Execution time: 5.32 seconds

## License

MIT License


