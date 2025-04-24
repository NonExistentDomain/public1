param (
    [Parameter(Mandatory=$true)]
    [string[]]$ComputerNames,
    [string]$OutputCsvPath = "report.csv"
)

# Initialize an array to store report data
$report = @()

# Process each computer
foreach ($computer in $ComputerNames) {
    try {
        # Execute script block on the remote computer
        $results = Invoke-Command -ComputerName $computer -ScriptBlock {
            # Get all file system drives
            $drives = Get-PSDrive -PSProvider FileSystem
            $files = @()

            # Search each drive for target files
            foreach ($drive in $drives) {
                $files += Get-ChildItem -Path $drive.Root -Recurse -File -ErrorAction SilentlyContinue |
                          Where-Object { $_.Name -match '\.abc$' -or $_.Name -match '\.cft$' -or 
                                         $_.Name -match '\.abc\.old$' -or $_.Name -match '\.cft\.old$' }
            }

            # Process each file found
            foreach ($file in $files) {
                $status = $null

                if ($file.Name -match '\.abc\.old$' -or $file.Name -match '\.cft\.old$') {
                    # File is already renamed
                    $status = "Already Renamed"
                }
                elseif ($file.Name -match '\.abc$' -or $file.Name -match '\.cft$') {
                    # File needs renaming
                    $newName = $file.FullName + ".old"
                    try {
                        Rename-Item -Path $file.FullName -NewName $newName -ErrorAction Stop
                        $status = "Renamed"
                    }
                    catch {
                        $status = "Failed to Rename"
                    }
                }

                # Create report entry
                [PSCustomObject]@{
                    ComputerName = $env:COMPUTERNAME
                    FilePath     = $file.FullName
                    Status       = $status
                }
            }
        } -ErrorAction Stop

        # Add results to report
        $report += $results
    }
    catch {
        # Record connection failure
        $report += [PSCustomObject]@{
            ComputerName = $computer
            FilePath     = $null
            Status       = "Failed to Connect"
        }
    }
}

# Export the report to CSV
$report | Export-Csv -Path $OutputCsvPath -NoTypeInformation


#------------------------------------------------------------------------------


<#
.SYNOPSIS
    Finds and renames files with .abc or .cft extensions to .abc.old or .cft.old on remote computers.

.DESCRIPTION
    This script scans all drives on specified remote computers for files with .abc or .cft extensions,
    renames them to add .old to their existing extension, and generates a comprehensive CSV report.

.PARAMETER ComputerList
    Path to a text file containing the list of computer names, one per line.

.PARAMETER OutputCSV
    Path where the CSV report will be saved.

.EXAMPLE
    .\Find-RenameExtensions.ps1 -ComputerList "C:\Temp\computers.txt" -OutputCSV "C:\Temp\RenameReport.csv"
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$ComputerList,
    
    [Parameter(Mandatory = $true)]
    [string]$OutputCSV
)

# Initialize results collection
$results = @()

# Check if ComputerList file exists
if (-not (Test-Path -Path $ComputerList)) {
    Write-Error "Computer list file not found: $ComputerList"
    exit 1
}

# Read computer names from file
$computers = Get-Content -Path $ComputerList

# Create CSV directory if it doesn't exist
$csvDirectory = Split-Path -Path $OutputCSV -Parent
if ($csvDirectory -and -not (Test-Path -Path $csvDirectory)) {
    New-Item -Path $csvDirectory -ItemType Directory -Force | Out-Null
}

Write-Host "Starting file scan and rename operation across $(($computers | Measure-Object).Count) computers..."

# Process each computer
foreach ($computer in $computers) {
    $computer = $computer.Trim()
    if (-not $computer) { continue }
    
    Write-Host "Processing computer: $computer" -ForegroundColor Cyan
    
    # Check if computer is online
    if (-not (Test-Connection -ComputerName $computer -Count 1 -Quiet)) {
        Write-Host "  Computer is offline: $computer" -ForegroundColor Yellow
        
        # Add offline computer to results
        $results += [PSCustomObject]@{
            ComputerName = $computer
            Status = "Offline"
            FilePath = $null
            OriginalName = $null
            NewName = $null
            Action = "Connection Failed"
            ErrorMessage = "Computer is offline or unreachable"
        }
        continue
    }
    
    try {
        # Create a script block to run on the remote computer
        $scriptBlock = {
            param($extensionsToFind)
            
            # Get all drives excluding network drives
            $drives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Root -match '^[A-Z]:' }
            
            # Initialize results
            $fileResults = @()
            
            foreach ($drive in $drives) {
                Write-Output "  Scanning drive $($drive.Root) on $env:COMPUTERNAME..."
                
                try {
                    # Find files with specified extensions
                    $foundFiles = @()
                    
                    # Find .abc files
                    Write-Output "    Searching for .abc files..."
                    $abcFiles = Get-ChildItem -Path $drive.Root -Filter "*.abc" -File -Recurse -ErrorAction SilentlyContinue -ErrorVariable getErr
                    $foundFiles += $abcFiles
                    
                    # Find .cft files
                    Write-Output "    Searching for .cft files..."
                    $cftFiles = Get-ChildItem -Path $drive.Root -Filter "*.cft" -File -Recurse -ErrorAction SilentlyContinue -ErrorVariable +getErr
                    $foundFiles += $cftFiles
                    
                    # Find already renamed files
                    Write-Output "    Searching for .abc.old files..."
                    $abcOldFiles = Get-ChildItem -Path $drive.Root -Filter "*.abc.old" -File -Recurse -ErrorAction SilentlyContinue -ErrorVariable +getErr
                    
                    Write-Output "    Searching for .cft.old files..."
                    $cftOldFiles = Get-ChildItem -Path $drive.Root -Filter "*.cft.old" -File -Recurse -ErrorAction SilentlyContinue -ErrorVariable +getErr
                    
                    # Process found files
                    foreach ($file in $foundFiles) {
                        $newName = "$($file.FullName).old"
                        
                        try {
                            # Check if file is not already renamed
                            if (-not ($file.Name -like "*.old")) {
                                # Attempt to rename the file
                                Rename-Item -Path $file.FullName -NewName $newName -ErrorAction Stop
                                
                                $fileResults += [PSCustomObject]@{
                                    FilePath = $file.FullName
                                    OriginalName = $file.Name
                                    NewName = "$($file.Name).old"
                                    Action = "Renamed"
                                    ErrorMessage = $null
                                }
                            }
                        }
                        catch {
                            $fileResults += [PSCustomObject]@{
                                FilePath = $file.FullName
                                OriginalName = $file.Name
                                NewName = "$($file.Name).old"
                                Action = "Failed to Rename"
                                ErrorMessage = $_.Exception.Message
                            }
                        }
                    }
                    
                    # Process already renamed files
                    foreach ($file in $abcOldFiles + $cftOldFiles) {
                        $fileResults += [PSCustomObject]@{
                            FilePath = $file.FullName
                            OriginalName = $file.Name
                            NewName = $file.Name
                            Action = "Already Renamed"
                            ErrorMessage = $null
                        }
                    }
                    
                    # Process any access errors during file search
                    foreach ($err in $getErr) {
                        $fileResults += [PSCustomObject]@{
                            FilePath = $err.TargetObject
                            OriginalName = $null
                            NewName = $null
                            Action = "Access Error"
                            ErrorMessage = $err.Exception.Message
                        }
                    }
                }
                catch {
                    $fileResults += [PSCustomObject]@{
                        FilePath = $drive.Root
                        OriginalName = $null
                        NewName = $null
                        Action = "Drive Access Error"
                        ErrorMessage = $_.Exception.Message
                    }
                }
            }
            
            return $fileResults
        }
        
        # Execute the script block on the remote computer
        $remoteResults = Invoke-Command -ComputerName $computer -ScriptBlock $scriptBlock -ArgumentList @(".abc", ".cft") -ErrorAction Stop
        
        # Process the results from the remote computer
        if ($remoteResults.Count -eq 0) {
            Write-Host "  No matching files found on $computer" -ForegroundColor Green
            
            # Add entry for computer with no files
            $results += [PSCustomObject]@{
                ComputerName = $computer
                Status = "Online"
                FilePath = $null
                OriginalName = $null
                NewName = $null
                Action = "No Files Found"
                ErrorMessage = $null
            }
        }
        else {
            Write-Host "  Found $($remoteResults.Count) files on $computer" -ForegroundColor Green
            
            # Add all remote results to the main results collection
            foreach ($item in $remoteResults) {
                $results += [PSCustomObject]@{
                    ComputerName = $computer
                    Status = "Online"
                    FilePath = $item.FilePath
                    OriginalName = $item.OriginalName
                    NewName = $item.NewName
                    Action = $item.Action
                    ErrorMessage = $item.ErrorMessage
                }
            }
        }
    }
    catch {
        Write-Host "  Error connecting to $computer: $($_.Exception.Message)" -ForegroundColor Red
        
        # Add error entry
        $results += [PSCustomObject]@{
            ComputerName = $computer
            Status = "Error"
            FilePath = $null
            OriginalName = $null
            NewName = $null
            Action = "Remote Execution Failed"
            ErrorMessage = $_.Exception.Message
        }
    }
}

# Export results to CSV
try {
    $results | Export-Csv -Path $OutputCSV -NoTypeInformation -Force
    Write-Host "Results exported to: $OutputCSV" -ForegroundColor Green
}
catch {
    Write-Host "Error exporting results to CSV: $($_.Exception.Message)" -ForegroundColor Red
}

# Display summary
$summary = $results | Group-Object Action | Select-Object Name, Count
Write-Host "`nSummary:" -ForegroundColor Cyan
$summary | Format-Table -AutoSize

Write-Host "Total computers processed: $(($computers | Measure-Object).Count)" -ForegroundColor Cyan
Write-Host "Total files processed: $(($results | Where-Object { $_.FilePath -ne $null } | Measure-Object).Count)" -ForegroundColor Cyan



#--------------------------------------------------------------------------

<#
.SYNOPSIS
    Finds and renames files with .abc or .cft extensions to .abc.old or .cft.old on the local computer.

.DESCRIPTION
    This script scans all drives on the local computer for files with .abc or .cft extensions,
    renames them to add .old to their existing extension, and generates a comprehensive CSV report.

.PARAMETER OutputCSV
    Path where the CSV report will be saved.

.EXAMPLE
    .\Find-RenameExtensionsLocal.ps1 -OutputCSV "C:\Temp\RenameReport.csv"
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$OutputCSV
)

# Initialize results collection
$results = @()
$computerName = $env:COMPUTERNAME

# Create CSV directory if it doesn't exist
$csvDirectory = Split-Path -Path $OutputCSV -Parent
if ($csvDirectory -and -not (Test-Path -Path $csvDirectory)) {
    New-Item -Path $csvDirectory -ItemType Directory -Force | Out-Null
}

Write-Host "Starting file scan and rename operation on local computer: $computerName..."

# Get all drives excluding network drives
$drives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Root -match '^[A-Z]:' }

foreach ($drive in $drives) {
    Write-Host "Scanning drive $($drive.Root)..." -ForegroundColor Cyan
    
    try {
        # Find files with specified extensions
        $foundFiles = @()
        
        # Find .abc files
        Write-Host "  Searching for .abc files..."
        $abcFiles = Get-ChildItem -Path $drive.Root -Filter "*.abc" -File -Recurse -ErrorAction SilentlyContinue -ErrorVariable getErr
        $foundFiles += $abcFiles
        
        # Find .cft files
        Write-Host "  Searching for .cft files..."
        $cftFiles = Get-ChildItem -Path $drive.Root -Filter "*.cft" -File -Recurse -ErrorAction SilentlyContinue -ErrorVariable +getErr
        $foundFiles += $cftFiles
        
        # Find already renamed files
        Write-Host "  Searching for .abc.old files..."
        $abcOldFiles = Get-ChildItem -Path $drive.Root -Filter "*.abc.old" -File -Recurse -ErrorAction SilentlyContinue -ErrorVariable +getErr
        
        Write-Host "  Searching for .cft.old files..."
        $cftOldFiles = Get-ChildItem -Path $drive.Root -Filter "*.cft.old" -File -Recurse -ErrorAction SilentlyContinue -ErrorVariable +getErr
        
        # Process found files
        foreach ($file in $foundFiles) {
            $newName = "$($file.FullName).old"
            
            try {
                # Check if file is not already renamed
                if (-not ($file.Name -like "*.old")) {
                    # Attempt to rename the file
                    Rename-Item -Path $file.FullName -NewName $newName -ErrorAction Stop
                    
                    $results += [PSCustomObject]@{
                        ComputerName = $computerName
                        DriveLabel = $drive.Name
                        FilePath = $file.FullName
                        OriginalName = $file.Name
                        NewName = "$($file.Name).old"
                        Action = "Renamed"
                        ErrorMessage = $null
                    }
                    
                    Write-Host "  Renamed: $($file.FullName) -> $newName" -ForegroundColor Green
                }
            }
            catch {
                $results += [PSCustomObject]@{
                    ComputerName = $computerName
                    DriveLabel = $drive.Name
                    FilePath = $file.FullName
                    OriginalName = $file.Name
                    NewName = "$($file.Name).old"
                    Action = "Failed to Rename"
                    ErrorMessage = $_.Exception.Message
                }
                
                Write-Host "  Failed to rename: $($file.FullName) - $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        
        # Process already renamed files
        foreach ($file in $abcOldFiles + $cftOldFiles) {
            $results += [PSCustomObject]@{
                ComputerName = $computerName
                DriveLabel = $drive.Name
                FilePath = $file.FullName
                OriginalName = $file.Name
                NewName = $file.Name
                Action = "Already Renamed"
                ErrorMessage = $null
            }
            
            Write-Host "  Already renamed: $($file.FullName)" -ForegroundColor Yellow
        }
        
        # Process any access errors during file search
        foreach ($err in $getErr) {
            $results += [PSCustomObject]@{
                ComputerName = $computerName
                DriveLabel = $drive.Name
                FilePath = $err.TargetObject
                OriginalName = $null
                NewName = $null
                Action = "Access Error"
                ErrorMessage = $err.Exception.Message
            }
            
            Write-Host "  Access error: $($err.TargetObject) - $($err.Exception.Message)" -ForegroundColor Red
        }
    }
    catch {
        $results += [PSCustomObject]@{
            ComputerName = $computerName
            DriveLabel = $drive.Name
            FilePath = $drive.Root
            OriginalName = $null
            NewName = $null
            Action = "Drive Access Error"
            ErrorMessage = $_.Exception.Message
        }
        
        Write-Host "Error accessing drive $($drive.Root): $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Add summary entry if no files were found
if ($results.Count -eq 0) {
    $results += [PSCustomObject]@{
        ComputerName = $computerName
        DriveLabel = $null
        FilePath = $null
        OriginalName = $null
        NewName = $null
        Action = "No Files Found"
        ErrorMessage = $null
    }
    
    Write-Host "No matching files found on any drives" -ForegroundColor Yellow
}

# Export results to CSV
try {
    $results | Export-Csv -Path $OutputCSV -NoTypeInformation -Force
    Write-Host "Results exported to: $OutputCSV" -ForegroundColor Green
}
catch {
    Write-Host "Error exporting results to CSV: $($_.Exception.Message)" -ForegroundColor Red
}

# Display summary
$summary = $results | Group-Object Action | Select-Object Name, Count
Write-Host "`nSummary:" -ForegroundColor Cyan
$summary | Format-Table -AutoSize

Write-Host "Total files processed: $($results.Count)" -ForegroundColor Cyan

# Optional: Open the CSV file
$openCsv = Read-Host "Do you want to open the CSV file now? (Y/N)"
if ($openCsv -eq "Y" -or $openCsv -eq "y") {
    Start-Process $OutputCSV
}
