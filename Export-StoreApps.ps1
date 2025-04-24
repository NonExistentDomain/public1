<#
.SYNOPSIS
    Exports Microsoft Store applications from a Windows 11 computer for offline installation.
.DESCRIPTION
    This script allows administrators to extract installed Microsoft Store applications
    from a Windows 11 computer. It modifies permissions on the WindowsApps folder,
    copies the specified application packages to a destination folder, and restores
    the original permissions when complete.
.PARAMETER AppNames
    Names of the applications to export (e.g., Microsoft.WindowsCalculator)
.PARAMETER DestinationPath
    Path where the applications will be saved
.EXAMPLE
    .\Export-StoreApps.ps1 -AppNames "Microsoft.WindowsCalculator","Microsoft.ScreenSketch" -DestinationPath "E:\StoreApps"
.NOTES
    Requires administrative privileges
#>

param (
    [Parameter(Mandatory=$true)]
    [string[]]$AppNames,
    
    [Parameter(Mandatory=$true)]
    [string]$DestinationPath
)

# Function to handle errors
function Write-ErrorLog {
    param (
        [string]$Message
    )
    Write-Host "ERROR: $Message" -ForegroundColor Red
}

# Function to handle information messages
function Write-InfoLog {
    param (
        [string]$Message
    )
    Write-Host "INFO: $Message" -ForegroundColor Cyan
}

# Check for administrative privileges
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-ErrorLog "This script requires administrative privileges. Please restart PowerShell as an Administrator."
    exit 1
}

# Create the destination directory if it doesn't exist
if (-not (Test-Path -Path $DestinationPath)) {
    try {
        New-Item -ItemType Directory -Path $DestinationPath -Force | Out-Null
        Write-InfoLog "Created destination directory: $DestinationPath"
    }
    catch {
        Write-ErrorLog "Failed to create destination directory: $_"
        exit 1
    }
}

# Define the WindowsApps folder path
$windowsAppsPath = "C:\Program Files\WindowsApps"

# Store the original ACL for restoration later
$originalAcl = Get-Acl -Path $windowsAppsPath

try {
    Write-InfoLog "Modifying permissions on $windowsAppsPath..."
    
    # Take ownership of the WindowsApps directory
    $takeown = Start-Process -FilePath "takeown.exe" -ArgumentList "/f `"$windowsAppsPath`" /r /d y" -PassThru -Wait -NoNewWindow
    if ($takeown.ExitCode -ne 0) {
        throw "Failed to take ownership of $windowsAppsPath"
    }
    
    # Grant the current user full control permissions
    $icacls = Start-Process -FilePath "icacls.exe" -ArgumentList "`"$windowsAppsPath`" /grant `"$($env:USERNAME)`":F /t" -PassThru -Wait -NoNewWindow
    if ($icacls.ExitCode -ne 0) {
        throw "Failed to grant permissions on $windowsAppsPath"
    }
    
    Write-InfoLog "Permissions modified successfully."
    
    # Process each specified app
    foreach ($appName in $AppNames) {
        Write-InfoLog "Processing $appName..."
        
        # Find all matching app folders (including version variants)
        $appFolders = Get-ChildItem -Path $windowsAppsPath -Directory | Where-Object { $_.Name -like "$appName*" }
        
        if ($appFolders.Count -eq 0) {
            Write-ErrorLog "No installation found for $appName"
            continue
        }
        
        # Find the latest version of the app (assuming version is in the folder name)
        $latestAppFolder = $appFolders | Sort-Object Name -Descending | Select-Object -First 1
        $appSourcePath = $latestAppFolder.FullName
        $appDestPath = Join-Path -Path $DestinationPath -ChildPath $latestAppFolder.Name
        
        Write-InfoLog "Copying $($latestAppFolder.Name) to $appDestPath..."
        
        try {
            # Create app destination folder
            New-Item -ItemType Directory -Path $appDestPath -Force | Out-Null
            
            # Copy app files (excluding any access-denied items)
            Copy-Item -Path "$appSourcePath\*" -Destination $appDestPath -Recurse -Force -ErrorAction SilentlyContinue
            
            # Verify the AppxManifest.xml file was copied
            $manifestPath = Join-Path -Path $appDestPath -ChildPath "AppxManifest.xml"
            if (-not (Test-Path -Path $manifestPath)) {
                Write-ErrorLog "Failed to copy AppxManifest.xml for $appName"
            } else {
                Write-InfoLog "Successfully exported $appName to $appDestPath"
            }
        }
        catch {
            Write-ErrorLog "Failed to copy $appName: $_"
        }
    }
}
catch {
    Write-ErrorLog "An error occurred: $_"
}
finally {
    # Restore the original permissions
    Write-InfoLog "Restoring original permissions on $windowsAppsPath..."
    try {
        Set-Acl -Path $windowsAppsPath -AclObject $originalAcl
        Write-InfoLog "Original permissions restored successfully."
    }
    catch {
        Write-ErrorLog "Failed to restore original permissions: $_"
    }
}

Write-InfoLog "App export operation completed. Applications saved to $DestinationPath"