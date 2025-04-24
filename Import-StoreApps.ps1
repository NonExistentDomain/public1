<#
.SYNOPSIS
    Installs Microsoft Store applications on an offline Windows 11 computer.
.DESCRIPTION
    This script installs Microsoft Store applications that were previously exported
    from another Windows 11 computer. It copies the application packages to the
    WindowsApps folder, registers them, and restores the original permissions.
.PARAMETER SourcePath
    Path where the exported applications are located
.EXAMPLE
    .\Import-StoreApps.ps1 -SourcePath "E:\StoreApps"
.NOTES
    Requires administrative privileges
#>

param (
    [Parameter(Mandatory=$true)]
    [string]$SourcePath
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

# Check if source path exists
if (-not (Test-Path -Path $SourcePath)) {
    Write-ErrorLog "Source path does not exist: $SourcePath"
    exit 1
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
    
    # Process each app package in the source folder
    $appFolders = Get-ChildItem -Path $SourcePath -Directory
    
    if ($appFolders.Count -eq 0) {
        Write-ErrorLog "No app packages found in $SourcePath"
        exit 1
    }
    
    foreach ($appFolder in $appFolders) {
        $appName = $appFolder.Name
        $appSourcePath = $appFolder.FullName
        $appDestPath = Join-Path -Path $windowsAppsPath -ChildPath $appName
        
        Write-InfoLog "Processing $appName..."
        
        # Check if the app is already installed
        if (Test-Path -Path $appDestPath) {
            Write-InfoLog "$appName is already present in $windowsAppsPath - will attempt to register anyway"
        } else {
            # Copy the app package to the WindowsApps folder
            try {
                Write-InfoLog "Copying $appName to $appDestPath..."
                New-Item -ItemType Directory -Path $appDestPath -Force | Out-Null
                Copy-Item -Path "$appSourcePath\*" -Destination $appDestPath -Recurse -Force
                Write-InfoLog "Copy completed successfully."
            }
            catch {
                Write-ErrorLog "Failed to copy $appName: $_"
                continue
            }
        }
        
        # Register the app package
        try {
            $manifestPath = Join-Path -Path $appDestPath -ChildPath "AppxManifest.xml"
            
            if (-not (Test-Path -Path $manifestPath)) {
                Write-ErrorLog "AppxManifest.xml not found for $appName"
                continue
            }
            
            Write-InfoLog "Registering $appName..."
            Add-AppxPackage -Register $manifestPath -DisableDevelopmentMode -ForceApplicationShutdown
            Write-InfoLog "$appName registered successfully."
        }
        catch {
            Write-ErrorLog "Failed to register $appName: $_"
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

Write-InfoLog "App import operation completed."