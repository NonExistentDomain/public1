<#
.SYNOPSIS
    Installs Microsoft Store applications on an offline Windows 11 computer.
.DESCRIPTION
    This script installs or updates Microsoft Store applications that were previously exported
    from another Windows 11 computer.
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

# Check if dependencies path exists
$dependencyPath = Join-Path -Path $SourcePath -ChildPath "Dependencies"
$hasDependencies = Test-Path -Path $dependencyPath

# Process all .appx and .msix files in the source folder
$appPackages = Get-ChildItem -Path $SourcePath -Filter "*.appx" -ErrorAction SilentlyContinue
$appPackages += Get-ChildItem -Path $SourcePath -Filter "*.msix" -ErrorAction SilentlyContinue
$appPackages += Get-ChildItem -Path $SourcePath -Filter "*.appxbundle" -ErrorAction SilentlyContinue
$appPackages += Get-ChildItem -Path $SourcePath -Filter "*.msixbundle" -ErrorAction SilentlyContinue

if ($appPackages.Count -eq 0) {
    Write-InfoLog "No .appx/.msix packages found. Checking for exported app directories..."
    
    # Fallback: Process directories that might contain app files
    $appFolders = Get-ChildItem -Path $SourcePath -Directory | Where-Object { $_.Name -like "Microsoft.*" }
    
    if ($appFolders.Count -eq 0) {
        Write-ErrorLog "No app packages found in $SourcePath"
        exit 1
    }
    
    Write-InfoLog "Found $($appFolders.Count) app directories. Will attempt to register them directly."
    
    foreach ($appFolder in $appFolders) {
        $appName = $appFolder.Name
        $appSourcePath = $appFolder.FullName
        
        Write-InfoLog "Processing $appName..."
        
        # Register the app package
        try {
            $manifestPath = Join-Path -Path $appSourcePath -ChildPath "AppxManifest.xml"
            
            if (-not (Test-Path -Path $manifestPath)) {
                Write-ErrorLog "AppxManifest.xml not found for $appName"
                continue
            }
            
            Write-InfoLog "Registering $appName..."
            
            # Try to register with AllUsers if possible (requires Windows 10 1809 or later)
            try {
                Add-AppxPackage -Register $manifestPath -DisableDevelopmentMode -ForceUpdateFromAnyVersion -AllUsers
            }
            catch {
                # Fall back to per-user registration
                Add-AppxPackage -Register $manifestPath -DisableDevelopmentMode -ForceUpdateFromAnyVersion
            }
            
            Write-InfoLog "$appName registered successfully."
        }
        catch {
            Write-ErrorLog "Failed to register $appName: $_"
        }
    }
} else {
    Write-InfoLog "Found $($appPackages.Count) .appx/.msix package files. Processing..."
    
    # First, prepare the dependency packages if they exist
    $dependencies = @()
    if ($hasDependencies) {
        $depPackages = Get-ChildItem -Path $dependencyPath -Filter "*.appx" -ErrorAction SilentlyContinue
        $depPackages += Get-ChildItem -Path $dependencyPath -Filter "*.msix" -ErrorAction SilentlyContinue
        $depPackages += Get-ChildItem -Path $dependencyPath -Filter "*.appxbundle" -ErrorAction SilentlyContinue
        $depPackages += Get-ChildItem -Path $dependencyPath -Filter "*.msixbundle" -ErrorAction SilentlyContinue
        
        if ($depPackages.Count -gt 0) {
            $dependencies = $depPackages.FullName
            Write-InfoLog "Found $($depPackages.Count) dependency packages."
        }
    }
    
    # Install each app package
    foreach ($package in $appPackages) {
        Write-InfoLog "Installing $($package.Name)..."
        
        try {
            # Check if it's an update or new install
            $packageInfoPattern = "^(.*?)_(\d+\.\d+\.\d+\.\d+)_"
            if ($package.Name -match $packageInfoPattern) {
                $packageName = $matches[1]
                $packageVersion = $matches[2]
                
                $existingPackage = Get-AppxPackage -Name $packageName -ErrorAction SilentlyContinue
                
                if ($existingPackage) {
                    Write-InfoLog "$packageName is already installed (version $($existingPackage.Version)). Attempting update to version $packageVersion..."
                    
                    # Try to use AllUsers if possible
                    try {
                        if ($dependencies.Count -gt 0) {
                            Add-AppxPackage -Path $package.FullName -DependencyPath $dependencies -ForceUpdateFromAnyVersion -AllUsers
                        } else {
                            Add-AppxPackage -Path $package.FullName -ForceUpdateFromAnyVersion -AllUsers
                        }
                    }
                    catch {
                        # Fall back to per-user installation
                        if ($dependencies.Count -gt 0) {
                            Add-AppxPackage -Path $package.FullName -DependencyPath $dependencies -ForceUpdateFromAnyVersion
                        } else {
                            Add-AppxPackage -Path $package.FullName -ForceUpdateFromAnyVersion
                        }
                    }
                } else {
                    Write-InfoLog "Installing new package $packageName version $packageVersion..."
                    
                    # Try to use AllUsers if possible
                    try {
                        if ($dependencies.Count -gt 0) {
                            Add-AppxPackage -Path $package.FullName -DependencyPath $dependencies -AllUsers
                        } else {
                            Add-AppxPackage -Path $package.FullName -AllUsers
                        }
                    }
                    catch {
                        # Fall back to per-user installation
                        if ($dependencies.Count -gt 0) {
                            Add-AppxPackage -Path $package.FullName -DependencyPath $dependencies
                        } else {
                            Add-AppxPackage -Path $package.FullName
                        }
                    }
                }
                
                Write-InfoLog "$($package.Name) installed successfully."
            } else {
                # For packages that don't match the naming pattern
                Write-InfoLog "Installing $($package.Name) (non-standard naming)..."
                
                # Try to use AllUsers if possible
                try {
                    if ($dependencies.Count -gt 0) {
                        Add-AppxPackage -Path $package.FullName -DependencyPath $dependencies -ForceUpdateFromAnyVersion -AllUsers
                    } else {
                        Add-AppxPackage -Path $package.FullName -ForceUpdateFromAnyVersion -AllUsers
                    }
                }
                catch {
                    # Fall back to per-user installation
                    if ($dependencies.Count -gt 0) {
                        Add-AppxPackage -Path $package.FullName -DependencyPath $dependencies -ForceUpdateFromAnyVersion
                    } else {
                        Add-AppxPackage -Path $package.FullName -ForceUpdateFromAnyVersion
                    }
                }
                
                Write-InfoLog "$($package.Name) installed successfully."
            }
        }
        catch {
            Write-ErrorLog "Failed to install $($package.Name): $_"
            
            # Try an alternative method if the primary method fails
            Write-InfoLog "Attempting alternative installation method for $($package.Name)..."
            try {
                Add-AppxPackage -Path $package.FullName -ForceApplicationShutdown -ForceUpdateFromAnyVersion
                Write-InfoLog "Alternative installation method succeeded for $($package.Name)."
            }
            catch {
                Write-ErrorLog "Alternative installation method also failed for $($package.Name): $_"
            }
        }
    }
}

Write-InfoLog "App import operation completed."
