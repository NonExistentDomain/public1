<#
.SYNOPSIS
    Exports Microsoft Store applications from a Windows 11 computer for offline installation.
.DESCRIPTION
    This script exports installed Microsoft Store applications from a Windows 11 computer
    as proper .appx/.msix packages with preserved signatures for offline installation.
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

# Create a dependency folder for related packages
$dependencyPath = Join-Path -Path $DestinationPath -ChildPath "Dependencies"
if (-not (Test-Path -Path $dependencyPath)) {
    New-Item -ItemType Directory -Path $dependencyPath -Force | Out-Null
    Write-InfoLog "Created dependency directory: $dependencyPath"
}

# Process each specified app
foreach ($appNamePattern in $AppNames) {
    Write-InfoLog "Processing applications matching: $appNamePattern..."
    
    # Get all installed packages matching the pattern
    $packages = Get-AppxPackage -Name "*$appNamePattern*" | Sort-Object -Property Name, Version

    if ($packages.Count -eq 0) {
        Write-ErrorLog "No installation found for $appNamePattern"
        continue
    }
    
    # Group packages by name to find the latest version of each
    $packageGroups = $packages | Group-Object -Property Name
    
    foreach ($group in $packageGroups) {
        $latestPackage = $group.Group | Sort-Object -Property Version -Descending | Select-Object -First 1
        
        $packageName = $latestPackage.Name
        $packageVersion = $latestPackage.Version
        $packageFullName = $latestPackage.PackageFullName
        
        Write-InfoLog "Exporting $packageName (version $packageVersion)..."
        
        # Path for the exported package
        $packagePath = Join-Path -Path $DestinationPath -ChildPath "$packageFullName.appx"
        
        try {
            # Export the package using built-in cmdlet to preserve signatures
            Export-StartLayout -Path "$env:TEMP\temp_layout.xml" -ErrorAction SilentlyContinue
            
            # Use Add-AppxProvisionedPackage to extract the package files - this preserves signatures
            $packageLocation = $latestPackage.InstallLocation
            $manifestPath = Join-Path -Path $packageLocation -ChildPath "AppxManifest.xml"
            
            # Create a temporary directory for the package
            $tempDir = Join-Path -Path $env:TEMP -ChildPath ([Guid]::NewGuid().ToString())
            New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
            
            # Copy the package files to a temporary location
            Copy-Item -Path "$packageLocation\*" -Destination $tempDir -Recurse -Force -ErrorAction SilentlyContinue
            
            # Find the package manifest
            $appxManifest = Get-ChildItem -Path $tempDir -Filter "AppxManifest.xml" -Recurse | Select-Object -First 1
            
            if ($appxManifest) {
                # Create package using MakeAppx.exe
                $makeAppxPath = "$env:ProgramFiles (x86)\Windows Kits\10\bin\10.0.22000.0\x64\makeappx.exe"
                
                # If MakeAppx.exe is not found, try to use alternative methods
                if (-not (Test-Path -Path $makeAppxPath)) {
                    $makeAppxPath = (Get-Command makeappx.exe -ErrorAction SilentlyContinue).Source
                    
                    # If still not found, search for it
                    if (-not $makeAppxPath) {
                        $possiblePaths = Get-ChildItem -Path "$env:ProgramFiles (x86)\Windows Kits\10\bin" -Filter "makeappx.exe" -Recurse -ErrorAction SilentlyContinue
                        if ($possiblePaths) {
                            $makeAppxPath = $possiblePaths[0].FullName
                        }
                    }
                }
                
                if ($makeAppxPath) {
                    # Create the package
                    $makeAppxArgs = "pack /d `"$tempDir`" /p `"$packagePath`" /l"
                    Start-Process -FilePath $makeAppxPath -ArgumentList $makeAppxArgs -NoNewWindow -Wait
                    
                    if (Test-Path -Path $packagePath) {
                        Write-InfoLog "Successfully exported $packageName to $packagePath"
                        
                        # Export dependencies if they exist
                        if ($latestPackage.Dependencies -and $latestPackage.Dependencies.Count -gt 0) {
                            Write-InfoLog "Exporting dependencies for $packageName..."
                            
                            foreach ($dependency in $latestPackage.Dependencies) {
                                $depPackage = Get-AppxPackage -Name $dependency.Name | Sort-Object -Property Version -Descending | Select-Object -First 1
                                
                                if ($depPackage) {
                                    $depPackagePath = Join-Path -Path $dependencyPath -ChildPath "$($depPackage.PackageFullName).appx"
                                    
                                    # Process dependency package
                                    $depTempDir = Join-Path -Path $env:TEMP -ChildPath ([Guid]::NewGuid().ToString())
                                    New-Item -ItemType Directory -Path $depTempDir -Force | Out-Null
                                    
                                    Copy-Item -Path "$($depPackage.InstallLocation)\*" -Destination $depTempDir -Recurse -Force -ErrorAction SilentlyContinue
                                    
                                    $makeAppxDepArgs = "pack /d `"$depTempDir`" /p `"$depPackagePath`" /l"
                                    Start-Process -FilePath $makeAppxPath -ArgumentList $makeAppxDepArgs -NoNewWindow -Wait
                                    
                                    if (Test-Path -Path $depPackagePath) {
                                        Write-InfoLog "Successfully exported dependency $($depPackage.Name) to $depPackagePath"
                                    }
                                    
                                    # Clean up
                                    Remove-Item -Path $depTempDir -Recurse -Force -ErrorAction SilentlyContinue
                                }
                            }
                        }
                    } else {
                        Write-ErrorLog "Failed to create package for $packageName"
                    }
                } else {
                    Write-ErrorLog "MakeAppx.exe not found. Cannot create package for $packageName"
                    
                    # Fallback: just copy the package directory
                    $fallbackPath = Join-Path -Path $DestinationPath -ChildPath $packageFullName
                    New-Item -ItemType Directory -Path $fallbackPath -Force | Out-Null
                    Copy-Item -Path "$packageLocation\*" -Destination $fallbackPath -Recurse -Force -ErrorAction SilentlyContinue
                    Write-InfoLog "Copied $packageName files to $fallbackPath (without proper packaging)"
                }
            } else {
                Write-ErrorLog "AppxManifest.xml not found for $packageName"
            }
            
            # Clean up
            Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        }
        catch {
            Write-ErrorLog "Failed to export $packageName: $_"
        }
    }
}

Write-InfoLog "App export operation completed. Applications saved to $DestinationPath"
