# PowerShell script to automatically identify and remediate NetFx40_LegacySecurityPolicy in .NET config files

# Define the search path
targetPath = "C:\"

# Find all .exe.config files and scan for NetFx40_LegacySecurityPolicy
$configFiles = Get-ChildItem -Path $targetPath -Filter "*.exe.config" -Recurse -ErrorAction SilentlyContinue

$foundIssues = @()

foreach ($file in $configFiles) {
    [xml]$xmlContent = Get-Content $file.FullName -ErrorAction SilentlyContinue

    if ($xmlContent) {
        $legacyPolicyNodes = $xmlContent.configuration.runtime.NetFx40_LegacySecurityPolicy

        if ($legacyPolicyNodes -and $legacyPolicyNodes.enabled -eq "true") {
            $foundIssues += [PSCustomObject]@{
                FilePath  = $file.FullName
                PolicySet = $legacyPolicyNodes.enabled
            }

            # Disable the legacy security policy
            $legacyPolicyNodes.enabled = "false"

            # Backup original file
            Copy-Item -Path $file.FullName -Destination ($file.FullName + ".backup") -Force

            # Save updated config file
            $xmlContent.Save($file.FullName)
        }
    }
}

# Output the results
if ($foundIssues.Count -gt 0) {
    Write-Output "Applications with NetFx40_LegacySecurityPolicy enabled found and corrected:"
    $foundIssues | Format-Table -AutoSize
    Write-Warning "Original configuration files have been backed up with a .backup extension."
} else {
    Write-Output "No applications with NetFx40_LegacySecurityPolicy enabled were detected."
}