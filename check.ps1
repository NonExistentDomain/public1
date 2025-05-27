# PowerShell script to identify applications using NetFx40_LegacySecurityPolicy
# and guide the administrator on remediation actions per DoD STIG guidelines

# Define the registry paths to check
$registryPaths = @(
    "HKLM:\SOFTWARE\Microsoft\.NETFramework",
    "HKLM:\SOFTWARE\Wow6432Node\Microsoft\.NETFramework"
)

$legacyPolicyApps = @()

# Function to check registry keys for LegacySecurityPolicy
foreach ($path in $registryPaths) {
    if (Test-Path $path) {
        Get-ChildItem -Path $path | ForEach-Object {
            $subKey = $_.PSPath
            $legacyPolicy = Get-ItemProperty -Path $subKey -Name "NetFx40_LegacySecurityPolicy" -ErrorAction SilentlyContinue
            if ($legacyPolicy -and $legacyPolicy.NetFx40_LegacySecurityPolicy -eq 1) {
                $appDetails = [PSCustomObject]@{
                    ApplicationRegistryPath = $subKey
                    LegacyPolicyEnabled     = $legacyPolicy.NetFx40_LegacySecurityPolicy
                }
                $legacyPolicyApps += $appDetails
            }
        }
    }
}

# Output findings
if ($legacyPolicyApps.Count -gt 0) {
    Write-Output "Applications with NetFx40_LegacySecurityPolicy enabled detected:"
    $legacyPolicyApps | Format-Table -AutoSize

    Write-Warning "CAS policy is enabled for these applications. Ensure previous .NET STIG guidance is applied."
    Write-Warning "Refer to the DISA .NET Framework STIG document to apply necessary security configurations."

    # Example remediation action (manual review required):
    # Set NetFx40_LegacySecurityPolicy to 0 if it's safe to disable
    # foreach ($app in $legacyPolicyApps) {
    #     Set-ItemProperty -Path $app.ApplicationRegistryPath -Name "NetFx40_LegacySecurityPolicy" -Value 0
    # }
} else {
    Write-Output "No applications with NetFx40_LegacySecurityPolicy enabled were detected."
}