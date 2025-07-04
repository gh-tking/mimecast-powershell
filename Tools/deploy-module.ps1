[CmdletBinding()]
param (
    [Parameter()]
    [string]$ModulePath = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)),
    
    [Parameter()]
    [string]$ModuleName = 'MimecastApi',
    
    [Parameter()]
    [string]$DestinationPath = "$env:USERPROFILE\Documents\WindowsPowerShell\Modules\$ModuleName"
)

# Ensure destination directory exists
if (-not (Test-Path -Path $DestinationPath)) {
    New-Item -Path $DestinationPath -ItemType Directory -Force | Out-Null
}

# Install required modules
if (-not (Get-Module -Name Microsoft.PowerShell.SecretManagement -ListAvailable)) {
    Install-Module -Name Microsoft.PowerShell.SecretManagement -Force
}
if (-not (Get-Module -Name Microsoft.PowerShell.SecretStore -ListAvailable)) {
    Install-Module -Name Microsoft.PowerShell.SecretStore -Force
}

# Copy module files to PowerShell modules directory
Copy-Item -Path "$ModulePath\*" -Destination $DestinationPath -Recurse -Force

# Create scripts directory in Program Files
$scriptsPath = "C:\Program Files\MimecastApi\Scripts"
if (-not (Test-Path -Path $scriptsPath)) {
    New-Item -Path $scriptsPath -ItemType Directory -Force | Out-Null
}

# Copy log collector script
Copy-Item -Path "$ModulePath\Tools\Invoke-MimecastLogCollection.ps1" -Destination $scriptsPath -Force

# Create state and logs directories
$stateDir = "C:\ProgramData\MimecastApi\State"
$logDir = "C:\ProgramData\MimecastApi\Logs"
New-Item -Path $stateDir, $logDir -ItemType Directory -Force | Out-Null

Write-Host "MimecastApi module deployed successfully to $DestinationPath"
Write-Host "Log collector script deployed to $scriptsPath"
Write-Host "State directory created at $stateDir"
Write-Host "Log directory created at $logDir"

# Import the module to verify deployment
Import-Module -Name $ModuleName -Force
Get-Module -Name $ModuleName | Format-List