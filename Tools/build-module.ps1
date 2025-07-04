[CmdletBinding()]
param (
    [Parameter()]
    [string]$ModulePath = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)),
    
    [Parameter()]
    [string]$OutputPath = ".\build",
    
    [Parameter()]
    [switch]$RunTests
)

# Create build directory if it doesn't exist
if (-not (Test-Path -Path $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

# Run Pester tests if specified
if ($RunTests) {
    if (-not (Get-Module -Name Pester -ListAvailable)) {
        Install-Module -Name Pester -Force -SkipPublisherCheck
    }
    
    # Install required modules for testing
    if (-not (Get-Module -Name Microsoft.PowerShell.SecretManagement -ListAvailable)) {
        Install-Module -Name Microsoft.PowerShell.SecretManagement -Force
    }
    if (-not (Get-Module -Name Microsoft.PowerShell.SecretStore -ListAvailable)) {
        Install-Module -Name Microsoft.PowerShell.SecretStore -Force
    }
    
    $testResults = Invoke-Pester -Path "$ModulePath\Tests" -PassThru
    
    if ($testResults.FailedCount -gt 0) {
        throw "One or more tests failed. Build aborted."
    }
}

# Copy module files to build directory
Copy-Item -Path "$ModulePath\*" -Destination $OutputPath -Recurse -Force

Write-Host "MimecastApi module built successfully at $OutputPath"