# Module-level PrivateData
$PrivateData = @{
    BaseUri = "https://api.mimecast.com/api/v2/" # Base URI for Mimecast API 2.0
    LogFilePath = "$($PSScriptRoot)\Mimecast.log"
    ApiVaultName = "MimecastVault"
    # Mimecast-specific configurations
    ApiVersion = "v2"
    RegionUri = "" # Will be set during connection based on account region
    ClientId = "" # OAuth2 client ID
    ClientSecret = "" # OAuth2 client secret
    AccessToken = "" # OAuth2 access token
    TokenExpiresAt = $null # DateTime when the token expires
}

# Get all script files
$Public = @(Get-ChildItem -Path $PSScriptRoot\Public\*.ps1 -ErrorAction SilentlyContinue)
$Private = @(Get-ChildItem -Path $PSScriptRoot\Private\*.ps1 -ErrorAction SilentlyContinue)
$Resources = @(Get-ChildItem -Path $PSScriptRoot\Public\Resources\*.ps1 -ErrorAction SilentlyContinue)

# Dot source the files
foreach ($import in @($Public + $Private + $Resources)) {
    try {
        . $import.FullName
    }
    catch {
        Write-Error -Message "Failed to import function $($import.FullName): $_"
    }
}

# Export Public functions
Export-ModuleMember -Function $Public.BaseName
Export-ModuleMember -Function $Resources.BaseName

# Export the PrivateData
$ExecutionContext.SessionState.Module.PrivateData = $PrivateData