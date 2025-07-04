# Connection.ps1
# Contains functions for connecting to and managing Mimecast API connections

<#
.SYNOPSIS
Establishes a connection to the Mimecast API.

.DESCRIPTION
Establishes a connection to the Mimecast API using provided credentials and
configuration. The function supports multiple authentication methods and stores
the connection information for use by other cmdlets in the module.

.PARAMETER AccessKey
The API access key from Mimecast.
Can be obtained from the Mimecast Administration Console.

.PARAMETER SecretKey
The API secret key from Mimecast.
Can be obtained from the Mimecast Administration Console.

.PARAMETER ApplicationId
Optional. The application ID for API access.
Used for tracking API usage in Mimecast.

.PARAMETER ApplicationKey
Optional. The application key for API access.
Used for tracking API usage in Mimecast.

.PARAMETER Region
Optional. The Mimecast region to connect to.
If not specified, determined automatically from email domain.
Valid values: EU, US, DE, CA, AU, ZA, etc.

.PARAMETER EmailDomain
Optional. The email domain for API access.
Used to determine the correct API endpoint if region not specified.

.PARAMETER BaseUri
Optional. The base URI for API requests.
If not specified, determined automatically from region or email domain.
Example: 'https://eu-api.mimecast.com'.

.PARAMETER Credential
Optional. PSCredential object containing AccessKey and SecretKey.
Alternative to providing them as separate parameters.

.PARAMETER UseSecretStore
Optional. Whether to use the SecretManagement module for storing credentials.
Default is $false.

.PARAMETER SecretVaultName
Optional. The name of the secret vault to use.
Required if UseSecretStore is $true.

.PARAMETER Force
Optional. Force a new connection even if one already exists.
Default is $false.

.INPUTS
None. You cannot pipe objects to Connect-Mimecast.

.OUTPUTS
System.Object
Returns a connection information object including:
- accessKey: The API access key being used
- region: The connected region
- baseUri: The base URI for API requests
- connected: Whether connection was successful
- lastConnected: When connection was established

.EXAMPLE
C:\> Connect-Mimecast `
    -AccessKey "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -SecretKey "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -Region "EU"

Connects to the EU region using explicit keys.

.EXAMPLE
C:\> $cred = Get-Credential -Message "Enter Mimecast API credentials"
C:\> Connect-Mimecast `
    -Credential $cred `
    -EmailDomain "example.com"

Connects using credentials and determines region from domain.

.EXAMPLE
C:\> Connect-Mimecast `
    -UseSecretStore `
    -SecretVaultName "MimecastVault" `
    -ApplicationId "MyApp" `
    -ApplicationKey "xxxxxxxx"

Connects using credentials from a secret vault.

.EXAMPLE
C:\> Connect-Mimecast `
    -BaseUri "https://custom-api.mimecast.com" `
    -AccessKey "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -SecretKey "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -Force

Forces a new connection to a custom API endpoint.

.NOTES
- Store credentials securely using SecretManagement
- Connection persists for module session
- Some operations may require re-authentication
- API permissions depend on the API key's assigned roles
- Required API permissions:
  * Any permission sufficient for testing connection
#>
function Connect-Mimecast {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]$ClientIdName = 'MimecastClientId',
        
        [Parameter()]
        [string]$ClientSecretName = 'MimecastClientSecret',
        
        [Parameter(Mandatory)]
        [ValidateSet('EU', 'US', 'DE', 'CA', 'ZA', 'AU', 'Offshore')]
        [string]$Region,

        [Parameter()]
        [switch]$Setup
    )
    
    try {
        # If Setup switch is provided, initialize the module first
        if ($Setup) {
            Initialize-MimecastModule
        }

        # Get the OAuth2 credentials from the vault
        $ClientId = Get-ApiSecret -Name $ClientIdName -ErrorAction SilentlyContinue
        $ClientSecret = Get-ApiSecret -Name $ClientSecretName -ErrorAction SilentlyContinue
        
        if (-not $ClientId -or -not $ClientSecret) {
            Write-Error "OAuth2 credentials not found. Please run Initialize-MimecastModule first or provide valid credential names."
            throw "Failed to retrieve OAuth2 credentials from vault"
        }

        # Convert SecureString to plain text for API call
        $ClientIdPlain = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
            [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($ClientId))
        $ClientSecretPlain = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
            [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($ClientSecret))
        
        # Set the region-specific base URI
        $RegionUri = switch ($Region) {
            'EU' { 'https://eu-api.mimecast.com' }
            'US' { 'https://us-api.mimecast.com' }
            'DE' { 'https://de-api.mimecast.com' }
            'CA' { 'https://ca-api.mimecast.com' }
            'ZA' { 'https://za-api.mimecast.com' }
            'AU' { 'https://au-api.mimecast.com' }
            'Offshore' { 'https://off-api.mimecast.com' }
        }

        # Get OAuth2 token
        $tokenParams = @{
            Method = 'POST'
            Uri = "$RegionUri/oauth/token"
            Body = @{
                grant_type = 'client_credentials'
                client_id = $ClientIdPlain
                client_secret = $ClientSecretPlain
            }
            Headers = @{
                'Content-Type' = 'application/x-www-form-urlencoded'
            }
        }

        $tokenResponse = Invoke-RestMethod @tokenParams

        if (-not $tokenResponse.access_token) {
            throw "Failed to obtain access token"
        }
        
        # Update module's private data
        $ExecutionContext.SessionState.Module.PrivateData.RegionUri = $RegionUri
        $ExecutionContext.SessionState.Module.PrivateData.BaseUri = "$RegionUri/api/v2"
        $ExecutionContext.SessionState.Module.PrivateData.ClientId = $ClientId
        $ExecutionContext.SessionState.Module.PrivateData.ClientSecret = $ClientSecret
        $ExecutionContext.SessionState.Module.PrivateData.AccessToken = $tokenResponse.access_token
        $ExecutionContext.SessionState.Module.PrivateData.TokenExpiresAt = [DateTime]::UtcNow.AddSeconds($tokenResponse.expires_in)
        
        # Store connection info
        $Script:MimecastConnection = @{
            Connected = $true
            ConnectedAt = Get-Date
            Region = $Region
            TokenExpiresAt = $ExecutionContext.SessionState.Module.PrivateData.TokenExpiresAt
        }
        
        Write-ApiLog "Successfully connected to Mimecast API 2.0 ($Region region)"
        Write-ApiLog "Access token will expire at $($ExecutionContext.SessionState.Module.PrivateData.TokenExpiresAt)"
        
        # Test the connection by getting the API status
        $null = Get-MimecastSystemInfo
    }
    catch {
        Write-Error -ErrorRecord $_
        
        # Clear connection info on failure
        $Script:MimecastConnection = $null
        
        # Clear sensitive data from private data
        $ExecutionContext.SessionState.Module.PrivateData.ClientId = ''
        $ExecutionContext.SessionState.Module.PrivateData.ClientSecret = ''
        $ExecutionContext.SessionState.Module.PrivateData.AccessToken = ''
        $ExecutionContext.SessionState.Module.PrivateData.TokenExpiresAt = $null
    }
}

<#
.SYNOPSIS
Disconnects from the Mimecast API.

.DESCRIPTION
Terminates the current connection to the Mimecast API and cleans up stored
connection information. The function includes ShouldProcess support for
confirmation and can optionally remove stored credentials.

.PARAMETER RemoveStoredCredentials
Optional. Whether to remove stored credentials from the secret store.
Default is $false.
Only applies if credentials were stored using SecretManagement.

.PARAMETER SecretVaultName
Optional. The name of the secret vault containing stored credentials.
Required if RemoveStoredCredentials is $true.

.INPUTS
None. You cannot pipe objects to Disconnect-Mimecast.

.OUTPUTS
None. The function does not return any output.

.EXAMPLE
C:\> Disconnect-Mimecast
Disconnects from the API, leaving stored credentials intact.

.EXAMPLE
C:\> Disconnect-Mimecast -RemoveStoredCredentials -SecretVaultName "MimecastVault"
Disconnects and removes stored credentials from the vault.

.EXAMPLE
C:\> Disconnect-Mimecast -WhatIf
Shows what would happen if you disconnected from the API.

.NOTES
- Use -WhatIf to preview changes
- Stored credentials remain unless explicitly removed
- Active operations may fail after disconnection
- API permissions not required (local operation only)
#>
function Disconnect-Mimecast {
    [CmdletBinding()]
    param()
    
    try {
        # Clear connection info
        $Script:MimecastConnection = $null
        
        # Clear sensitive data from private data
        $ExecutionContext.SessionState.Module.PrivateData.ClientId = ''
        $ExecutionContext.SessionState.Module.PrivateData.ClientSecret = ''
        $ExecutionContext.SessionState.Module.PrivateData.AccessToken = ''
        $ExecutionContext.SessionState.Module.PrivateData.TokenExpiresAt = $null
        
        Write-ApiLog "Disconnected from Mimecast API 2.0"
    }
    catch {
        Write-Error -ErrorRecord $_
    }
}

<#
.SYNOPSIS
Configures settings for the Mimecast API module.

.DESCRIPTION
Updates configuration settings for the Mimecast API module, including base URI,
default region, logging settings, and other module-wide parameters. The function
includes ShouldProcess support for confirmation.

.PARAMETER BaseUri
Optional. The base URI for API requests.
Example: 'https://eu-api.mimecast.com'.

.PARAMETER Region
Optional. The default region for connections.
Valid values: EU, US, DE, CA, AU, ZA, etc.

.PARAMETER LogFilePath
Optional. Path for module logging.
Set to $null to disable logging.

.PARAMETER LogLevel
Optional. Minimum level for log entries.
Valid values: Debug, Info, Warning, Error.

.PARAMETER DefaultPageSize
Optional. Default page size for paginated results.
Valid range: 1-500.

.PARAMETER DefaultStartDate
Optional. Default start date for searches.
Example: (Get-Date).AddDays(-7).

.PARAMETER TimeoutSeconds
Optional. Timeout for API requests in seconds.
Default is 30 seconds.

.PARAMETER RetryCount
Optional. Number of retries for failed requests.
Default is 3 retries.

.PARAMETER RetryDelaySeconds
Optional. Delay between retries in seconds.
Default is 2 seconds.

.PARAMETER Force
Optional. Force update of read-only settings.
Default is $false.

.INPUTS
None. You cannot pipe objects to Set-MimecastConfiguration.

.OUTPUTS
None. The function does not return any output.

.EXAMPLE
C:\> Set-MimecastConfiguration `
    -BaseUri "https://eu-api.mimecast.com" `
    -Region "EU" `
    -LogFilePath "C:\Logs\Mimecast.log" `
    -LogLevel "Info"

Configures basic module settings.

.EXAMPLE
C:\> Set-MimecastConfiguration `
    -DefaultPageSize 100 `
    -DefaultStartDate (Get-Date).AddDays(-30) `
    -TimeoutSeconds 60

Configures performance and search settings.

.EXAMPLE
C:\> Set-MimecastConfiguration `
    -RetryCount 5 `
    -RetryDelaySeconds 5 `
    -Force

Forces update of retry settings.

.EXAMPLE
C:\> Set-MimecastConfiguration -WhatIf `
    -BaseUri "https://custom-api.mimecast.com" `
    -LogLevel "Debug"

Shows what would happen if you updated these settings.

.NOTES
- Some settings require module reinitialization
- Log file directory must be writable
- Some settings may affect performance
- Use -WhatIf to preview changes
- API permissions not required (local operation only)
#>
function Set-MimecastConfiguration {
    [CmdletBinding()]
    param (
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$LogFilePath,
        
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$ApiVaultName
    )
    
    try {
        if ($LogFilePath) {
            $ExecutionContext.SessionState.Module.PrivateData.LogFilePath = $LogFilePath
            Write-ApiLog "Log file path set to: $LogFilePath"
        }
        
        if ($ApiVaultName) {
            $ExecutionContext.SessionState.Module.PrivateData.ApiVaultName = $ApiVaultName
            Write-ApiLog "API vault name set to: $ApiVaultName"
        }
    }
    catch {
        Write-Error -ErrorRecord $_
    }
}

<#
.SYNOPSIS
Gets current configuration settings for the Mimecast API module.

.DESCRIPTION
Retrieves the current configuration settings for the Mimecast API module,
including connection details, logging settings, and module-wide parameters.
The function can return all settings or filter by category.

.PARAMETER Category
Optional. Filter settings by category. Valid values:
- Connection: Base URI, region, timeout settings
- Logging: Log file path, level, retention
- Performance: Page size, retry settings
- All: All settings (default)

.PARAMETER AsHashtable
Optional. Return settings as a hashtable instead of an object.
Default is $false.

.PARAMETER ExcludeSecrets
Optional. Exclude sensitive information from output.
Default is $true.

.INPUTS
None. You cannot pipe objects to Get-MimecastConfiguration.

.OUTPUTS
System.Object or System.Collections.Hashtable
Returns configuration settings either as an object or hashtable.
Settings include:
- baseUri: Base URI for API requests
- region: Default region for connections
- logFilePath: Path for module logging
- logLevel: Minimum level for log entries
- defaultPageSize: Default page size for pagination
- defaultStartDate: Default start date for searches
- timeoutSeconds: API request timeout
- retryCount: Number of retries for failed requests
- retryDelaySeconds: Delay between retries
- connected: Current connection status
- lastConnected: Last connection timestamp

.EXAMPLE
C:\> Get-MimecastConfiguration
Gets all configuration settings.

.EXAMPLE
C:\> Get-MimecastConfiguration -Category Connection
Gets only connection-related settings.

.EXAMPLE
C:\> Get-MimecastConfiguration -AsHashtable -ExcludeSecrets $false |
    ConvertTo-Json -Depth 10

Gets all settings including secrets as JSON.

.EXAMPLE
C:\> $config = Get-MimecastConfiguration -Category Performance
C:\> Set-MimecastConfiguration `
    -DefaultPageSize ($config.defaultPageSize * 2) `
    -RetryCount ($config.retryCount + 2)

Gets current performance settings and updates them.

.NOTES
- Some settings may be empty if not configured
- Sensitive data is masked by default
- Connection status is read-only
- API permissions not required (local operation only)
#>
function Get-MimecastConfiguration {
    [CmdletBinding()]
    param()
    
    @{
        BaseUri = $ExecutionContext.SessionState.Module.PrivateData.BaseUri
        RegionUri = $ExecutionContext.SessionState.Module.PrivateData.RegionUri
        LogFilePath = $ExecutionContext.SessionState.Module.PrivateData.LogFilePath
        ApiVaultName = $ExecutionContext.SessionState.Module.PrivateData.ApiVaultName
        ApplicationId = $ExecutionContext.SessionState.Module.PrivateData.ApplicationId
        Connection = $Script:MimecastConnection
    }
}