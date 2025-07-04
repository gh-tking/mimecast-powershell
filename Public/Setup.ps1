function Initialize-MimecastModule {
    [CmdletBinding()]
    param (
        [Parameter()]
        [switch]$ForceSetup
    )

    try {
        # Get the vault name from module's private data
        $vaultName = $ExecutionContext.SessionState.Module.PrivateData.ApiVaultName
        
        # Check if the vault exists, if not create it
        if (-not (Get-SecretVault -Name $vaultName -ErrorAction SilentlyContinue)) {
            Write-Verbose "Creating secret vault: $vaultName"
            Register-SecretVault -Name $vaultName -ModuleName Microsoft.PowerShell.SecretStore -DefaultVault
        }

        # Check for existing secrets
        $clientIdExists = $null -ne (Get-ApiSecret -Name 'MimecastClientId' -ErrorAction SilentlyContinue)
        $clientSecretExists = $null -ne (Get-ApiSecret -Name 'MimecastClientSecret' -ErrorAction SilentlyContinue)

        if (-not $ForceSetup -and $clientIdExists -and $clientSecretExists) {
            Write-Host "Mimecast API credentials are already configured. Use -ForceSetup to reconfigure."
            return
        }

        Write-Host "Initializing Mimecast API module..."
        
        # Prompt for credentials
        $clientId = Read-Host -Prompt "Enter your Mimecast OAuth2 Client ID" -AsSecureString
        $clientSecret = Read-Host -Prompt "Enter your Mimecast OAuth2 Client Secret" -AsSecureString

        # Store credentials
        Set-ApiSecret -Name 'MimecastClientId' -Secret $clientId -Vault $vaultName
        Set-ApiSecret -Name 'MimecastClientSecret' -Secret $clientSecret -Vault $vaultName

        Write-Host "Mimecast API credentials have been securely stored."
        Write-Host "You can now use Connect-Mimecast -Region <region> to connect to the API."
    }
    catch {
        Write-Error -ErrorRecord $_
        throw "Failed to initialize Mimecast API module: $_"
    }
}