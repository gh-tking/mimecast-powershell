function Set-ApiSecret {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$Name,
        
        [Parameter(Mandatory)]
        [securestring]$Secret,
        
        [Parameter()]
        [string]$Vault = 'ApiSecrets'
    )
    
    try {
        # Ensure the vault exists
        if (-not (Get-SecretVault -Name $Vault -ErrorAction SilentlyContinue)) {
            Register-SecretVault -Name $Vault -ModuleName Microsoft.PowerShell.SecretStore -DefaultVault
        }
        
        Set-Secret -Name $Name -SecureStringSecret $Secret -Vault $Vault
    }
    catch {
        Write-Error -ErrorRecord $_
    }
}

function Get-ApiSecret {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$Name,
        
        [Parameter()]
        [string]$Vault = 'ApiSecrets'
    )
    
    try {
        Get-Secret -Name $Name -Vault $Vault -AsSecureString
    }
    catch {
        Write-Error -ErrorRecord $_
    }
}