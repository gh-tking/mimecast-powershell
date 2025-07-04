@{
    RootModule = 'Mimecast.psm1'
    ModuleVersion = '0.1.0'
    GUID = [guid]::NewGuid().ToString()
    Author = 'Your Name'
    CompanyName = 'Your Company'
    Copyright = '(c) 2025 Your Company. All rights reserved.'
    Description = 'PowerShell module for interacting with Mimecast services'
    PowerShellVersion = '7.0'
    RequiredModules = @(
        @{ModuleName = 'Microsoft.PowerShell.SecretManagement'; ModuleVersion = '1.0.0'},
        @{ModuleName = 'Microsoft.PowerShell.SecretStore'; ModuleVersion = '1.0.0'}
    )
    FunctionsToExport = @(
        'Initialize-MimecastModule',
        'Connect-Mimecast',
        'Disconnect-Mimecast',
        'Set-MimecastConfiguration',
        'Get-MimecastConfiguration',
        'Get-MimecastSystemInfo',
        'Register-MimecastLogCollectorTask',
        'Unregister-MimecastLogCollectorTask',
        'Get-MimecastLogCollectorTask',
        'Start-MimecastLogCollectorTask',
        'Stop-MimecastLogCollectorTask',
        'Get-MimecastMessageTrace',
        'Search-MimecastMessage',
        'Get-MimecastMessageInfo',
        'Get-MimecastHeldMessage',
        'Get-MimecastProcessingMessage'
    )
    PrivateData = $true # Load PrivateData from .psm1
}