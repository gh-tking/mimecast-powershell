# Mimecast PowerShell Module

A PowerShell module for interacting with Mimecast services, providing cmdlets for email security, archiving, and continuity features.

## Features

- Message tracking and tracing
- Archive search and retrieval
- Message processing and quarantine management
- Secure credential management
- Comprehensive logging
- Pipeline support
- Automatic pagination

## Installation

```powershell
# Install from PowerShell Gallery (coming soon)
Install-Module -Name Mimecast -Scope CurrentUser

# Or clone and import manually
git clone https://github.com/yourusername/mimecast-powershell.git
Import-Module ./mimecast-powershell/Mimecast.psd1
```

## Quick Start

```powershell
# Connect to Mimecast
Connect-Mimecast -Region 'US' -Setup

# Search archived messages
Search-MimecastMessage -FromAddress 'user@example.com' -Subject '*report*'

# Track message delivery
Get-MimecastMessageTrace -ToAddress 'recipient@example.com' -Status 'delivered'

# Get message details
Get-MimecastMessageInfo -Id 'message-id'

# Check quarantined messages
Get-MimecastHeldMessage -Route 'inbound'
```

## Documentation

- [User Guide](Docs/UserGuide.md) - Detailed usage instructions and examples
- [API Reference](https://www.mimecast.com/developer/) - Mimecast API documentation
- [Contributing](CONTRIBUTING.md) - Guidelines for contributing to the project

## Requirements

- PowerShell 7.0 or later
- Mimecast account with API access
- Required PowerShell modules:
  * Microsoft.PowerShell.SecretManagement
  * Microsoft.PowerShell.SecretStore

## Features

### Message Management

- `Search-MimecastMessage` - Search archived messages
- `Get-MimecastMessageTrace` - Track message delivery
- `Get-MimecastMessageInfo` - Get detailed message information
- `Get-MimecastHeldMessage` - Search quarantined messages
- `Get-MimecastProcessingMessage` - View messages being processed

### Connection Management

- `Connect-Mimecast` - Establish API connection
- `Disconnect-Mimecast` - Terminate API connection
- `Get-MimecastConfiguration` - View current settings
- `Set-MimecastConfiguration` - Update module settings

### Security

- Secure credential storage using SecretManagement
- OAuth2 authentication support
- Automatic token refresh
- Region-specific endpoints

## Examples

### Message Search

```powershell
# Search with multiple criteria
Search-MimecastMessage `
    -FromAddress '*@example.com' `
    -Subject '*confidential*' `
    -HasAttachment `
    -StartDate (Get-Date).AddDays(-7)

# Process results
Search-MimecastMessage -Subject 'report' |
    Where-Object { $_.Size -gt 10MB } |
    Select-Object Subject, From, To, Size |
    Export-Csv 'large-reports.csv'
```

### Message Tracking

```powershell
# Track failed deliveries
Get-MimecastMessageTrace `
    -Status 'failed' `
    -StartDate (Get-Date).AddHours(-1) |
    ForEach-Object {
        Write-Warning "Delivery failed: $($_.Subject)"
        $_.ProcessingDetails | Format-List
    }

# Monitor specific recipients
Get-MimecastMessageTrace `
    -ToAddress 'vip@example.com' `
    -Route 'inbound' |
    Where-Object { $_.Status -ne 'delivered' }
```

### Batch Processing

```powershell
# Process messages in batches
Search-MimecastMessage -Subject 'invoice' |
    ForEach-Object -Begin {
        $batch = @()
        $count = 0
    } -Process {
        $batch += $_
        $count++
        if ($count % 100 -eq 0) {
            Process-MessageBatch $batch
            $batch = @()
        }
    } -End {
        if ($batch) {
            Process-MessageBatch $batch
        }
    }
```

## Contributing

Contributions are welcome! Please see our [Contributing Guide](CONTRIBUTING.md) for details.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Support

- For bugs and feature requests, please create an issue
- For questions and discussions, please use GitHub Discussions
- For security issues, please see our [Security Policy](SECURITY.md)