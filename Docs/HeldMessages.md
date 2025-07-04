# Held Messages (Quarantine) Management

The Mimecast PowerShell module provides comprehensive cmdlets for managing held (quarantined) messages, allowing you to search, review, release, and block messages that have been held by Mimecast's security policies.

## Available Cmdlets

- `Get-MimecastHeldMessage`: Search and list held messages
- `Release-MimecastHeldMessage`: Release messages from quarantine
- `Block-MimecastHeldMessage`: Block held messages
- `Get-MimecastHeldMessageContent`: View message content
- `Save-MimecastHeldMessageAttachment`: Download message attachments
- `Get-MimecastHeldMessageHeader`: View message headers

## Basic Usage

### Search Held Messages

```powershell
# Get all held messages
Get-MimecastHeldMessage

# Filter by route
Get-MimecastHeldMessage -Route 'inbound'

# Filter by hold type
Get-MimecastHeldMessage -HoldType 'spam'

# Search by sender
Get-MimecastHeldMessage -Sender '*@suspiciousdomain.com'

# Search by recipient
Get-MimecastHeldMessage -Recipient 'user@yourdomain.com'

# Search by subject
Get-MimecastHeldMessage -Subject '*invoice*'

# Combine filters
Get-MimecastHeldMessage `
    -Route 'inbound' `
    -HoldType 'malware_sandbox' `
    -StartDate (Get-Date).AddDays(-7)
```

### Release Messages

```powershell
# Release a single message
Release-MimecastHeldMessage -Id 'eNo1j8EKgjAURf9l1wSnZVnQwdIoUjSmZuCpOWWp23JpWfTv6aH7e7z3AZLg'

# Release with custom reason
Release-MimecastHeldMessage `
    -Id 'eNo1j8EKgjAURf9l1wSnZVnQwdIoUjSmZuCpOWWp23JpWfTv6aH7e7z3AZLg' `
    -Reason 'Approved by security team'

# Release with tag
Release-MimecastHeldMessage `
    -Id 'eNo1j8EKgjAURf9l1wSnZVnQwdIoUjSmZuCpOWWp23JpWfTv6aH7e7z3AZLg' `
    -Action 'release_with_tag'

# Release multiple messages
Get-MimecastHeldMessage -HoldType 'suspected_spam' |
    Release-MimecastHeldMessage -Reason 'Bulk release - false positives'
```

### Block Messages

```powershell
# Block a single message
Block-MimecastHeldMessage -Id 'eNo1j8EKgjAURf9l1wSnZVnQwdIoUjSmZuCpOWWp23JpWfTv6aH7e7z3AZLg'

# Block with custom reason
Block-MimecastHeldMessage `
    -Id 'eNo1j8EKgjAURf9l1wSnZVnQwdIoUjSmZuCpOWWp23JpWfTv6aH7e7z3AZLg' `
    -Reason 'Confirmed malicious content'

# Block multiple messages
Get-MimecastHeldMessage -HoldType 'malware_sandbox' |
    Block-MimecastHeldMessage -Reason 'Malware detected'
```

### View Message Content

```powershell
# View HTML content
Get-MimecastHeldMessageContent `
    -Id 'eNo1j8EKgjAURf9l1wSnZVnQwdIoUjSmZuCpOWWp23JpWfTv6aH7e7z3AZLg' `
    -Type 'html'

# View plain text
Get-MimecastHeldMessageContent `
    -Id 'eNo1j8EKgjAURf9l1wSnZVnQwdIoUjSmZuCpOWWp23JpWfTv6aH7e7z3AZLg' `
    -Type 'text'
```

### Download Attachments

```powershell
# Download a specific attachment
Save-MimecastHeldMessageAttachment `
    -Id 'eNo1j8EKgjAURf9l1wSnZVnQwdIoUjSmZuCpOWWp23JpWfTv6aH7e7z3AZLg' `
    -AttachmentId 'att123' `
    -OutputPath 'C:\Temp\attachment.pdf'

# Force overwrite existing file
Save-MimecastHeldMessageAttachment `
    -Id 'eNo1j8EKgjAURf9l1wSnZVnQwdIoUjSmZuCpOWWp23JpWfTv6aH7e7z3AZLg' `
    -AttachmentId 'att123' `
    -OutputPath 'C:\Temp\attachment.pdf' `
    -Force
```

### View Headers

```powershell
# Get message headers
Get-MimecastHeldMessageHeader -Id 'eNo1j8EKgjAURf9l1wSnZVnQwdIoUjSmZuCpOWWp23JpWfTv6aH7e7z3AZLg'

# Process multiple messages
Get-MimecastHeldMessage -HoldType 'dmarc' |
    Get-MimecastHeldMessageHeader
```

## Advanced Usage

### Batch Processing

```powershell
# Process messages in batches
Get-MimecastHeldMessage -HoldType 'spam' |
    ForEach-Object -Begin {
        $batch = @()
        $count = 0
    } -Process {
        $batch += $_
        $count++
        if ($count % 10 -eq 0) {
            Process-MessageBatch $batch
            $batch = @()
        }
    } -End {
        if ($batch) {
            Process-MessageBatch $batch
        }
    }
```

### Reporting

```powershell
# Generate hold type summary
Get-MimecastHeldMessage -StartDate (Get-Date).AddDays(-30) |
    Group-Object holdType |
    Select-Object @{
        Name = 'HoldType'
        Expression = { $_.Name }
    }, @{
        Name = 'Count'
        Expression = { $_.Count }
    } |
    Export-Csv 'quarantine-summary.csv'

# Track release/block actions
$messages = Get-MimecastHeldMessage -HoldType 'malware_sandbox'
foreach ($msg in $messages) {
    try {
        Block-MimecastHeldMessage -Id $msg.id -ErrorAction Stop
        "$(Get-Date) - Blocked: $($msg.subject)" |
            Add-Content 'quarantine-actions.log'
    }
    catch {
        "$(Get-Date) - Failed to block: $($msg.subject) - $_" |
            Add-Content 'quarantine-actions.log'
    }
}
```

### Error Handling

```powershell
# Handle release errors
try {
    $messages = Get-MimecastHeldMessage -HoldType 'suspected_spam'
    foreach ($msg in $messages) {
        try {
            $result = Release-MimecastHeldMessage `
                -Id $msg.id `
                -ErrorAction Stop

            Write-Host "Released: $($msg.subject)"
        }
        catch {
            Write-Warning "Failed to release $($msg.subject): $_"
            # Log error or implement retry logic
        }
    }
}
catch {
    Write-Error "Failed to get held messages: $_"
}
```

## Best Practices

1. Always provide meaningful reasons for release/block actions
2. Implement proper error handling for batch operations
3. Log all release and block actions
4. Review message content and headers before taking action
5. Use appropriate hold type filters to narrow down results
6. Consider implementing approval workflows for sensitive actions
7. Regularly review and clean up quarantine

## Required Permissions

The held message cmdlets require the following Mimecast API permissions:

- "Gateway > Hold Queue > Read"
- "Gateway > Hold Queue > Write"

Ensure your API application has these permissions assigned in the Mimecast Administration Console.

## Related Cmdlets

- `Search-MimecastMessage`: Search archived messages
- `Get-MimecastMessageTrace`: Track message delivery
- `Get-MimecastMessageInfo`: Get detailed message information