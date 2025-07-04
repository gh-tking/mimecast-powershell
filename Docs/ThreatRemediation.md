# Threat Remediation

The Mimecast PowerShell module provides cmdlets for managing threat remediations, allowing you to automate the process of handling identified threats in your email environment.

## Available Cmdlets

- `Get-MimecastThreatRemediation`: List threat remediations
- `Start-MimecastThreatRemediation`: Create a new threat remediation
- `Get-MimecastThreatRemediationStatus`: Get status of a specific remediation
- `Stop-MimecastThreatRemediation`: Cancel an in-progress remediation

## Basic Usage

### List Threat Remediations

```powershell
# Get all remediations
Get-MimecastThreatRemediation

# Get pending remediations
Get-MimecastThreatRemediation -Status 'pending'

# Get remediations from last 7 days
Get-MimecastThreatRemediation -StartDate (Get-Date).AddDays(-7)

# Get all remediations with automatic pagination
Get-MimecastThreatRemediation -All
```

### Start a Remediation

```powershell
# Delete a malicious message
Start-MimecastThreatRemediation `
    -Code 'eNo1j8EKgjAURf9l1wSnZVnQwdIoUjSmZuCpOWWp23JpWfTv6aH7e7z3AZLg' `
    -Action 'delete_message' `
    -Reason 'Malicious content detected'

# Restore a false positive
Start-MimecastThreatRemediation `
    -Code 'fMo2k9FLhkBURf9l1wSnZVnQwdIoUjSmZuCpOWWp23JpWfTv6aH7e7z3AZLg' `
    -Action 'restore_message' `
    -Reason 'False positive confirmed'
```

### Check Remediation Status

```powershell
# Get status of a specific remediation
Get-MimecastThreatRemediationStatus -Code 'eNo1j8EKgjAURf9l1wSnZVnQwdIoUjSmZuCpOWWp23JpWfTv6aH7e7z3AZLg'

# Monitor multiple remediations
Get-MimecastThreatRemediation -Status 'pending' | 
    Get-MimecastThreatRemediationStatus
```

### Stop a Remediation

```powershell
# Cancel a specific remediation
Stop-MimecastThreatRemediation `
    -Code 'eNo1j8EKgjAURf9l1wSnZVnQwdIoUjSmZuCpOWWp23JpWfTv6aH7e7z3AZLg' `
    -Reason 'No longer needed'

# Cancel multiple pending remediations
Get-MimecastThreatRemediation -Status 'pending' |
    Stop-MimecastThreatRemediation -Reason 'Bulk cancellation'
```

## Advanced Usage

### Batch Processing

```powershell
# Process multiple remediations in batches
Get-MimecastThreatRemediation -Status 'pending' |
    ForEach-Object -Begin {
        $batch = @()
        $count = 0
    } -Process {
        $batch += $_
        $count++
        if ($count % 10 -eq 0) {
            Process-RemediationBatch $batch
            $batch = @()
        }
    } -End {
        if ($batch) {
            Process-RemediationBatch $batch
        }
    }
```

### Monitoring and Reporting

```powershell
# Monitor remediation progress
$remediations = Get-MimecastThreatRemediation -Status 'pending'
$remediations | ForEach-Object {
    $status = Get-MimecastThreatRemediationStatus -Code $_.code
    if ($status.status -eq 'completed') {
        Write-Host "Remediation $($_.code) completed successfully"
    }
    elseif ($status.status -eq 'failed') {
        Write-Warning "Remediation $($_.code) failed: $($status.failureReason)"
    }
}

# Generate remediation report
Get-MimecastThreatRemediation -StartDate (Get-Date).AddDays(-30) |
    Group-Object status |
    Select-Object @{
        Name = 'Status'
        Expression = { $_.Name }
    }, @{
        Name = 'Count'
        Expression = { $_.Count }
    } |
    Export-Csv 'remediation-report.csv'
```

### Error Handling

```powershell
# Handle remediation errors
try {
    $remediation = Start-MimecastThreatRemediation `
        -Code $messageCode `
        -Action 'delete_message' `
        -ErrorAction Stop

    do {
        Start-Sleep -Seconds 5
        $status = Get-MimecastThreatRemediationStatus -Code $messageCode
    } while ($status.status -eq 'pending')

    if ($status.status -eq 'failed') {
        throw "Remediation failed: $($status.failureReason)"
    }
}
catch {
    Write-Error "Error processing remediation: $_"
    # Implement retry logic or notification
}
```

## Best Practices

1. Always provide a meaningful reason for remediations
2. Monitor long-running remediations
3. Implement proper error handling
4. Use batching for bulk operations
5. Keep detailed logs of remediation actions
6. Regularly review failed remediations
7. Consider implementing approval workflows for critical actions

## Required Permissions

The threat remediation cmdlets require the following Mimecast API permissions:

- "Remediation > Remediation > Read"
- "Remediation > Remediation > Write"

Ensure your API application has these permissions assigned in the Mimecast Administration Console.

## Related Cmdlets

- `Search-MimecastMessage`: Search for messages that might need remediation
- `Get-MimecastMessageInfo`: Get detailed message information
- `Get-MimecastHeldMessage`: Search quarantined messages that might need remediation