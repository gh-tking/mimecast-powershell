# Mimecast PowerShell Module User Guide

## Table of Contents
- [Installation](#installation)
- [Getting Started](#getting-started)
- [Authentication](#authentication)
- [Archive Message Search](#archive-message-search)
- [Message Management](#message-management)
- [Advanced Usage](#advanced-usage)
- [Troubleshooting](#troubleshooting)

## Installation

```powershell
# Install from PowerShell Gallery
Install-Module -Name Mimecast -Scope CurrentUser

# Import the module
Import-Module Mimecast
```

## Getting Started

Before using the module, you'll need to connect to the Mimecast API. The recommended secure approach is to use the `-Setup` parameter:

```powershell
# First-time setup and connection (recommended)
Connect-Mimecast -Setup -Region 'US'  # or EU, DE, CA, ZA, AU, Offshore
```

This will:
1. Initialize the secure credential store
2. Securely prompt for your API credentials
3. Store the credentials in an encrypted vault
4. Establish the API connection

For subsequent connections, simply use:
```powershell
Connect-Mimecast -Region 'US'
```

## Authentication

The module uses Microsoft.PowerShell.SecretManagement and Microsoft.PowerShell.SecretStore to securely manage your API credentials:

- Credentials are never exposed in scripts or command history
- All secrets are encrypted at rest
- Access is controlled through the SecretStore vault

### Security Best Practices

1. NEVER:
   - Include API keys or secrets in scripts
   - Pass credentials as command line parameters
   - Store credentials in source control
   - Log credentials or secrets
   - Include credentials in documentation

2. ALWAYS:
   - Use `-Setup` for initial configuration
   - Let the module prompt for credentials
   - Store credentials in the SecretStore vault
   - Use environment-specific credential names for different environments

### Managing Multiple Environments

For scenarios like CI/CD or multiple accounts:

```powershell
# Connect with environment-specific credentials
Connect-Mimecast `
    -ClientIdName 'Prod_ClientId' `
    -ClientSecretName 'Prod_ClientSecret' `
    -Region 'US'
```

### Rotating Credentials

To update stored credentials:

```powershell
# Securely update credentials
Connect-Mimecast -Setup -Region 'US' -ForceSetup
```

## Archive Message Search

The `Search-MimecastMessage` cmdlet provides powerful search capabilities across your Mimecast email archive. It supports both simple filtering parameters for common searches and advanced query syntax for complex requirements.

### Basic Usage

1. Simple Search Examples:
   ```powershell
   # Search by sender
   Search-MimecastMessage -FromAddress 'user@example.com'

   # Search by subject
   Search-MimecastMessage -Subject 'monthly report'

   # Search with multiple criteria
   Search-MimecastMessage `
       -FromAddress '*@example.com' `
       -Subject '*confidential*' `
       -HasAttachment
   ```

2. Date Range Filtering:
   ```powershell
   # Last 30 days
   Search-MimecastMessage `
       -StartDate (Get-Date).AddDays(-30) `
       -Subject 'invoice'

   # Specific date range
   Search-MimecastMessage `
       -StartDate '2024-01-01' `
       -EndDate '2024-01-31' `
       -FromAddress 'finance@example.com'
   ```

3. Attachment Searches:
   ```powershell
   # Messages with attachments
   Search-MimecastMessage -HasAttachment

   # Specific file types
   Search-MimecastMessage `
       -AttachmentFileName '*.pdf' `
       -FromAddress '*@example.com'

   # By file hash
   Search-MimecastMessage -AttachmentFileHash '1234...abcd'
   ```

### Advanced Query Syntax

For complex searches, use the -Query parameter with Mimecast's query syntax:

```powershell
# Complex sender/recipient combination
Search-MimecastMessage -Query '(from:"finance@example.com" OR from:"accounting@example.com") AND to:"executives@example.com"'

# Content and date conditions
Search-MimecastMessage -Query 'subject:"quarterly report" AND date:[2024-01-01 TO 2024-03-31] AND has:attachment'

# Multiple conditions with grouping
Search-MimecastMessage -Query '(subject:"confidential" OR subject:"sensitive") AND NOT from:"external@domain.com"'
```

### Pagination and Result Management

1. Controlling Result Size:
   ```powershell
   # Get first N results
   Search-MimecastMessage -Subject 'report' -First 100

   # Skip and take (for manual paging)
   Search-MimecastMessage -Subject 'report' -Skip 100 -First 100

   # Get all results (automatic pagination)
   Search-MimecastMessage -Subject 'report' -All
   ```

2. Custom Page Size:
   ```powershell
   # Adjust page size for performance
   Search-MimecastMessage -Subject 'report' -PageSize 500
   ```

### Working with Results

1. Pipeline Integration:
   ```powershell
   # Get detailed info for search results
   Search-MimecastMessage -Subject 'urgent' |
       Get-MimecastMessageInfo

   # Export to CSV
   Search-MimecastMessage -FromAddress 'user@example.com' |
       Export-Csv -Path 'messages.csv'
   ```

2. Filtering Results:
   ```powershell
   # Filter by size
   Search-MimecastMessage -HasAttachment |
       Where-Object { $_.Size -gt 10MB }

   # Filter by date
   Search-MimecastMessage -Subject 'report' |
       Where-Object { $_.Received.DayOfWeek -eq 'Monday' }
   ```

### Performance Considerations

1. Query Optimization:
   - Use specific search criteria when possible
   - Limit date ranges for faster results
   - Use indexed fields (subject, from, to) before body content
   - Combine multiple conditions effectively

2. Resource Management:
   - Use -First when you don't need all results
   - Adjust -PageSize based on result size and performance
   - Consider using -Skip/-First for large result sets

### Best Practices

1. Error Handling:
   ```powershell
   try {
       $messages = Search-MimecastMessage -Subject 'important' -ErrorAction Stop
   }
   catch {
       Write-Error "Search failed: $_"
   }
   ```

2. Parameter Validation:
   ```powershell
   # Validate date ranges
   if ($EndDate -lt $StartDate) {
       throw "End date must be after start date"
   }

   # Validate attachment parameters
   if ($HasAttachment -and $NoAttachment) {
       throw "Cannot specify both HasAttachment and NoAttachment"
   }
   ```

3. Result Processing:
   ```powershell
   # Process results in batches
   Search-MimecastMessage -Subject 'report' |
       ForEach-Object -Begin {
           $count = 0
           $batch = @()
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

### Common Issues and Solutions

1. Query Timeout:
   ```powershell
   # Problem: Search taking too long
   # Solution: Narrow the date range and use more specific criteria
   Search-MimecastMessage `
       -StartDate (Get-Date).AddDays(-7) `
       -FromAddress 'user@example.com' `
       -Subject 'specific topic'
   ```

2. Too Many Results:
   ```powershell
   # Problem: Memory usage with large result sets
   # Solution: Process in batches
   Search-MimecastMessage -Subject 'report' -First 1000 |
       ForEach-Object {
           # Process each message
           $_
       }
   ```

3. No Results Found:
   ```powershell
   # Problem: Search criteria too restrictive
   # Solution: Broaden search terms and use wildcards
   Search-MimecastMessage `
       -FromAddress '*@example.com' `
       -Subject '*report*'
   ```

### API Permissions

The Search-MimecastMessage cmdlet requires the following Mimecast API permissions:
- "Email Archive > Archive Search > Read"

Ensure your API keys have these permissions assigned in the Mimecast Administration Console.

### Related Cmdlets

- `Get-MimecastMessageInfo`: Get detailed information about specific messages
- `Get-MimecastHeldMessage`: Search quarantined messages
- `Get-MimecastProcessingMessage`: Search messages currently being processed

## Message Management

### Message Tracking (Get-MessageTrace)

The `Get-MessageTrace` cmdlet provides access to Mimecast's Track and Trace functionality, allowing you to monitor message flow through the Mimecast service. This is distinct from Archive Search (`Search-MimecastMessage`) in several ways:

1. Purpose:
   - Track and Trace: Real-time message tracking and delivery status
   - Archive Search: Historical message content search and retrieval

2. Data Retention:
   - Track and Trace: Typically 30 days
   - Archive Search: Based on your archiving configuration (often years)

3. Data Available:
   - Track and Trace: Detailed processing events, policy matches, delivery status
   - Archive Search: Message content, attachments, metadata

#### Basic Usage

1. Simple Tracking:
   ```powershell
   # Track messages from a sender
   Get-MessageTrace -FromEmailAddress 'user@example.com'

   # Track messages to a recipient
   Get-MessageTrace -ToEmailAddress 'recipient@example.com'

   # Track by subject
   Get-MessageTrace -Subject '*invoice*'
   ```

2. Route and Status Filtering:
   ```powershell
   # Track outbound messages
   Get-MessageTrace -Route 'outbound'

   # Track delivered messages
   Get-MessageTrace -Status 'delivered'

   # Combine filters
   Get-MessageTrace `
       -Route 'outbound' `
       -Status 'failed' `
       -FromEmailAddress '*@example.com'
   ```

3. Date Range Filtering:
   ```powershell
   # Last 24 hours
   Get-MessageTrace `
       -StartDate (Get-Date).AddHours(-24) `
       -EndDate (Get-Date)

   # Specific date range
   Get-MessageTrace `
       -StartDate '2024-01-01T00:00:00Z' `
       -EndDate '2024-01-31T23:59:59Z'
   ```

#### Advanced Usage

1. Message Flow Analysis:
   ```powershell
   # Get processing details for failed messages
   Get-MessageTrace -Status 'failed' |
       Select-Object Subject, From, To, ProcessingDetails |
       Format-List

   # Analyze policy matches
   Get-MessageTrace -Route 'inbound' |
       Where-Object { $_.PolicyMatches.Count -gt 0 } |
       Select-Object Subject, @{
           Name = 'Policies'
           Expression = { $_.PolicyMatches.PolicyName -join ', ' }
       }
   ```

2. Delivery Status Monitoring:
   ```powershell
   # Check recent delivery failures
   Get-MessageTrace `
       -Status 'failed' `
       -StartDate (Get-Date).AddHours(-1) |
       Select-Object Subject, From, To, ProcessingDetails |
       Export-Csv 'delivery-failures.csv'

   # Monitor specific recipient
   Get-MessageTrace `
       -ToEmailAddress 'vip@example.com' `
       -StartDate (Get-Date).AddMinutes(-30) |
       Where-Object { $_.Status -ne 'delivered' }
   ```

3. Result Management:
   ```powershell
   # Get oldest messages first
   Get-MessageTrace `
       -StartDate (Get-Date).AddDays(-7) `
       -OldestFirst

   # Paginate results manually
   Get-MessageTrace `
       -First 100 `
       -Skip 0

   Get-MessageTrace `
       -First 100 `
       -Skip 100
   ```

#### Performance Considerations

1. Date Range:
   - Keep date ranges narrow for better performance
   - Consider using multiple smaller queries instead of one large query
   - Remember the 30-day data retention limit

2. Filter Optimization:
   - Use specific filters when possible
   - Combine filters to reduce result set
   - Consider using -First to limit results

3. Resource Usage:
   - Process results in batches
   - Export large result sets to CSV
   - Use pagination for manual processing

#### Best Practices

1. Error Handling:
   ```powershell
   try {
       $messages = Get-MessageTrace `
           -FromEmailAddress 'user@example.com' `
           -ErrorAction Stop
   }
   catch {
       Write-Error "Message trace failed: $_"
   }
   ```

2. Result Processing:
   ```powershell
   # Process in batches
   Get-MessageTrace -Status 'failed' |
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

3. Monitoring:
   ```powershell
   # Create monitoring script
   $monitorParams = @{
       Route = 'inbound'
       Status = 'failed'
       StartDate = (Get-Date).AddMinutes(-5)
       EndDate = Get-Date
   }

   while ($true) {
       $failures = Get-MessageTrace @monitorParams
       if ($failures) {
           Send-Alert -Messages $failures
       }
       Start-Sleep -Seconds 300
   }
   ```

#### Common Issues and Solutions

1. No Results:
   ```powershell
   # Problem: Date range too narrow
   # Solution: Expand date range
   Get-MessageTrace `
       -StartDate (Get-Date).AddHours(-24) `
       -EndDate (Get-Date)

   # Problem: Too specific filters
   # Solution: Use wildcards
   Get-MessageTrace `
       -FromEmailAddress '*@example.com' `
       -Subject '*report*'
   ```

2. Performance Issues:
   ```powershell
   # Problem: Large date range
   # Solution: Split into smaller queries
   $start = (Get-Date).AddDays(-7)
   $end = Get-Date
   $interval = New-TimeSpan -Days 1

   while ($start -lt $end) {
       $nextEnd = $start + $interval
       if ($nextEnd -gt $end) { $nextEnd = $end }

       Get-MessageTrace `
           -StartDate $start `
           -EndDate $nextEnd |
           Process-Messages

       $start = $nextEnd
   }
   ```

3. Missing Data:
   ```powershell
   # Problem: Message not found
   # Solution: Check archive instead
   if (-not (Get-MessageTrace -MessageId $id)) {
       Search-MimecastMessage -MessageId $id
   }
   ```

#### API Permissions

The Get-MessageTrace cmdlet requires the following Mimecast API permissions:
- "Message > Track and Trace > Read"

Ensure your API keys have these permissions assigned in the Mimecast Administration Console.

#### Related Cmdlets

- `Search-MimecastMessage`: Search archived messages
- `Get-MimecastMessageInfo`: Get detailed message information
- `Get-MimecastHeldMessage`: Search quarantined messages

### Message Archive Search

[Previous Archive Search section content...]

## Advanced Usage

[Additional sections...]

## Troubleshooting

[Additional sections...]