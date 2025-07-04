<#
.SYNOPSIS
Gets messages held in Mimecast quarantine.

.DESCRIPTION
Retrieves messages that have been held in Mimecast's quarantine for review. These
messages may have been held due to content policies, spam detection, or other
security measures. The function supports filtering by route and date range.

.PARAMETER Route
Optional. Filter messages by route. Valid values:
- inbound: Messages coming into the organization
- outbound: Messages going out of the organization
- internal: Messages between internal users
- all: All routes (default)

.PARAMETER StartDate
Optional. Start date for message search.
Default is 7 days ago.

.PARAMETER EndDate
Optional. End date for message search.
Default is current date/time.

.PARAMETER FirstPageOnly
Optional. Return only the first page of results.
Default is $false (returns all pages).

.PARAMETER PageSize
Optional. Number of messages to return per page.
Default is 100. Maximum is 500.

.PARAMETER Sender
Optional. Filter by sender email address.
Supports wildcards (e.g., '*@example.com').

.PARAMETER Recipient
Optional. Filter by recipient email address.
Supports wildcards (e.g., '*@example.com').

.PARAMETER Subject
Optional. Filter by message subject.
Supports wildcards (e.g., '*confidential*').

.PARAMETER HoldReason
Optional. Filter by reason message was held. Valid values:
- policy: Content policy match
- spam: Spam detection
- malware: Malware detection
- attachment: Attachment policy match
- all: All reasons (default)

.INPUTS
None. You cannot pipe objects to Get-MimecastHeldMessage.

.OUTPUTS
System.Object[]
Returns an array of held message objects. Each object includes:
- id: Unique message identifier
- sender: Sender email address
- recipient: Recipient email address
- subject: Message subject
- received: Date/time message was received
- holdReason: Why the message was held
- route: Message route (inbound/outbound/internal)
- size: Message size in bytes
- hasAttachments: Whether message has attachments
- holdDate: Date/time message was held
- releaseDate: Date/time message was released (if applicable)
- status: Current message status

.EXAMPLE
C:\> Get-MimecastHeldMessage
Gets all held messages from the last 7 days.

.EXAMPLE
C:\> Get-MimecastHeldMessage -Route 'outbound' -HoldReason 'policy'
Gets outbound messages held due to policy violations.

.EXAMPLE
C:\> $startDate = (Get-Date).AddDays(-30)
C:\> $endDate = Get-Date
C:\> Get-MimecastHeldMessage `
    -StartDate $startDate `
    -EndDate $endDate `
    -Sender '*@example.com' `
    -Subject '*confidential*'

Gets held messages from the last 30 days from a specific domain
with 'confidential' in the subject.

.EXAMPLE
C:\> Get-MimecastHeldMessage -FirstPageOnly -PageSize 50
Gets only the first 50 held messages.

.NOTES
- Results are paginated for performance
- Date range affects search performance
- Some fields may be empty if not applicable
- API permissions required:
  * "Message > Message Information > Read"
  * "Message > Message Hold > Read"
#>
function Get-MimecastHeldMessage {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]$Id,

        [Parameter()]
        [string]$FromAddress,

        [Parameter()]
        [string]$ToAddress,

        [Parameter()]
        [string]$Subject,

        [Parameter()]
        [ValidateSet('inbound', 'outbound')]
        [string]$Route,

        [Parameter()]
        [switch]$FirstPageOnly
    )

    try {
        $body = @{
            data = @(
                @{}
            )
        }

        if ($Id) {
            $body.data[0].id = $Id
        }
        if ($FromAddress) {
            $body.data[0].fromAddress = $FromAddress
        }
        if ($ToAddress) {
            $body.data[0].toAddress = $ToAddress
        }
        if ($Subject) {
            $body.data[0].subject = $Subject
        }
        if ($Route) {
            $body.data[0].route = $Route
        }

        # Use pagination helper
        Invoke-ApiRequestWithPagination -Path '/api/gateway/get-hold-message-list' `
            -InitialBody $body `
            -ReturnFirstPageOnly:$FirstPageOnly
    }
    catch {
        Write-Error -ErrorRecord $_
    }
}

<#
.SYNOPSIS
Gets detailed information about a specific Mimecast message.

.DESCRIPTION
Retrieves detailed information about a message processed by Mimecast, including
tracking data, policy matches, and security scan results. The function supports
pipeline input from other message-related cmdlets.

.PARAMETER Id
The unique identifier of the message. This parameter accepts pipeline input
by property name from Get-MimecastHeldMessage and Search-MimecastMessage.

.INPUTS
System.Object. You can pipe message objects from Get-MimecastHeldMessage or Search-MimecastMessage.

.OUTPUTS
System.Object
Returns a detailed message object that includes:
- id: Unique message identifier
- sender: Sender information (email, name, domain)
- recipient: Recipient information (email, name, domain)
- subject: Message subject
- received: Date/time message was received
- size: Message size in bytes
- hasAttachments: Whether message has attachments
- attachments: Array of attachment details (if present)
- headers: Message headers
- route: Message route (inbound/outbound/internal)
- status: Current message status
- processingDetails: Array of processing events
  * timestamp: When event occurred
  * type: Type of event (policy, spam, malware, etc.)
  * result: Result of processing
  * details: Additional event details
- securityDetails: Security scan results
  * spamScore: Spam detection score
  * malwareDetected: Whether malware was found
  * urlsRewritten: Whether URLs were rewritten
  * attachmentsScanned: Whether attachments were scanned
- policyMatches: Array of policy matches
  * policyId: ID of matched policy
  * policyName: Name of matched policy
  * action: Action taken by policy
  * matchDetails: Why policy matched

.EXAMPLE
C:\> Get-MimecastMessageInfo -Id '4d24b932-5c57-4b57-b7b5-9c3f511e5c5c'
Gets detailed information about a specific message.

.EXAMPLE
C:\> Get-MimecastHeldMessage -Route 'outbound' |
    Where-Object { $_.holdReason -eq 'policy' } |
    Get-MimecastMessageInfo

Gets detailed information about all outbound messages held due to policy violations.

.EXAMPLE
C:\> Search-MimecastMessage -Subject '*confidential*' |
    Get-MimecastMessageInfo |
    Where-Object { $_.policyMatches.Count -gt 0 }

Gets detailed information about messages with 'confidential' in the subject
that matched any policies.

.NOTES
- Some fields may be empty if not applicable
- Processing details retention varies by configuration
- Security details depend on enabled features
- API permissions required:
  * "Message > Message Information > Read"
#>
function Get-MimecastMessageInfo {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$Id
    )

    try {
        $body = @{
            data = @(
                @{
                    id = $Id
                }
            )
        }

        $response = Invoke-ApiRequest -Method 'POST' -Path '/api/message-finder/get-message-info' -Body $body
        $response.data
    }
    catch {
        Write-Error -ErrorRecord $_
    }
}

<#
.SYNOPSIS
Searches for messages in the Mimecast archive.

.DESCRIPTION
Performs a search across messages in the Mimecast archive based on various criteria
such as sender, recipient, subject, body content, and attachments. The function
supports both simple filtering parameters and advanced Mimecast query syntax.

.PARAMETER FromAddress
Optional. Array of sender email addresses to search for.
Multiple addresses are combined with OR logic.
Supports wildcards (e.g., '*@example.com').

.PARAMETER ToAddress
Optional. Array of recipient email addresses to search for.
Multiple addresses are combined with OR logic.
Supports wildcards (e.g., '*@example.com').

.PARAMETER Subject
Optional. Array of subject keywords or phrases to search for.
Multiple terms are combined with OR logic.
Supports wildcards (e.g., '*confidential*').

.PARAMETER Body
Optional. Array of body content keywords or phrases to search for.
Multiple terms are combined with OR logic.
Supports wildcards (e.g., '*confidential*').

.PARAMETER Keyword
Optional. Array of keywords to search across all common fields.
Multiple terms are combined with OR logic.
Searches subject, body, sender, and recipient fields.

.PARAMETER HasAttachment
Optional. Search for messages with attachments.
Cannot be used with -NoAttachment.

.PARAMETER NoAttachment
Optional. Search for messages without attachments.
Cannot be used with -HasAttachment.

.PARAMETER AttachmentFileName
Optional. Search for specific attachment file names.
Supports wildcards (e.g., '*.pdf').

.PARAMETER AttachmentFileHash
Optional. Search for specific attachment file hashes.
Supports MD5, SHA1, and SHA256 hashes.

.PARAMETER Query
Optional. Raw Mimecast query string for advanced searches.
Cannot be used with simple filtering parameters.
Example: 'subject:"report" AND from:"user@domain.com"'

.PARAMETER StartDate
Optional. Start of the date range for the search.
Default is 7 days ago.

.PARAMETER EndDate
Optional. End of the date range for the search.
Default is current date/time.

.PARAMETER FolderId
Optional. Array of specific archive folder IDs to search.

.PARAMETER MessageId
Optional. Array of specific message IDs to retrieve.

.PARAMETER AccountId
Optional. Array of specific account IDs to search.

.PARAMETER All
Optional. Retrieve all matching results by handling pagination.
Default is $true if neither -First nor -Skip are specified.

.PARAMETER First
Optional. Retrieve only the first N results.
Cannot be used with -All.

.PARAMETER Skip
Optional. Skip the first N results.
Useful for manual pagination.

.INPUTS
None. You cannot pipe objects to Search-MimecastMessage.

.OUTPUTS
System.Object[]
Returns an array of message objects. Each object includes:
- id: Unique message identifier
- subject: Message subject
- from: Sender information
  * address: Email address
  * name: Display name
- to: Array of recipient information
  * address: Email address
  * name: Display name
- received: Date/time message was received
- size: Message size in bytes
- hasAttachments: Whether message has attachments
- attachmentCount: Number of attachments
- status: Message status
- folder: Archive folder information
- score: Search relevance score

.EXAMPLE
C:\> Search-MimecastMessage -FromAddress 'user@example.com' -Subject 'report'
Searches for messages from a specific sender containing 'report' in the subject.

.EXAMPLE
C:\> Search-MimecastMessage `
    -Keyword 'confidential', 'urgent' `
    -HasAttachment `
    -StartDate (Get-Date).AddDays(-30) `
    -First 100

Searches for the first 100 messages from the last 30 days containing either
'confidential' or 'urgent' and having attachments.

.EXAMPLE
C:\> Search-MimecastMessage `
    -Query 'subject:"monthly report" AND (from:"finance@example.com" OR from:"accounting@example.com")' `
    -StartDate '2024-01-01' `
    -EndDate '2024-01-31'

Uses advanced query syntax to search for monthly reports from finance or accounting
in January 2024.

.EXAMPLE
C:\> Search-MimecastMessage `
    -ToAddress '*@example.com' `
    -AttachmentFileName '*.pdf', '*.docx' `
    -All

Searches for all messages sent to example.com with PDF or Word attachments.

.NOTES
- Date ranges affect search performance
- Complex queries may take longer to execute
- Some features require specific permissions
- Results are sorted by relevance by default
- API permissions required:
  * "Email Archive > Archive Search > Read"
#>
function Search-MimecastMessage {
    [CmdletBinding(DefaultParameterSetName='SimpleSearch')]
    param (
        [Parameter(ParameterSetName='SimpleSearch')]
        [string[]]$FromAddress,

        [Parameter(ParameterSetName='SimpleSearch')]
        [string[]]$ToAddress,

        [Parameter(ParameterSetName='SimpleSearch')]
        [string[]]$Subject,

        [Parameter(ParameterSetName='SimpleSearch')]
        [string[]]$Body,

        [Parameter(ParameterSetName='SimpleSearch')]
        [string[]]$Keyword,

        [Parameter(ParameterSetName='SimpleSearch')]
        [switch]$HasAttachment,

        [Parameter(ParameterSetName='SimpleSearch')]
        [switch]$NoAttachment,

        [Parameter(ParameterSetName='SimpleSearch')]
        [string]$AttachmentFileName,

        [Parameter(ParameterSetName='SimpleSearch')]
        [string]$AttachmentFileHash,

        [Parameter(ParameterSetName='RawQuery', Mandatory=$true)]
        [string]$Query,

        [Parameter()]
        [datetime]$StartDate = (Get-Date).AddDays(-7),

        [Parameter()]
        [datetime]$EndDate = (Get-Date),

        [Parameter()]
        [string[]]$FolderId,

        [Parameter()]
        [string[]]$MessageId,

        [Parameter()]
        [string[]]$AccountId,

        [Parameter()]
        [switch]$All = $true,

        [Parameter()]
        [ValidateRange(1, [int]::MaxValue)]
        [int]$First,

        [Parameter()]
        [ValidateRange(0, [int]::MaxValue)]
        [int]$Skip = 0,

        [Parameter()]
        [ValidateRange(1, 500)]
        [int]$PageSize = 100
    )

    begin {
        # Validate parameter combinations
        if ($HasAttachment -and $NoAttachment) {
            throw "Cannot use both -HasAttachment and -NoAttachment parameters."
        }

        if ($First -and $All) {
            Write-Warning "-All parameter is ignored when -First is specified."
            $All = $false
        }

        # Set up pagination variables
        $startRow = $Skip
        $maxResults = if ($First) { $First } else { [int]::MaxValue }
        $results = [System.Collections.ArrayList]::new()
    }

    process {
        try {
            # Build the XML query
            $xmlQuery = [System.Xml.XmlDocument]::new()
            $xmlQuery.LoadXml('<?xml version="1.0"?><xmlquery trace="iql,muse"><metadata query-type="emailarchive" archive="true" active="false"><smartfolders/><return-fields><return-field>attachmentcount</return-field><return-field>status</return-field><return-field>subject</return-field><return-field>size</return-field><return-field>receiveddate</return-field><return-field>displayfrom</return-field><return-field>displayfromaddress</return-field><return-field>id</return-field><return-field>displayto</return-field><return-field>displaytoaddresslist</return-field><return-field>smash</return-field></return-fields></metadata><muse></muse></xmlquery>')

            # Add query conditions
            $museNode = $xmlQuery.SelectSingleNode("//muse")
            if ($PSCmdlet.ParameterSetName -eq 'RawQuery') {
                $textNode = $xmlQuery.CreateElement("text")
                $textNode.InnerText = $Query
                $museNode.AppendChild($textNode)
            }
            else {
                $conditions = [System.Collections.ArrayList]::new()

                # Add from conditions
                if ($FromAddress) {
                    $fromConditions = $FromAddress | ForEach-Object { "from:`"$_`"" }
                    $conditions.Add("(" + ($fromConditions -join " OR ") + ")")
                }

                # Add to conditions
                if ($ToAddress) {
                    $toConditions = $ToAddress | ForEach-Object { "to:`"$_`"" }
                    $conditions.Add("(" + ($toConditions -join " OR ") + ")")
                }

                # Add subject conditions
                if ($Subject) {
                    $subjectConditions = $Subject | ForEach-Object { "subject:`"$_`"" }
                    $conditions.Add("(" + ($subjectConditions -join " OR ") + ")")
                }

                # Add body conditions
                if ($Body) {
                    $bodyConditions = $Body | ForEach-Object { "body:`"$_`"" }
                    $conditions.Add("(" + ($bodyConditions -join " OR ") + ")")
                }

                # Add keyword conditions
                if ($Keyword) {
                    $keywordConditions = $Keyword | ForEach-Object { "`"$_`"" }
                    $conditions.Add("(" + ($keywordConditions -join " OR ") + ")")
                }

                # Add attachment conditions
                if ($HasAttachment) {
                    $conditions.Add("has:attachment")
                }
                if ($NoAttachment) {
                    $conditions.Add("NOT has:attachment")
                }
                if ($AttachmentFileName) {
                    $conditions.Add("attachment.name:`"$AttachmentFileName`"")
                }
                if ($AttachmentFileHash) {
                    $conditions.Add("attachment.hash:`"$AttachmentFileHash`"")
                }

                # Add folder conditions
                if ($FolderId) {
                    $folderConditions = $FolderId | ForEach-Object { "folder:`"$_`"" }
                    $conditions.Add("(" + ($folderConditions -join " OR ") + ")")
                }

                # Add message ID conditions
                if ($MessageId) {
                    $messageConditions = $MessageId | ForEach-Object { "id:`"$_`"" }
                    $conditions.Add("(" + ($messageConditions -join " OR ") + ")")
                }

                # Add account conditions
                if ($AccountId) {
                    $accountConditions = $AccountId | ForEach-Object { "account:`"$_`"" }
                    $conditions.Add("(" + ($accountConditions -join " OR ") + ")")
                }

                # Add date range conditions
                $conditions.Add("date:[" + $StartDate.ToString("yyyy-MM-dd") + " TO " + $EndDate.ToString("yyyy-MM-dd") + "]")

                # Combine all conditions
                $textNode = $xmlQuery.CreateElement("text")
                $textNode.InnerText = $conditions -join " AND "
                $museNode.AppendChild($textNode)
            }

            # Update metadata attributes
            $metadataNode = $xmlQuery.SelectSingleNode("//metadata")
            $metadataNode.SetAttribute("page-size", $PageSize.ToString())
            $metadataNode.SetAttribute("startrow", $startRow.ToString())

            # Initialize pagination
            $pageToken = $null
            $totalRetrieved = 0
            $hasMore = $true

            while ($hasMore -and $totalRetrieved -lt $maxResults) {
                # Prepare request body
                $body = @{
                    data = @(
                        @{
                            query = $xmlQuery.OuterXml
                            admin = $true
                        }
                    )
                }

                if ($pageToken) {
                    $body.data[0].pageToken = $pageToken
                }

                # Make API request
                $response = Invoke-ApiRequest -Method 'POST' -Path '/api/archive/search' -Body $body

                # Process results
                foreach ($message in $response.data) {
                    if ($totalRetrieved -ge $maxResults) {
                        $hasMore = $false
                        break
                    }

                    # Transform message object
                    $messageObj = [PSCustomObject]@{
                        Id = $message.id
                        Subject = $message.subject
                        From = [PSCustomObject]@{
                            Address = $message.displayfromaddress
                            Name = $message.displayfrom
                        }
                        To = @($message.displaytoaddresslist | ForEach-Object {
                            [PSCustomObject]@{
                                Address = $_
                                Name = $message.displayto
                            }
                        })
                        Received = [datetime]::Parse($message.receiveddate)
                        Size = $message.size
                        HasAttachments = $message.attachmentcount -gt 0
                        AttachmentCount = $message.attachmentcount
                        Status = $message.status
                        Score = $message.score
                    }

                    [void]$results.Add($messageObj)
                    $totalRetrieved++
                }

                # Check for more pages
                if ($response.meta.pagination.next) {
                    $pageToken = $response.meta.pagination.next
                    $startRow += $PageSize
                    $metadataNode.SetAttribute("startrow", $startRow.ToString())
                }
                else {
                    $hasMore = $false
                }

                # Stop if we're not retrieving all results
                if (-not $All) {
                    $hasMore = $false
                }
            }
        }
        catch {
            Write-Error -ErrorRecord $_
        }
    }

    end {
        return $results.ToArray()
    }
}

<#
.SYNOPSIS
Gets messages currently being processed by Mimecast.

.DESCRIPTION
Retrieves messages that are currently being processed by Mimecast's email security
services. These messages are in transit and have not yet been delivered or
quarantined. The function supports filtering by route and other criteria.

.PARAMETER Id
Optional. Filter by specific message ID.

.PARAMETER FromAddress
Optional. Filter by sender email address.
Supports wildcards (e.g., '*@example.com').

.PARAMETER ToAddress
Optional. Filter by recipient email address.
Supports wildcards (e.g., '*@example.com').

.PARAMETER Subject
Optional. Filter by message subject.
Supports wildcards (e.g., '*confidential*').

.PARAMETER Route
Optional. Filter messages by route. Valid values:
- inbound: Messages coming into the organization
- outbound: Messages going out of the organization

.PARAMETER FirstPageOnly
Optional. Return only the first page of results.
Default is $false (returns all pages).

.INPUTS
None. You cannot pipe objects to Get-MimecastProcessingMessage.

.OUTPUTS
System.Object[]
Returns an array of processing message objects. Each object includes:
- id: Unique message identifier
- sender: Sender email address
- recipient: Recipient email address
- subject: Message subject
- received: Date/time message was received
- route: Message route (inbound/outbound)
- size: Message size in bytes
- hasAttachments: Whether message has attachments
- processingStage: Current processing stage
- processingStatus: Current processing status
- processingTime: Time spent in processing
- nextAction: Next processing action
- queuePosition: Position in processing queue

.EXAMPLE
C:\> Get-MimecastProcessingMessage
Gets all messages currently being processed.

.EXAMPLE
C:\> Get-MimecastProcessingMessage -Route 'inbound' -FromAddress '*@example.com'
Gets inbound messages being processed from example.com.

.EXAMPLE
C:\> Get-MimecastProcessingMessage -Subject '*urgent*' -FirstPageOnly
Gets the first page of messages with 'urgent' in the subject.

.NOTES
- Results are real-time and may change quickly
- Processing time varies by message complexity
- Some fields may be empty if not applicable
- API permissions required:
  * "Message > Message Processing > Read"
#>
<#
.SYNOPSIS
Gets message tracking information from Mimecast.

.DESCRIPTION
Retrieves message tracking logs from Mimecast's Track and Trace functionality.
This cmdlet provides visibility into message flow through the Mimecast service,
including delivery status, processing details, and policy matches. Data is typically
available for up to 30 days.

.PARAMETER FromEmailAddress
Optional. Array of sender email addresses to search for.
Multiple addresses are combined with OR logic.
Supports wildcards (e.g., '*@example.com').

.PARAMETER ToEmailAddress
Optional. Array of recipient email addresses to search for.
Multiple addresses are combined with OR logic.
Supports wildcards (e.g., '*@example.com').

.PARAMETER Subject
Optional. Keywords or phrases in the subject line.
Supports wildcards (e.g., '*confidential*').

.PARAMETER Route
Optional. Message route. Valid values:
- inbound: Messages coming into the organization
- outbound: Messages going out of the organization
- internal: Messages between internal users
- all: All routes (default)

.PARAMETER Status
Optional. Message status. Valid values:
- accepted: Message accepted by Mimecast
- rejected: Message rejected by policy
- delivered: Message successfully delivered
- held: Message held in quarantine
- failed: Delivery failed
- processing: Currently being processed
- all: All statuses (default)

.PARAMETER MessageId
Optional. Specific internet message ID to search for.

.PARAMETER StartDate
Optional. Start of the date range for the search.
Default is 7 days ago.
Maximum lookback is typically 30 days.

.PARAMETER EndDate
Optional. End of the date range for the search.
Default is current date/time.

.PARAMETER OldestFirst
Optional. Sort results with oldest messages first.
Default is newest first.

.PARAMETER All
Optional. Retrieve all matching results by handling pagination.
Default is $true if neither -First nor -Skip are specified.

.PARAMETER First
Optional. Retrieve only the first N results.
Cannot be used with -All.

.PARAMETER Skip
Optional. Skip the first N results.
Useful for manual pagination.

.INPUTS
None. You cannot pipe objects to Get-MessageTrace.

.OUTPUTS
System.Object[]
Returns an array of message tracking objects. Each object includes:
- id: Unique message identifier
- messageId: Internet message ID
- subject: Message subject
- from: Sender information
  * header: From header address
  * envelope: SMTP envelope sender
- to: Array of recipient information
  * address: Email address
  * status: Delivery status
  * details: Processing details
- sent: Date/time message was sent
- received: Date/time message was received
- route: Message route (inbound/outbound/internal)
- status: Current message status
- size: Message size in bytes
- processingDetails: Array of processing events
  * timestamp: When event occurred
  * type: Type of event
  * result: Result of processing
  * details: Additional event details
- policyMatches: Array of policy matches
  * policyType: Type of policy
  * policyName: Name of policy
  * action: Action taken
  * details: Match details

.EXAMPLE
C:\> Get-MessageTrace -FromEmailAddress 'user@example.com'
Gets message tracking for a specific sender.

.EXAMPLE
C:\> Get-MessageTrace `
    -StartDate (Get-Date).AddDays(-7) `
    -EndDate (Get-Date) `
    -Route 'outbound' `
    -Status 'delivered'

Gets delivered outbound messages from the last 7 days.

.EXAMPLE
C:\> Get-MessageTrace `
    -ToEmailAddress '*@example.com' `
    -Subject '*invoice*' `
    -First 100 `
    -OldestFirst

Gets the first 100 messages sent to example.com with 'invoice'
in the subject, ordered oldest first.

.EXAMPLE
C:\> Get-MessageTrace -MessageId '<message-id@domain.com>'
Gets tracking information for a specific message ID.

.NOTES
- Data retention is typically 30 days
- Date ranges affect search performance
- Some fields may be empty if not applicable
- API permissions required:
  * "Message > Track and Trace > Read"
#>
function Get-MessageTrace {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string[]]$FromEmailAddress,

        [Parameter()]
        [string[]]$ToEmailAddress,

        [Parameter()]
        [string]$Subject,

        [Parameter()]
        [ValidateSet('inbound', 'outbound', 'internal', 'all')]
        [string]$Route = 'all',

        [Parameter()]
        [ValidateSet('accepted', 'rejected', 'delivered', 'held', 'failed', 'processing', 'all')]
        [string]$Status = 'all',

        [Parameter()]
        [string]$MessageId,

        [Parameter()]
        [datetime]$StartDate = (Get-Date).AddDays(-7),

        [Parameter()]
        [datetime]$EndDate = (Get-Date),

        [Parameter()]
        [switch]$OldestFirst,

        [Parameter()]
        [switch]$All = $true,

        [Parameter()]
        [ValidateRange(1, [int]::MaxValue)]
        [int]$First,

        [Parameter()]
        [ValidateRange(0, [int]::MaxValue)]
        [int]$Skip = 0,

        [Parameter()]
        [ValidateRange(1, 500)]
        [int]$PageSize = 100
    )

    begin {
        if ($First -and $All) {
            Write-Warning "-All parameter is ignored when -First is specified."
            $All = $false
        }

        # Set up pagination variables
        $startRow = $Skip
        $maxResults = if ($First) { $First } else { [int]::MaxValue }
        $results = [System.Collections.ArrayList]::new()
    }

    process {
        try {
            # Initialize pagination
            $pageToken = $null
            $totalRetrieved = 0
            $hasMore = $true

            while ($hasMore -and $totalRetrieved -lt $maxResults) {
                # Prepare request body
                $body = @{
                    data = @(
                        @{
                            start = $StartDate.ToString('yyyy-MM-ddTHH:mm:ssZ')
                            end = $EndDate.ToString('yyyy-MM-ddTHH:mm:ssZ')
                            oldestFirst = $OldestFirst.IsPresent
                            pageSize = [Math]::Min($PageSize, $maxResults - $totalRetrieved)
                            advancedTrackAndTraceOptions = @{}
                        }
                    )
                }

                # Add search criteria
                if ($FromEmailAddress) {
                    $body.data[0].advancedTrackAndTraceOptions.from = $FromEmailAddress -join ','
                }
                if ($ToEmailAddress) {
                    $body.data[0].advancedTrackAndTraceOptions.to = $ToEmailAddress -join ','
                }
                if ($Subject) {
                    $body.data[0].advancedTrackAndTraceOptions.subject = $Subject
                }
                if ($Route -ne 'all') {
                    $body.data[0].advancedTrackAndTraceOptions.route = $Route
                }
                if ($Status -ne 'all') {
                    $body.data[0].advancedTrackAndTraceOptions.status = $Status
                }
                if ($MessageId) {
                    $body.data[0].advancedTrackAndTraceOptions.messageId = $MessageId
                }

                if ($pageToken) {
                    $body.data[0].pageToken = $pageToken
                }

                # Make API request
                $response = Invoke-ApiRequest -Method 'POST' -Path '/api/message-finder/search' -Body $body

                # Process results
                foreach ($message in $response.data) {
                    if ($totalRetrieved -ge $maxResults) {
                        $hasMore = $false
                        break
                    }

                    # Transform message object
                    $messageObj = [PSCustomObject]@{
                        Id = $message.id
                        MessageId = $message.recipientInfo.messageInfo.messageId
                        Subject = $message.recipientInfo.messageInfo.subject
                        From = [PSCustomObject]@{
                            Header = $message.recipientInfo.messageInfo.fromHeader
                            Envelope = $message.recipientInfo.messageInfo.fromEnvelope
                        }
                        To = @($message.recipientInfo.messageInfo.to | ForEach-Object {
                            [PSCustomObject]@{
                                Address = $_
                                Status = $message.deliveredMessage.$_.messageInfo.status
                                Details = $message.deliveredMessage.$_.txInfo.queueDetailStatus
                            }
                        })
                        Sent = [datetime]::Parse($message.recipientInfo.messageInfo.sent)
                        Received = [datetime]::Parse($message.recipientInfo.messageInfo.processed)
                        Route = $message.recipientInfo.messageInfo.route
                        Status = $message.status
                        Size = $message.recipientInfo.txInfo.transmissionSize
                        ProcessingDetails = @(
                            if ($message.recipientInfo.txInfo) {
                                [PSCustomObject]@{
                                    Timestamp = [datetime]::Parse($message.recipientInfo.txInfo.processStartTime)
                                    Type = 'Receipt'
                                    Result = $message.recipientInfo.txInfo.txEvent
                                    Details = $message.recipientInfo.txInfo.queueDetailStatus
                                }
                            }
                            if ($message.deliveredMessage) {
                                foreach ($recipient in $message.deliveredMessage.PSObject.Properties) {
                                    [PSCustomObject]@{
                                        Timestamp = [datetime]::Parse($recipient.Value.txInfo.processStartTime)
                                        Type = 'Delivery'
                                        Result = $recipient.Value.txInfo.txEvent
                                        Details = $recipient.Value.txInfo.queueDetailStatus
                                    }
                                }
                            }
                        )
                        PolicyMatches = @(
                            if ($message.deliveredMessage) {
                                foreach ($recipient in $message.deliveredMessage.PSObject.Properties) {
                                    foreach ($policy in $recipient.Value.policyInfo) {
                                        [PSCustomObject]@{
                                            PolicyType = $policy.policyType
                                            PolicyName = $policy.policyName
                                            Action = $policy.action
                                            Details = $policy.details
                                        }
                                    }
                                }
                            }
                        )
                    }

                    [void]$results.Add($messageObj)
                    $totalRetrieved++
                }

                # Check for more pages
                if ($response.meta.pagination.next) {
                    $pageToken = $response.meta.pagination.next
                }
                else {
                    $hasMore = $false
                }

                # Stop if we're not retrieving all results
                if (-not $All) {
                    $hasMore = $false
                }
            }
        }
        catch {
            Write-Error -ErrorRecord $_
        }
    }

    end {
        return $results.ToArray()
    }
}

function Get-MimecastProcessingMessage {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]$Id,

        [Parameter()]
        [string]$FromAddress,

        [Parameter()]
        [string]$ToAddress,

        [Parameter()]
        [string]$Subject,

        [Parameter()]
        [ValidateSet('inbound', 'outbound')]
        [string]$Route,

        [Parameter()]
        [switch]$FirstPageOnly
    )

    try {
        $body = @{
            data = @(
                @{}
            )
        }

        if ($Id) {
            $body.data[0].id = $Id
        }
        if ($FromAddress) {
            $body.data[0].fromAddress = $FromAddress
        }
        if ($ToAddress) {
            $body.data[0].toAddress = $ToAddress
        }
        if ($Subject) {
            $body.data[0].subject = $Subject
        }
        if ($Route) {
            $body.data[0].route = $Route
        }

        # Use pagination helper
        Invoke-ApiRequestWithPagination -Path '/api/gateway/get-processing-message-list' `
            -InitialBody $body `
            -ReturnFirstPageOnly:$FirstPageOnly
    }
    catch {
        Write-Error -ErrorRecord $_
    }
}