[CmdletBinding(DefaultParameterSetName='Default')]
param (
    [Parameter()]
    [string]$ClientIdName = 'MimecastClientId',

    [Parameter()]
    [string]$ClientSecretName = 'MimecastClientSecret',

    [Parameter(Mandatory)]
    [string]$SIEMEndpoint,

    [Parameter()]
    [ValidateSet('receipt', 'entities')]
    [string]$Type = 'receipt',

    [Parameter()]
    [ValidateRange(1, 1000)]
    [int]$PageSize = 100,

    [Parameter()]
    [datetime]$StartDate,

    [Parameter()]
    [datetime]$EndDate = (Get-Date),

    [Parameter()]
    [string]$LogOutputDirectory,

    [Parameter()]
    [string]$StateDirectory
)

# Initialize transcript logging
$transcriptPath = Join-Path ($StateDirectory ?? $LogOutputDirectory ?? $PWD) "MimecastLogCollection_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
Start-Transcript -Path $transcriptPath -Force

try {
    Write-Verbose "Starting Mimecast log collection at $(Get-Date)"
    
    # Import the Mimecast module
    $modulePath = Join-Path $PSScriptRoot "..\Mimecast.psd1"
    Import-Module $modulePath -ErrorAction Stop
    Write-Verbose "Successfully imported Mimecast module"

    # Connect to Mimecast API
    Write-Verbose "Connecting to Mimecast API..."
    Connect-Mimecast -ClientIdName $ClientIdName -ClientSecretName $ClientSecretName -ErrorAction Stop
    Write-Verbose "Successfully connected to Mimecast API"

    # Create log output directory if specified
    if ($LogOutputDirectory) {
        if (-not (Test-Path -Path $LogOutputDirectory)) {
            New-Item -Path $LogOutputDirectory -ItemType Directory -Force | Out-Null
        }
        Write-Verbose "Log output directory: $LogOutputDirectory"
    }

    # Initialize counters
    $totalEventsCollected = 0
    $batchCount = 0

    # Build base URL for SIEM events
    $basePath = "/v1/siem/events"
    if ($Type -eq 'receipt') {
        $basePath += "/cg" # Customer Gateway events
    }
    else {
        $basePath += "/ci" # Customer Intelligence events
    }

    # Build query parameters
    $queryParams = @{
        pageSize = $PageSize
    }
    if ($StartDate) {
        $queryParams.dateRangeStartsAt = $StartDate.ToString('yyyy-MM-ddTHH:mm:ssZ')
    }
    if ($EndDate) {
        $queryParams.dateRangeEndsAt = $EndDate.ToString('yyyy-MM-ddTHH:mm:ssZ')
    }

    # Initialize pagination token
    $nextToken = $null

    # Collect logs in batches
    do {
        $batchCount++
        Write-Verbose "Processing batch #$batchCount..."

        try {
            # Add pagination token if we have one
            if ($nextToken) {
                $queryParams.next = $nextToken
            }

            # Call the SIEM events endpoint
            $response = Invoke-ApiRequest -Method 'GET' -Path $basePath -QueryParameters $queryParams
            
            # Check if we got any events
            if (-not $response.data -or $response.data.Count -eq 0) {
                Write-Verbose "No new events available"
                break
            }

            $eventsInBatch = $response.data.Count
            $totalEventsCollected += $eventsInBatch
            Write-Verbose "Retrieved $eventsInBatch events in this batch (Total: $totalEventsCollected)"

            # Save events locally if directory specified
            if ($LogOutputDirectory) {
                $batchFile = Join-Path $LogOutputDirectory "MimecastLogs_$(Get-Date -Format 'yyyyMMdd_HHmmss')_Batch$batchCount.json"
                $response.data | ConvertTo-Json -Depth 10 | Out-File -FilePath $batchFile -Force
                Write-Verbose "Saved batch to: $batchFile"
            }

            # Forward events to SIEM endpoint
            foreach ($event in $response.data) {
                # TODO: Implement SIEM forwarding logic
                # For now, just output to console in a syslog-like format
                $eventMessage = "$($event.timestamp) mimecast[$($event.id)]: $($event | ConvertTo-Json -Compress)"
                Write-Output $eventMessage
            }

            # Get the next token from pagination info
            $pagination = $response.meta.pagination
            if ($pagination -and $pagination.next) {
                $nextToken = $pagination.next
                Write-Verbose "Retrieved next page token"
            }
            else {
                Write-Verbose "No next page token, ending collection"
                break
            }
        }
        catch {
            Write-Error "Error processing batch #$batchCount: $_"
            throw
        }

        # Small delay between batches to avoid rate limiting
        Start-Sleep -Seconds 1

    } while ($true)

    Write-Verbose "Log collection completed. Total events collected: $totalEventsCollected"
}
catch {
    Write-Error "Error during log collection: $_"
    exit 1
}
finally {
    # Always disconnect from the API
    try {
        Disconnect-Mimecast
        Write-Verbose "Disconnected from Mimecast API"
    }
    catch {
        Write-Warning "Error during API disconnect: $_"
    }

    Stop-Transcript
}