function Get-MimecastHeldMessage {
    [CmdletBinding()]
    param (
        [Parameter()]
        [ValidateSet('inbound', 'outbound', 'internal')]
        [string]$Route,

        [Parameter()]
        [string]$Sender,

        [Parameter()]
        [string]$Recipient,

        [Parameter()]
        [string]$Subject,

        [Parameter()]
        [ValidateSet(
            'spam', 'suspected_spam', 'signature', 'av', 'content', 'policy', 
            'tlsv1', 'dkim', 'spf', 'dmarc', 'impersonation', 'malware_sandbox'
        )]
        [string]$HoldType,

        [Parameter()]
        [DateTime]$StartDate,

        [Parameter()]
        [DateTime]$EndDate,

        [Parameter()]
        [int]$PageSize = 100,

        [Parameter()]
        [switch]$All
    )

    begin {
        $endpoint = "api/gateway/hold-messages/get-hold-messages"
        $results = @()
    }

    process {
        $body = @{
            data = @(
                @{}
            )
        }

        # Add optional parameters
        if ($Route) { $body.data[0].route = $Route }
        if ($Sender) { $body.data[0].from = $Sender }
        if ($Recipient) { $body.data[0].to = $Recipient }
        if ($Subject) { $body.data[0].subject = $Subject }
        if ($HoldType) { $body.data[0].holdType = $HoldType }
        if ($StartDate) { $body.data[0].from = $StartDate.ToString('o') }
        if ($EndDate) { $body.data[0].to = $EndDate.ToString('o') }
        if ($PageSize) { $body.data[0].pageSize = $PageSize }

        # Handle pagination
        do {
            $response = Invoke-ApiRequest -Method Post -Path $endpoint -Body $body
            
            if (-not $response.success) {
                Write-Error "Failed to get held messages: $($response.fail)"
                return
            }

            $results += $response.data

            # Update pagination
            if ($response.pagination.next) {
                $body.data[0].next = $response.pagination.next
            }

        } while ($All -and $response.pagination.next)

        $results
    }
}

function Release-MimecastHeldMessage {
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string[]]$Id,

        [Parameter()]
        [ValidateSet('release_clean', 'release_with_tag')]
        [string]$Action = 'release_clean',

        [Parameter()]
        [string]$Reason = "Released via PowerShell module"
    )

    begin {
        $endpoint = "api/gateway/hold-messages/release-hold-message"
    }

    process {
        foreach ($messageId in $Id) {
            if (-not $PSCmdlet.ShouldProcess($messageId, "Release held message with action: $Action")) {
                continue
            }

            $body = @{
                data = @(
                    @{
                        id = $messageId
                        action = $Action
                        reason = $Reason
                    }
                )
            }

            $response = Invoke-ApiRequest -Method Post -Path $endpoint -Body $body

            if (-not $response.success) {
                Write-Error "Failed to release message $messageId`: $($response.fail)"
                continue
            }

            $response.data
        }
    }
}

function Block-MimecastHeldMessage {
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string[]]$Id,

        [Parameter()]
        [string]$Reason = "Blocked via PowerShell module"
    )

    begin {
        $endpoint = "api/gateway/hold-messages/block-hold-message"
    }

    process {
        foreach ($messageId in $Id) {
            if (-not $PSCmdlet.ShouldProcess($messageId, "Block held message")) {
                continue
            }

            $body = @{
                data = @(
                    @{
                        id = $messageId
                        reason = $Reason
                    }
                )
            }

            $response = Invoke-ApiRequest -Method Post -Path $endpoint -Body $body

            if (-not $response.success) {
                Write-Error "Failed to block message $messageId`: $($response.fail)"
                continue
            }

            $response.data
        }
    }
}

function Get-MimecastHeldMessageContent {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]$Id,

        [Parameter()]
        [ValidateSet('html', 'text')]
        [string]$Type = 'html'
    )

    begin {
        $endpoint = "api/gateway/hold-messages/get-hold-message-content"
    }

    process {
        $body = @{
            data = @(
                @{
                    id = $Id
                    type = $Type
                }
            )
        }

        $response = Invoke-ApiRequest -Method Post -Path $endpoint -Body $body

        if (-not $response.success) {
            Write-Error "Failed to get message content for $Id`: $($response.fail)"
            return
        }

        $response.data
    }
}

function Save-MimecastHeldMessageAttachment {
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]$Id,

        [Parameter(Mandatory)]
        [string]$AttachmentId,

        [Parameter(Mandatory)]
        [string]$OutputPath,

        [Parameter()]
        [switch]$Force
    )

    begin {
        $endpoint = "api/gateway/hold-messages/get-hold-message-attachment"
    }

    process {
        # Check if file exists
        if (Test-Path $OutputPath) {
            if (-not $Force) {
                Write-Error "File already exists at $OutputPath. Use -Force to overwrite."
                return
            }
        }

        if (-not $PSCmdlet.ShouldProcess($Id, "Download attachment to $OutputPath")) {
            return
        }

        $body = @{
            data = @(
                @{
                    id = $Id
                    attachmentId = $AttachmentId
                }
            )
        }

        try {
            $response = Invoke-ApiRequest -Method Post -Path $endpoint -Body $body

            if (-not $response.success) {
                Write-Error "Failed to download attachment: $($response.fail)"
                return
            }

            # Save attachment content
            [System.Convert]::FromBase64String($response.data.attachmentContent) | 
                Set-Content -Path $OutputPath -Encoding Byte -Force:$Force

            # Return file info
            Get-Item $OutputPath
        }
        catch {
            Write-Error "Failed to save attachment: $_"
        }
    }
}

function Get-MimecastHeldMessageHeader {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]$Id
    )

    begin {
        $endpoint = "api/gateway/hold-messages/get-hold-message-headers"
    }

    process {
        $body = @{
            data = @(
                @{
                    id = $Id
                }
            )
        }

        $response = Invoke-ApiRequest -Method Post -Path $endpoint -Body $body

        if (-not $response.success) {
            Write-Error "Failed to get message headers for $Id`: $($response.fail)"
            return
        }

        $response.data
    }
}