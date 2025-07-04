function Get-MimecastThreatRemediation {
    [CmdletBinding()]
    param (
        [Parameter()]
        [ValidateSet('all', 'pending', 'completed', 'failed')]
        [string]$Status = 'all',

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
        $endpoint = "api/ttp/remediation/get-remediations"
        $results = @()
    }

    process {
        $body = @{
            data = @(
                @{
                    status = $Status
                }
            )
        }

        if ($StartDate) {
            $body.data[0].from = $StartDate.ToString('o')
        }
        if ($EndDate) {
            $body.data[0].to = $EndDate.ToString('o')
        }

        # Handle pagination
        do {
            $response = Invoke-ApiRequest -Method Post -Path $endpoint -Body $body
            
            if (-not $response.success) {
                Write-Error "Failed to get threat remediations: $($response.fail)"
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

function Start-MimecastThreatRemediation {
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]$Code,

        [Parameter(Mandatory)]
        [ValidateSet('delete_message', 'restore_message')]
        [string]$Action,

        [Parameter()]
        [string]$Reason = "Automated remediation via PowerShell module"
    )

    begin {
        $endpoint = "api/ttp/remediation/create-remediation"
    }

    process {
        if (-not $PSCmdlet.ShouldProcess($Code, "Start threat remediation with action: $Action")) {
            return
        }

        $body = @{
            data = @(
                @{
                    code = $Code
                    action = $Action
                    reason = $Reason
                }
            )
        }

        $response = Invoke-ApiRequest -Method Post -Path $endpoint -Body $body

        if (-not $response.success) {
            Write-Error "Failed to start threat remediation: $($response.fail)"
            return
        }

        $response.data
    }
}

function Get-MimecastThreatRemediationStatus {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]$Code
    )

    begin {
        $endpoint = "api/ttp/remediation/get-remediation"
    }

    process {
        $body = @{
            data = @(
                @{
                    code = $Code
                }
            )
        }

        $response = Invoke-ApiRequest -Method Post -Path $endpoint -Body $body

        if (-not $response.success) {
            Write-Error "Failed to get threat remediation status: $($response.fail)"
            return
        }

        $response.data
    }
}

function Stop-MimecastThreatRemediation {
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]$Code,

        [Parameter()]
        [string]$Reason = "Cancelled via PowerShell module"
    )

    begin {
        $endpoint = "api/ttp/remediation/delete-remediation"
    }

    process {
        if (-not $PSCmdlet.ShouldProcess($Code, "Stop threat remediation")) {
            return
        }

        $body = @{
            data = @(
                @{
                    code = $Code
                    reason = $Reason
                }
            )
        }

        $response = Invoke-ApiRequest -Method Post -Path $endpoint -Body $body

        if (-not $response.success) {
            Write-Error "Failed to stop threat remediation: $($response.fail)"
            return
        }

        $response.data
    }
}