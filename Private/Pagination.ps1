function Invoke-ApiRequestWithPagination {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter()]
        [object]$InitialBody = @{ data = @( @{} ) },

        [Parameter()]
        [switch]$ReturnFirstPageOnly
    )

    try {
        $allResults = @()
        $body = $InitialBody.Clone()
        $hasMorePages = $true

        Write-Verbose "Starting paginated request to $Path"

        while ($hasMorePages) {
            $response = Invoke-ApiRequest -Method 'POST' -Path $Path -Body $body

            if ($response.data) {
                $allResults += $response.data
            }

            # Check if we should stop after first page
            if ($ReturnFirstPageOnly) {
                Write-Verbose "Returning first page only as requested"
                return $response.data
            }

            # Check for pagination info
            $pagination = $response.meta.pagination
            if (-not $pagination -or -not $pagination.next) {
                $hasMorePages = $false
                Write-Verbose "No more pages available"
            }
            else {
                Write-Verbose "Found next page token: $($pagination.next)"
                # Update body with next page token
                if (-not $body.data[0].ContainsKey('pagination')) {
                    $body.data[0].Add('pagination', @{})
                }
                $body.data[0].pagination.pageToken = $pagination.next
            }
        }

        Write-Verbose "Retrieved total of $($allResults.Count) items"
        return $allResults
    }
    catch {
        Write-Error -ErrorRecord $_
    }
}