function Invoke-ApiRequest {
    [CmdletBinding(DefaultParameterSetName='Default')]
    param (
        [Parameter(Mandatory)]
        [ValidateSet('GET', 'POST', 'PUT', 'DELETE')]
        [string]$Method,

        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter()]
        [object]$Body,

        [Parameter()]
        [hashtable]$QueryParameters,

        [Parameter()]
        [hashtable]$Headers = @{},

        [Parameter()]
        [switch]$Raw
    )

    begin {
        $BaseUri = $ExecutionContext.SessionState.Module.PrivateData.BaseUri
        if (-not $BaseUri) {
            throw "BaseUri not configured. Use Set-MimecastApiConfiguration to set the base URI."
        }

        # Check if we have a valid access token
        $accessToken = $ExecutionContext.SessionState.Module.PrivateData.AccessToken
        $tokenExpiresAt = $ExecutionContext.SessionState.Module.PrivateData.TokenExpiresAt

        if (-not $accessToken -or -not $tokenExpiresAt -or $tokenExpiresAt -le [DateTime]::UtcNow) {
            throw "No valid access token found. Please connect using Connect-MimecastApi first."
        }

        # Set OAuth2 Bearer token
        $Headers['Authorization'] = "Bearer $accessToken"
        $Headers['Content-Type'] = 'application/json'
    }

    process {
        try {
            # Build the URI
            $Uri = [System.Uri]::new([System.Uri]::new($BaseUri), $Path)
            if ($QueryParameters) {
                $QueryString = [System.Web.HttpUtility]::ParseQueryString('')
                foreach ($param in $QueryParameters.GetEnumerator()) {
                    $QueryString[$param.Key] = $param.Value
                }
                $UriBuilder = [System.UriBuilder]::new($Uri)
                $UriBuilder.Query = $QueryString.ToString()
                $Uri = $UriBuilder.Uri
            }

            # Build parameters for Invoke-RestMethod
            $params = @{
                Method = $Method
                Uri = $Uri
                Headers = $Headers
                UseBasicParsing = $true
                ErrorAction = 'Stop'
            }

            if ($Body) {
                $params['Body'] = ($Body | ConvertTo-Json -Depth 10)
            }

            Write-Verbose "Sending $Method request to $Uri"
            Write-Verbose "Request ID: $RequestId"
            
            $response = Invoke-RestMethod @params

            if ($Raw) {
                $response
            }
            else {
                # Mimecast API typically wraps responses in a data object
                if ($response.data) {
                    $response.data
                }
                else {
                    $response
                }
            }
        }
        catch {
            Write-Error -ErrorRecord $_
        }
    }
}