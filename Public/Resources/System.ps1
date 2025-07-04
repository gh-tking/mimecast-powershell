function Get-MimecastApiSystemInfo {
    [CmdletBinding()]
    param()
    
    try {
        # Call the Mimecast API status endpoint
        $response = Invoke-ApiRequest -Method 'GET' -Path '/api/system/info'
        
        # Return the response
        $response
    }
    catch {
        Write-Error -ErrorRecord $_
    }
}