function Write-ApiLog {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$Message,
        
        [Parameter()]
        [ValidateSet('Information', 'Warning', 'Error')]
        [string]$Level = 'Information',
        
        [Parameter()]
        [switch]$PassThru
    )
    
    $LogFilePath = $ExecutionContext.SessionState.Module.PrivateData.LogFilePath
    
    $TimeStamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $LogMessage = "[$TimeStamp] [$Level] $Message"
    
    Write-Verbose $LogMessage
    
    if ($LogFilePath) {
        try {
            $LogMessage | Add-Content -Path $LogFilePath -ErrorAction Stop
        }
        catch {
            Write-Warning "Failed to write to log file: $_"
        }
    }
    
    if ($PassThru) {
        $LogMessage
    }
}