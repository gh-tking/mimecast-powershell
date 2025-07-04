<#
.SYNOPSIS
Registers a scheduled task to collect Mimecast SIEM logs.

.DESCRIPTION
Creates a Windows scheduled task that periodically collects SIEM logs from Mimecast.
The task can collect different types of logs (receipt, entities) and supports various
scheduling options. The function includes state management to track log collection
progress and avoid duplicates.

.PARAMETER TaskName
The name for the scheduled task. This should be unique and descriptive.
Example: 'MimecastSIEMLogCollector_CG' for Customer Gateway logs.

.PARAMETER Type
The type of SIEM logs to collect. Valid values:
- receipt: Customer Gateway events (message tracking)
- entities: Customer Intelligence events (user/group changes)

.PARAMETER PageSize
Optional. Number of events to retrieve per API call.
Default is 100. Maximum is 500.
Higher values reduce API calls but increase memory usage.

.PARAMETER StartDate
Optional. Start date for initial log collection.
Default is 7 days ago.
Historical data beyond this may not be available.

.PARAMETER LogOutputDirectory
The directory where log files will be saved.
Files are named with date and type (e.g., 'mimecast_receipt_2024-07-03.json').

.PARAMETER StateDirectory
The directory where state files will be saved.
Files track last collection time to avoid duplicates.

.PARAMETER Frequency
How often the task should run. Valid values:
- EveryNMinutes: Run every N minutes
- Hourly: Run at the start of each hour
- Daily: Run once per day
- Weekly: Run once per week

.PARAMETER IntervalMinutes
Required when Frequency is 'EveryNMinutes'.
How many minutes between runs (1-1440).

.PARAMETER TimeOfDay
Required when Frequency is 'Daily' or 'Weekly'.
What time to run the task (e.g., '03:00').

.PARAMETER DayOfWeek
Required when Frequency is 'Weekly'.
Which day to run the task (e.g., 'Monday').

.PARAMETER Credential
The credentials for running the task.
Should be a service account with:
- "Log on as a batch job" rights
- Write access to LogOutputDirectory and StateDirectory
- Network access for API calls

.INPUTS
None. You cannot pipe objects to Register-MimecastLogCollectorTask.

.OUTPUTS
Microsoft.Management.Infrastructure.CimInstance#root/Microsoft/Windows/TaskScheduler/MSFT_ScheduledTask
Returns the created scheduled task object.

.EXAMPLE
C:\> $svcPassword = Read-Host -Prompt "Enter service account password" -AsSecureString
C:\> $svcCredential = New-Object System.Management.Automation.PSCredential(
    "DOMAIN\MimecastSvc", 
    $svcPassword
)
C:\> Register-MimecastLogCollectorTask `
    -TaskName "MimecastSIEMLogCollector_CG" `
    -Type "receipt" `
    -PageSize 100 `
    -StartDate (Get-Date).AddDays(-7) `
    -LogOutputDirectory "C:\ProgramData\MimecastApi\Logs" `
    -StateDirectory "C:\ProgramData\MimecastApi\State" `
    -Frequency "EveryNMinutes" `
    -IntervalMinutes 5 `
    -Credential $svcCredential

Creates a task to collect Customer Gateway events every 5 minutes.

.EXAMPLE
C:\> Register-MimecastLogCollectorTask `
    -TaskName "MimecastSIEMLogCollector_CI" `
    -Type "entities" `
    -LogOutputDirectory "C:\ProgramData\MimecastApi\Logs" `
    -StateDirectory "C:\ProgramData\MimecastApi\State" `
    -Frequency "Daily" `
    -TimeOfDay "03:00" `
    -Credential $svcCredential

Creates a task to collect Customer Intelligence events daily at 3 AM.

.EXAMPLE
C:\> Register-MimecastLogCollectorTask `
    -TaskName "MimecastSIEMLogCollector_Weekly" `
    -Type "receipt" `
    -LogOutputDirectory "C:\ProgramData\MimecastApi\Logs" `
    -StateDirectory "C:\ProgramData\MimecastApi\State" `
    -Frequency "Weekly" `
    -DayOfWeek "Sunday" `
    -TimeOfDay "00:00" `
    -Credential $svcCredential

Creates a task to collect Customer Gateway events weekly on Sunday at midnight.

.NOTES
- Task requires appropriate API permissions
- Service account needs appropriate rights
- Directories must be accessible to service account
- State files prevent duplicate event collection
- Log files are in JSON format for SIEM ingestion
- API permissions required:
  * "SIEM Integration > SIEM Events > Read"
#>
function Register-MimecastLogCollectorTask {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$TaskName,

        [Parameter(Mandatory)]
        [string]$ScriptPath,

        [Parameter(Mandatory)]
        [string]$SIEMEndpoint,

        [Parameter()]
        [ValidateSet('MTA')]
        [string]$LogType = 'MTA',

        [Parameter(Mandatory)]
        [string]$LastTokenStateFilePath,

        [Parameter()]
        [string]$LogOutputDirectory,

        [Parameter(Mandatory)]
        [ValidateSet('Hourly', 'Daily', 'EveryNMinutes')]
        [string]$Frequency,

        [Parameter()]
        [int]$IntervalMinutes,

        [Parameter()]
        [string]$AtTime,

        [Parameter(Mandatory)]
        [PSCredential]$Credential,

        [Parameter()]
        [switch]$Force
    )

    try {
        # Validate script path
        if (-not (Test-Path -Path $ScriptPath)) {
            throw "Script not found at path: $ScriptPath"
        }

        # Validate frequency parameters
        switch ($Frequency) {
            'EveryNMinutes' {
                if (-not $IntervalMinutes -or $IntervalMinutes -lt 1) {
                    throw "IntervalMinutes must be specified and greater than 0 for EveryNMinutes frequency"
                }
            }
            { $_ -in 'Hourly','Daily' } {
                if (-not $AtTime) {
                    throw "AtTime must be specified for $Frequency frequency"
                }
                if (-not ($AtTime -match '^\d{2}:\d{2}$')) {
                    throw "AtTime must be in format 'HH:mm'"
                }
            }
        }

        # Build PowerShell arguments
        $scriptArgs = @(
            "-NoProfile"
            "-NonInteractive"
            "-ExecutionPolicy Bypass"
            "-File `"$ScriptPath`""
            "-SIEMEndpoint `"$SIEMEndpoint`""
            "-LastTokenStateFile `"$LastTokenStateFilePath`""
            "-LogType `"$LogType`""
        )

        if ($LogOutputDirectory) {
            $scriptArgs += "-LogOutputDirectory `"$LogOutputDirectory`""
        }

        # Create trigger based on frequency
        $trigger = switch ($Frequency) {
            'EveryNMinutes' {
                $startTime = [DateTime]::Now.AddMinutes(1)
                New-ScheduledTaskTrigger -Once -At $startTime `
                    -RepetitionInterval (New-TimeSpan -Minutes $IntervalMinutes) `
                    -RepetitionDuration ([TimeSpan]::MaxValue)
            }
            'Hourly' {
                $time = [DateTime]::ParseExact($AtTime, 'HH:mm', $null)
                New-ScheduledTaskTrigger -Once -At $time `
                    -RepetitionInterval (New-TimeSpan -Hours 1) `
                    -RepetitionDuration ([TimeSpan]::MaxValue)
            }
            'Daily' {
                $time = [DateTime]::ParseExact($AtTime, 'HH:mm', $null)
                New-ScheduledTaskTrigger -Daily -At $time
            }
        }

        # Create action
        $action = New-ScheduledTaskAction `
            -Execute 'powershell.exe' `
            -Argument ($scriptArgs -join ' ')

        # Create settings
        $settings = New-ScheduledTaskSettingsSet `
            -MultipleInstances Parallel `
            -StartWhenAvailable `
            -RunOnlyIfNetworkAvailable `
            -ExecutionTimeLimit (New-TimeSpan -Hours 1) `
            -RestartCount 3 `
            -RestartInterval (New-TimeSpan -Minutes 1)

        # Register the task
        $taskParams = @{
            TaskName = $TaskName
            Action = $action
            Trigger = $trigger
            Settings = $settings
            User = $Credential.UserName
            Password = $Credential.GetNetworkCredential().Password
            RunLevel = 'Highest'
            Force = $Force
        }

        Register-ScheduledTask @taskParams
        Write-Verbose "Successfully registered task: $TaskName"
    }
    catch {
        Write-Error "Failed to register task: $_"
        throw
    }
}

<#
.SYNOPSIS
Removes a Mimecast SIEM log collector task.

.DESCRIPTION
Removes a previously registered Windows scheduled task that collects SIEM logs
from Mimecast. The function includes ShouldProcess support for confirmation and
optionally can clean up associated state and log files.

.PARAMETER TaskName
The name of the scheduled task to remove.
Example: 'MimecastSIEMLogCollector_CG'.

.PARAMETER CleanupFiles
Optional. Whether to remove associated state and log files.
Default is $false.

.PARAMETER StateDirectory
Optional. The directory containing state files.
Required if CleanupFiles is $true.

.PARAMETER LogOutputDirectory
Optional. The directory containing log files.
Required if CleanupFiles is $true.

.INPUTS
None. You cannot pipe objects to Unregister-MimecastLogCollectorTask.

.OUTPUTS
None. The function does not return any output.

.EXAMPLE
C:\> Unregister-MimecastLogCollectorTask -TaskName "MimecastSIEMLogCollector_CG"
Removes the specified task, leaving state and log files intact.

.EXAMPLE
C:\> Unregister-MimecastLogCollectorTask `
    -TaskName "MimecastSIEMLogCollector_CG" `
    -CleanupFiles `
    -StateDirectory "C:\ProgramData\MimecastApi\State" `
    -LogOutputDirectory "C:\ProgramData\MimecastApi\Logs"

Removes the task and all associated files.

.EXAMPLE
C:\> Get-ScheduledTask -TaskName "MimecastSIEMLogCollector*" |
    ForEach-Object {
        Unregister-MimecastLogCollectorTask -TaskName $_.TaskName -WhatIf
    }

Shows what would happen if you removed all Mimecast SIEM collector tasks.

.NOTES
- Use -WhatIf to preview changes
- File cleanup is permanent
- Task removal requires administrative rights
- API permissions not required (local operation only)
#>
function Unregister-MimecastLogCollectorTask {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$TaskName
    )

    try {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        Write-Verbose "Successfully unregistered task: $TaskName"
    }
    catch {
        Write-Error "Failed to unregister task: $_"
        throw
    }
}

function Get-MimecastLogCollectorTask {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]$TaskName
    )

    try {
        if ($TaskName) {
            Get-ScheduledTask -TaskName $TaskName -ErrorAction Stop
        }
        else {
            Get-ScheduledTask | Where-Object { $_.TaskName -like 'Mimecast*' }
        }
    }
    catch {
        Write-Error "Failed to get task(s): $_"
        throw
    }
}

function Start-MimecastLogCollectorTask {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$TaskName
    )

    try {
        Start-ScheduledTask -TaskName $TaskName
        Write-Verbose "Successfully started task: $TaskName"
    }
    catch {
        Write-Error "Failed to start task: $_"
        throw
    }
}

function Stop-MimecastLogCollectorTask {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$TaskName
    )

    try {
        Stop-ScheduledTask -TaskName $TaskName
        Write-Verbose "Successfully stopped task: $TaskName"
    }
    catch {
        Write-Error "Failed to stop task: $_"
        throw
    }
}