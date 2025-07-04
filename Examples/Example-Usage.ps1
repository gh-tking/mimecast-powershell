# Import the module
Import-Module Mimecast

# Configure the module
Set-MimecastConfiguration -LogFilePath 'C:\Logs\Mimecast.log' -ApiVaultName 'MimecastVault'

# Method 1: Initialize first, then connect
# This will prompt for credentials if they haven't been stored yet
Initialize-MimecastModule

# Connect using stored credentials
Connect-Mimecast -Region 'US'  # or EU, DE, CA, ZA, AU, Offshore

# Method 2: Use the Setup switch (shortcut)
# This will initialize the module and connect in one step
Connect-Mimecast -Setup -Region 'US'

# Method 3: Use different credential names
# Useful for CI/CD or temporary overrides
$altClientId = Read-Host -Prompt "Enter alternative Client ID" -AsSecureString
$altClientSecret = Read-Host -Prompt "Enter alternative Client Secret" -AsSecureString

Set-ApiSecret -Name 'AlternativeClientId' -Secret $altClientId
Set-ApiSecret -Name 'AlternativeClientSecret' -Secret $altClientSecret

Connect-Mimecast `
    -ClientIdName 'AlternativeClientId' `
    -ClientSecretName 'AlternativeClientSecret' `
    -Region 'US'

# Get current configuration
Get-MimecastConfiguration

# Test the API connection
Get-MimecastSystemInfo

#
# User Management
#

# Get all users (automatically handles pagination)
Get-MimecastUser -Domain 'example.com'

# Get just the first page of users
Get-MimecastUser -Domain 'example.com' -FirstPageOnly

# Create a new user
$password = ConvertTo-SecureString 'MySecurePassword123!' -AsPlainText -Force
New-MimecastUser -EmailAddress 'newuser@example.com' -Name 'New User' -Password $password

# Update user
Get-MimecastUser -EmailAddress 'user@example.com' | Set-MimecastUser -AccountDisabled $true

# Get user aliases
Get-MimecastUser -UserEmailAddress 'user@example.com' -ParameterSetName 'Aliases'

#
# Group Management
#

# Get all groups (automatically handles pagination)
Get-MimecastGroup -Name 'Marketing'

# Get just the first page of groups
Get-MimecastGroup -Name 'Marketing' -FirstPageOnly

# Create a new group
New-MimecastGroup -Name 'Sales Team' -Description 'Sales department group' -Type 'distribution'

# Get all members of a group (automatically handles pagination)
$group = Get-MimecastGroup -Name 'Sales Team'
Get-MimecastGroup -GroupId $group.id -ParameterSetName 'Members'

# Add member to group
Add-MimecastGroupMember -GroupId $group.id -EmailAddress 'salesperson@example.com'

# Remove member from group
Remove-MimecastGroupMember -GroupId $group.id -EmailAddress 'formeremployee@example.com'

#
# Policy Management
#

# Create an address alteration definition
$altDef = New-MimecastAddressAlterationDefinition `
    -Description 'Redirect partner emails' `
    -OriginalAddress '*@old-partner.com' `
    -NewAddress '*@new-partner.com' `
    -AddressType 'both' `
    -Routing 'inbound' `
    -NotifySender $true `
    -NotifyRecipient $true `
    -CustomNotificationSubject 'Email address updated' `
    -CustomNotificationBody 'Your email was redirected to the new partner domain.'

# Create a policy using the address alteration definition
New-MimecastPolicy -PolicyType 'AddressAlteration' `
    -Description 'Partner domain redirection' `
    -FromType 'external' `
    -FromValue '*@old-partner.com' `
    -ToType 'internal' `
    -Direction 'inbound' `
    -DefinitionId $altDef.id

# Create a content examination definition with multiple rules
$rules = @(
    @{
        type = 'keyword'
        pattern = 'confidential'
        matchType = 'contains'
        caseSensitive = $false
        wholeWord = $true
    },
    @{
        type = 'pattern'
        pattern = '\b\d{16}\b'  # Credit card number pattern
        matchType = 'matches'
        caseSensitive = $false
    },
    @{
        type = 'size'
        size = 10485760  # 10MB
        operator = 'greater_than'
    },
    @{
        type = 'attachment'
        fileType = 'executable'
        includeArchives = $true
    }
)

$contentDef = New-MimecastContentExaminationDefinition `
    -Description 'DLP - Confidential Content' `
    -Action 'hold' `
    -ScanLocations @('subject', 'body', 'attachments') `
    -MatchCondition 'any' `
    -Rules $rules `
    -NotifyAdministrator $true `
    -CustomNotificationSubject 'DLP Alert: Confidential Content Detected' `
    -CustomNotificationBody 'A message containing potentially confidential content has been held for review.'

# Create a policy using the content examination definition
New-MimecastPolicy -PolicyType 'ContentExamination' `
    -Description 'DLP Policy - Confidential Content' `
    -FromType 'internal' `
    -ToType 'external' `
    -Direction 'outbound' `
    -DefinitionId $contentDef.id

# Get anti-spoofing bypass policies
Get-MimecastPolicy -PolicyType 'AntiSpoofingBypass'

# Create a new anti-spoofing bypass policy
New-MimecastPolicy -PolicyType 'AntiSpoofingBypass' `
    -Description 'Allow trusted partner domain' `
    -FromType 'external' `
    -FromValue '*@trusted-partner.com' `
    -ToType 'internal' `
    -Option 'enable_bypass'

# Update a policy
$policy = Get-MimecastPolicy -PolicyType 'AntiSpoofingBypass' | Select-Object -First 1
Set-MimecastPolicy -PolicyType 'AntiSpoofingBypass' -Id $policy.id -Enabled $false

# Remove a policy
Remove-MimecastPolicy -PolicyType 'AntiSpoofingBypass' -Id $policy.id

# Create a web white URL policy with targets
$targets = @(
    @{
        action = 'allow'
        type = 'domain'
        value = 'trusted-site.com'
    }
)

New-MimecastPolicy -PolicyType 'WebWhiteUrl' `
    -Description 'Allow trusted website' `
    -FromType 'internal' `
    -ToType 'everyone' `
    -Targets $targets

# Create an attachment management definition
$allowedContent = @(
    @{
        type = 'file_extension'
        value = 'pdf'
    },
    @{
        type = 'file_extension'
        value = 'docx'
    }
)

$blockedContent = @(
    @{
        type = 'file_extension'
        value = 'exe'
    },
    @{
        type = 'file_extension'
        value = 'bat'
    }
)

$attachDef = New-MimecastAttachmentManagementDefinition `
    -Description 'Standard Attachment Policy' `
    -DefaultBlockAllow 'block' `
    -PornographicImageSetting 'block' `
    -EncryptedArchives 'block' `
    -EncryptedDocuments 'block' `
    -UnreadableArchives 'block' `
    -LargeFileThreshold 10485760 `
    -EnableLargeFileSubstitution $true `
    -AllowedContent $allowedContent `
    -BlockedContent $blockedContent `
    -NotifyAdministrator $true `
    -CustomNotificationSubject 'Blocked Attachment Alert' `
    -CustomNotificationBody 'An email with a blocked attachment type was detected.'

# Create a policy using the attachment management definition
New-MimecastPolicy -PolicyType 'AttachmentManagement' `
    -Description 'Standard Attachment Policy' `
    -FromType 'external' `
    -ToType 'internal' `
    -Direction 'inbound' `
    -DefinitionId $attachDef.id

# Create a spam scanner definition
$spamDef = New-MimecastSpamScannerDefinition `
    -Description 'Aggressive Spam Policy' `
    -SpamDetectionLevel 'aggressive' `
    -ActionIfSpamDetected 'hold' `
    -SpamScoreThreshold 7 `
    -ActionIfSpamScoreHit 'hold' `
    -EnableSpamFeedback $true `
    -EnableGreylistingProtection $true `
    -EnableIntelligentRejection $true `
    -NotifyAdministrator $true `
    -CustomNotificationSubject 'Spam Alert' `
    -CustomNotificationBody 'A message was held due to high spam score.'

# Create a policy using the spam scanner definition
New-MimecastPolicy -PolicyType 'SpamScanning' `
    -Description 'Aggressive Spam Policy' `
    -FromType 'external' `
    -ToType 'internal' `
    -Direction 'inbound' `
    -DefinitionId $spamDef.id

# Create a notification set definition
$deliveryReceipts = @{
    subject = 'Message Delivered: {subject}'
    body = @'
Your message "{subject}" was delivered successfully.
From: {from}
To: {to}
Sent: {sent}
'@
}

$holdNotifications = @{
    subject = 'Message Held: {subject}'
    body = @'
Your message "{subject}" was held for review.
From: {from}
To: {to}
Sent: {sent}
Reason: {reason}
'@
}

$notificationDef = New-MimecastNotificationSetDefinition `
    -Description 'Standard Notification Set' `
    -PreferredLanguage 'en' `
    -UserMessage 'This is an automated notification from the email security system.' `
    -DeliveryReceipts $deliveryReceipts `
    -HoldNotifications $holdNotifications

# Create a policy using the notification set definition
New-MimecastPolicy -PolicyType 'NotificationSet' `
    -Description 'Standard Notification Set' `
    -FromType 'everyone' `
    -ToType 'everyone' `
    -DefinitionId $notificationDef.id

#
# Message Management
#

# Get all held messages (automatically handles pagination)
Get-MimecastHeldMessage -Route 'inbound'

# Get just the first page of held messages
Get-MimecastHeldMessage -Route 'inbound' -FirstPageOnly

# Search for messages (automatically handles pagination)
$startDate = (Get-Date).AddDays(-7)
$endDate = Get-Date
Search-MimecastMessage -FromAddress '*@example.com' -StartDate $startDate -EndDate $endDate

# Search messages but get only first page
Search-MimecastMessage -FromAddress '*@example.com' -StartDate $startDate -EndDate $endDate -FirstPageOnly

# Get message details
$message = Get-MimecastHeldMessage | Select-Object -First 1
Get-MimecastMessageInfo -Id $message.id

# Download message file
Get-MimecastMessageFile -Id $message.id -OutputPath 'C:\Temp\message.eml'

#
# SIEM Log Collection Setup
#

# IMPORTANT: Prerequisites
# 1. Create custom API role with "SIEM Integration > SIEM Events > Read" permissions
# 2. Create service account with "Log on as a batch job" rights
# 3. Store API credentials using Set-MimecastSecret
# 4. Note: Mimecast retains SIEM events for 7 days

# Define paths and settings
$scriptPath = "C:\Program Files\Mimecast\Scripts\Invoke-MimecastLogCollection.ps1"
$stateDir = "C:\ProgramData\Mimecast\State"
$logDir = "C:\ProgramData\Mimecast\Logs"

# Ensure directories exist
New-Item -Path $stateDir, $logDir -ItemType Directory -Force

# Create service account credential
$svcPassword = Read-Host -Prompt "Enter service account password" -AsSecureString
$svcCredential = New-Object System.Management.Automation.PSCredential("DOMAIN\MimecastSvc", $svcPassword)

# Register the SIEM log collector task for Customer Gateway events (MTA logs)
Register-MimecastLogCollectorTask `
    -TaskName "MimecastSIEMLogCollector_CG" `
    -ScriptPath $scriptPath `
    -SIEMEndpoint "syslog.example.com" `
    -Type "receipt" `
    -PageSize 100 `
    -StartDate (Get-Date).AddDays(-7) `
    -LogOutputDirectory $logDir `
    -StateDirectory $stateDir `
    -Frequency "EveryNMinutes" `
    -IntervalMinutes 5 `
    -Credential $svcCredential `
    -Force

# Register the SIEM log collector task for Customer Intelligence events
Register-MimecastLogCollectorTask `
    -TaskName "MimecastSIEMLogCollector_CI" `
    -ScriptPath $scriptPath `
    -SIEMEndpoint "syslog.example.com" `
    -Type "entities" `
    -PageSize 100 `
    -StartDate (Get-Date).AddDays(-7) `
    -LogOutputDirectory $logDir `
    -StateDirectory $stateDir `
    -Frequency "EveryNMinutes" `
    -IntervalMinutes 5 `
    -Credential $svcCredential `
    -Force

# View registered tasks
Get-MimecastLogCollectorTask -TaskName "MimecastSIEMLogCollector_CG"
Get-MimecastLogCollectorTask -TaskName "MimecastSIEMLogCollector_CI"

# Start the tasks immediately
Start-MimecastLogCollectorTask -TaskName "MimecastSIEMLogCollector_CG"
Start-MimecastLogCollectorTask -TaskName "MimecastSIEMLogCollector_CI"

# Wait a bit and check task status
Start-Sleep -Seconds 10
Get-MimecastLogCollectorTask -TaskName "MimecastSIEMLogCollector_CG"
Get-MimecastLogCollectorTask -TaskName "MimecastSIEMLogCollector_CI"

# Stop the tasks if needed
Stop-MimecastLogCollectorTask -TaskName "MimecastSIEMLogCollector_CG"
Stop-MimecastLogCollectorTask -TaskName "MimecastSIEMLogCollector_CI"

# Remove the tasks when no longer needed
# Unregister-MimecastLogCollectorTask -TaskName "MimecastSIEMLogCollector_CG"
# Unregister-MimecastLogCollectorTask -TaskName "MimecastSIEMLogCollector_CI"

# Disconnect when done
Disconnect-Mimecast