# Common rule types for content examination
$script:ContentRuleTypes = @{
    Keyword = @{
        Type = 'keyword'
        RequiredFields = @('pattern', 'matchType')
        OptionalFields = @('caseSensitive', 'wholeWord')
    }
    Pattern = @{
        Type = 'pattern'
        RequiredFields = @('pattern', 'matchType')
        OptionalFields = @('caseSensitive')
    }
    Dictionary = @{
        Type = 'dictionary'
        RequiredFields = @('dictionaryId')
        OptionalFields = @('matchType')
    }
    Size = @{
        Type = 'size'
        RequiredFields = @('size', 'operator')
        OptionalFields = @()
    }
    Attachment = @{
        Type = 'attachment'
        RequiredFields = @('fileType')
        OptionalFields = @('includeArchives')
    }
}

# Common notification settings parameters
$script:NotificationSettingsParams = @(
    @{
        Name = 'NotificationGroupId'
        Type = 'string'
        Mandatory = $false
        JsonName = 'notificationGroupId'
    }
    @{
        Name = 'NotifySender'
        Type = 'bool'
        Mandatory = $false
        JsonName = 'notifySender'
    }
    @{
        Name = 'NotifyRecipient'
        Type = 'bool'
        Mandatory = $false
        JsonName = 'notifyRecipient'
    }
    @{
        Name = 'NotifyAdministrator'
        Type = 'bool'
        Mandatory = $false
        JsonName = 'notifyAdministrator'
    }
    @{
        Name = 'CustomNotificationSubject'
        Type = 'string'
        Mandatory = $false
        JsonName = 'customNotificationSubject'
    }
    @{
        Name = 'CustomNotificationBody'
        Type = 'string'
        Mandatory = $false
        JsonName = 'customNotificationBody'
    }
)

function New-MimecastAddressAlterationDefinition {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$Description,

        [Parameter(Mandatory)]
        [string]$OriginalAddress,

        [Parameter(Mandatory)]
        [string]$NewAddress,

        [Parameter(Mandatory)]
        [ValidateSet('envelope_from', 'header_from', 'both')]
        [string]$AddressType,

        [Parameter(Mandatory)]
        [ValidateSet('inbound', 'outbound')]
        [string]$Routing,

        # Common notification settings
        [Parameter()]
        [string]$NotificationGroupId,

        [Parameter()]
        [bool]$NotifySender,

        [Parameter()]
        [bool]$NotifyRecipient,

        [Parameter()]
        [bool]$NotifyAdministrator,

        [Parameter()]
        [string]$CustomNotificationSubject,

        [Parameter()]
        [string]$CustomNotificationBody
    )

    try {
        # Build request body
        $body = @{
            data = @(
                @{
                    description = $Description
                    addressAlterations = @(
                        @{
                            originalAddress = $OriginalAddress
                            newAddress = $NewAddress
                            addressType = $AddressType
                            routing = $Routing
                        }
                    )
                }
            )
        }

        # Add notification settings if provided
        $notificationSettings = @{}
        $hasNotificationSettings = $false

        foreach ($param in $script:NotificationSettingsParams) {
            $value = $PSBoundParameters[$param.Name]
            if ($null -ne $value) {
                $notificationSettings[$param.JsonName] = $value
                $hasNotificationSettings = $true
            }
        }

        if ($hasNotificationSettings) {
            $body.data[0].notificationSettings = $notificationSettings
        }

        # Make API request
        $response = Invoke-ApiRequest -Method 'POST' -Path '/api/policy/address-alteration/create-definition' -Body $body
        $response.data
    }
    catch {
        Write-Error -ErrorRecord $_
    }
}

function Get-MimecastAddressAlterationDefinition {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]$Id,

        [Parameter()]
        [string]$Description,

        [Parameter()]
        [string]$OriginalAddress,

        [Parameter()]
        [string]$NewAddress
    )

    try {
        # Build request body
        $body = @{
            data = @(
                @{}
            )
        }

        if ($Id) {
            $body.data[0].id = $Id
        }
        if ($Description) {
            $body.data[0].description = $Description
        }
        if ($OriginalAddress) {
            $body.data[0].originalAddress = $OriginalAddress
        }
        if ($NewAddress) {
            $body.data[0].newAddress = $NewAddress
        }

        # Make API request
        $response = Invoke-ApiRequest -Method 'POST' -Path '/api/policy/address-alteration/get-definition' -Body $body
        $response.data
    }
    catch {
        Write-Error -ErrorRecord $_
    }
}

<#
.SYNOPSIS
Creates a new content examination definition.

.DESCRIPTION
Creates a new content examination definition in Mimecast. These definitions contain
the rules and actions used by content examination policies for DLP and content filtering.
The definition can include multiple rules of different types and specify actions to
take when content matches the rules.

.PARAMETER Description
A description of the definition. This helps identify the definition's purpose.

.PARAMETER Action
The action to take when content matches. Valid values:
- reject: Block and reject the message
- hold: Hold for review
- tag: Add tags/headers
- bcc: Send a copy to specified addresses

.PARAMETER ScanLocations
Optional. Which parts of the message to scan. Valid values:
- subject: Message subject
- body: Message body
- attachments: File attachments
- headers: Email headers
Default is @('subject', 'body').

.PARAMETER MatchCondition
Optional. How multiple rules are combined. Valid values:
- any: Match if any rule matches (OR)
- all: Match only if all rules match (AND)
Default is 'any'.

.PARAMETER Rules
Optional. Array of content examination rules. Each rule should be a hashtable with:
- type: Rule type (keyword, pattern, dictionary, size, attachment)
- Required fields based on type (see $ContentRuleTypes)
- Optional fields based on type

.PARAMETER NotificationGroupId
Optional. The notification group to use.

.PARAMETER NotifySender
Optional. Whether to notify the sender.

.PARAMETER NotifyRecipient
Optional. Whether to notify the recipient.

.PARAMETER NotifyAdministrator
Optional. Whether to notify administrators.

.PARAMETER CustomNotificationSubject
Optional. Custom notification subject template.

.PARAMETER CustomNotificationBody
Optional. Custom notification body template.

.INPUTS
None. You cannot pipe objects to New-MimecastContentExaminationDefinition.

.OUTPUTS
System.Object
Returns the newly created definition object.

.EXAMPLE
C:\> $rules = @(
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
    }
)
C:\> New-MimecastContentExaminationDefinition `
    -Description 'DLP - Confidential Content' `
    -Action 'hold' `
    -ScanLocations @('subject', 'body', 'attachments') `
    -MatchCondition 'any' `
    -Rules $rules `
    -NotifyAdministrator $true

Creates a content examination definition for detecting confidential content.

.EXAMPLE
C:\> $rules = @(
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
C:\> New-MimecastContentExaminationDefinition `
    -Description 'Block Large Files and Executables' `
    -Action 'reject' `
    -ScanLocations @('attachments') `
    -Rules $rules `
    -NotifySender $true `
    -CustomNotificationSubject 'Message Rejected: {subject}' `
    -CustomNotificationBody 'Your message was rejected due to attachment restrictions.'

Creates a content examination definition for blocking large files and executables.

.NOTES
- Rules must be valid according to $ContentRuleTypes
- Consider testing with a limited scope first
- Use with New-MimecastPolicy to apply the definition
#>
function New-MimecastContentExaminationDefinition {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$Description,

        [Parameter(Mandatory)]
        [ValidateSet('reject', 'hold', 'tag', 'bcc')]
        [string]$Action,

        [Parameter()]
        [ValidateSet('subject', 'body', 'attachments', 'headers')]
        [string[]]$ScanLocations = @('subject', 'body'),

        [Parameter()]
        [ValidateSet('any', 'all')]
        [string]$MatchCondition = 'any',

        [Parameter()]
        [object[]]$Rules,

        # Common notification settings
        [Parameter()]
        [string]$NotificationGroupId,

        [Parameter()]
        [bool]$NotifySender,

        [Parameter()]
        [bool]$NotifyRecipient,

        [Parameter()]
        [bool]$NotifyAdministrator,

        [Parameter()]
        [string]$CustomNotificationSubject,

        [Parameter()]
        [string]$CustomNotificationBody
    )

    try {
        # Validate rules
        if ($Rules) {
            foreach ($rule in $Rules) {
                if (-not $rule.type -or -not $script:ContentRuleTypes[$rule.type]) {
                    throw "Invalid rule type: $($rule.type)"
                }

                $ruleType = $script:ContentRuleTypes[$rule.type]
                foreach ($field in $ruleType.RequiredFields) {
                    if (-not $rule.$field) {
                        throw "Missing required field '$field' for rule type '$($rule.type)'"
                    }
                }
            }
        }

        # Build request body
        $body = @{
            data = @(
                @{
                    description = $Description
                    action = $Action
                    scanLocations = $ScanLocations
                    matchCondition = $MatchCondition
                }
            )
        }

        if ($Rules) {
            $body.data[0].rules = $Rules
        }

        # Add notification settings if provided
        $notificationSettings = @{}
        $hasNotificationSettings = $false

        foreach ($param in $script:NotificationSettingsParams) {
            $value = $PSBoundParameters[$param.Name]
            if ($null -ne $value) {
                $notificationSettings[$param.JsonName] = $value
                $hasNotificationSettings = $true
            }
        }

        if ($hasNotificationSettings) {
            $body.data[0].notificationSettings = $notificationSettings
        }

        # Make API request
        $response = Invoke-ApiRequest -Method 'POST' -Path '/api/policy/content-examination/create-definition' -Body $body
        $response.data
    }
    catch {
        Write-Error -ErrorRecord $_
    }
}

<#
.SYNOPSIS
Updates an existing content examination definition.

.DESCRIPTION
Modifies an existing content examination definition in Mimecast. This function allows
you to update the rules, actions, and notification settings for content-based DLP policies.
Only the specified parameters will be updated; all other settings remain unchanged.

.PARAMETER Id
The unique identifier of the definition to update. This parameter accepts pipeline input
by property name from Get-MimecastContentExaminationDefinition.

.PARAMETER Description
Optional. A new description for the definition.

.PARAMETER Action
Optional. Update the action to take when content matches. Valid values:
- reject: Block and reject the message
- hold: Hold for review
- tag: Add tags/headers
- bcc: Send a copy to specified addresses

.PARAMETER ScanLocations
Optional. Update which parts of the message to scan. Valid values:
- subject: Message subject
- body: Message body
- attachments: File attachments
- headers: Email headers

.PARAMETER MatchCondition
Optional. Update how multiple rules are combined. Valid values:
- any: Match if any rule matches (OR)
- all: Match only if all rules match (AND)

.PARAMETER Rules
Optional. Update the array of content examination rules. Each rule should be a hashtable with:
- type: Rule type (keyword, pattern, dictionary, size, attachment)
- Required fields based on type (see $ContentRuleTypes)
- Optional fields based on type

.PARAMETER NotificationGroupId
Optional. Update the notification group to use.

.PARAMETER NotifySender
Optional. Update whether to notify the sender.

.PARAMETER NotifyRecipient
Optional. Update whether to notify the recipient.

.PARAMETER NotifyAdministrator
Optional. Update whether to notify administrators.

.PARAMETER CustomNotificationSubject
Optional. Update the custom notification subject template.

.PARAMETER CustomNotificationBody
Optional. Update the custom notification body template.

.INPUTS
System.Object. You can pipe definition objects from Get-MimecastContentExaminationDefinition.

.OUTPUTS
System.Object
Returns the updated definition object.

.EXAMPLE
C:\> $rules = @(
    @{
        type = 'keyword'
        pattern = 'confidential'
        matchType = 'contains'
        caseSensitive = $false
    }
)
C:\> Get-MimecastContentExaminationDefinition -Description 'Old DLP Rules' |
    Set-MimecastContentExaminationDefinition -Description 'Updated DLP Rules' -Rules $rules

Updates the rules for an existing content examination definition.

.EXAMPLE
C:\> Set-MimecastContentExaminationDefinition -Id '4d24b932-5c57-4b57-b7b5-9c3f511e5c5c' `
    -Action 'hold' `
    -NotifyAdministrator $true `
    -CustomNotificationSubject 'DLP Alert: {subject}'

Updates the action and notification settings for a specific definition.

.NOTES
- Rules must be valid according to $ContentRuleTypes
- Consider testing changes with a limited scope first
- Changes take effect immediately for associated policies
#>
function Set-MimecastContentExaminationDefinition {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]$Id,

        [Parameter()]
        [string]$Description,

        [Parameter()]
        [ValidateSet('reject', 'hold', 'tag', 'bcc')]
        [string]$Action,

        [Parameter()]
        [ValidateSet('subject', 'body', 'attachments', 'headers')]
        [string[]]$ScanLocations,

        [Parameter()]
        [ValidateSet('any', 'all')]
        [string]$MatchCondition,

        [Parameter()]
        [object[]]$Rules,

        # Common notification settings
        [Parameter()]
        [string]$NotificationGroupId,

        [Parameter()]
        [bool]$NotifySender,

        [Parameter()]
        [bool]$NotifyRecipient,

        [Parameter()]
        [bool]$NotifyAdministrator,

        [Parameter()]
        [string]$CustomNotificationSubject,

        [Parameter()]
        [string]$CustomNotificationBody
    )

    process {
        try {
            # Validate rules if provided
            if ($Rules) {
                foreach ($rule in $Rules) {
                    if (-not $rule.type -or -not $script:ContentRuleTypes[$rule.type]) {
                        throw "Invalid rule type: $($rule.type)"
                    }

                    $ruleType = $script:ContentRuleTypes[$rule.type]
                    foreach ($field in $ruleType.RequiredFields) {
                        if (-not $rule.$field) {
                            throw "Missing required field '$field' for rule type '$($rule.type)'"
                        }
                    }
                }
            }

            # Build request body
            $body = @{
                data = @(
                    @{
                        id = $Id
                    }
                )
            }

            # Add optional fields
            if ($Description) {
                $body.data[0].description = $Description
            }
            if ($Action) {
                $body.data[0].action = $Action
            }
            if ($ScanLocations) {
                $body.data[0].scanLocations = $ScanLocations
            }
            if ($MatchCondition) {
                $body.data[0].matchCondition = $MatchCondition
            }
            if ($Rules) {
                $body.data[0].rules = $Rules
            }

            # Add notification settings if provided
            $notificationSettings = @{}
            $hasNotificationSettings = $false

            foreach ($param in $script:NotificationSettingsParams) {
                $value = $PSBoundParameters[$param.Name]
                if ($null -ne $value) {
                    $notificationSettings[$param.JsonName] = $value
                    $hasNotificationSettings = $true
                }
            }

            if ($hasNotificationSettings) {
                $body.data[0].notificationSettings = $notificationSettings
            }

            # Make API request
            $response = Invoke-ApiRequest -Method 'POST' -Path '/api/policy/content-examination/update-definition' -Body $body
            $response.data
        }
        catch {
            Write-Error -ErrorRecord $_
        }
    }
}

<#
.SYNOPSIS
Gets content examination definitions based on optional filters.

.DESCRIPTION
Retrieves content examination definitions from Mimecast. These definitions contain
the rules and actions used by content examination policies for DLP and content filtering.

.PARAMETER Id
Optional. The unique identifier of a specific definition to retrieve.

.PARAMETER Description
Optional. Filter definitions by description (exact match).

.INPUTS
None. You cannot pipe objects to Get-MimecastContentExaminationDefinition.

.OUTPUTS
System.Object[]
Returns an array of content examination definition objects.

.EXAMPLE
C:\> Get-MimecastContentExaminationDefinition
Gets all content examination definitions.

.EXAMPLE
C:\> Get-MimecastContentExaminationDefinition -Id '4d24b932-5c57-4b57-b7b5-9c3f511e5c5c'
Gets a specific content examination definition by ID.

.EXAMPLE
C:\> Get-MimecastContentExaminationDefinition -Description 'DLP Rules'
Gets content examination definitions with the exact description 'DLP Rules'.

.NOTES
- Definitions are used by content examination policies
- Each definition can contain multiple rules
- Rules can be of different types (keyword, pattern, dictionary, etc.)
#>
function Get-MimecastContentExaminationDefinition {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]$Id,

        [Parameter()]
        [string]$Description,

        [Parameter()]
        [ValidateSet('reject', 'hold', 'tag', 'bcc')]
        [string]$Action
    )

    try {
        # Build request body
        $body = @{
            data = @(
                @{}
            )
        }

        if ($Id) {
            $body.data[0].id = $Id
        }
        if ($Description) {
            $body.data[0].description = $Description
        }
        if ($Action) {
            $body.data[0].action = $Action
        }

        # Make API request
        $response = Invoke-ApiRequest -Method 'POST' -Path '/api/policy/content-examination/get-definition' -Body $body
        $response.data
    }
    catch {
        Write-Error -ErrorRecord $_
    }
}

<#
.SYNOPSIS
Removes a content examination definition.

.DESCRIPTION
Deletes a content examination definition from Mimecast. This operation is permanent
and cannot be undone. The function supports pipeline input from Get-MimecastContentExaminationDefinition
and includes ShouldProcess support for confirmation.

.PARAMETER Id
The unique identifier of the definition to remove. This parameter accepts pipeline input
by property name from Get-MimecastContentExaminationDefinition.

.INPUTS
System.Object. You can pipe definition objects from Get-MimecastContentExaminationDefinition.

.OUTPUTS
System.Object
Returns the result of the removal operation.

.EXAMPLE
C:\> Remove-MimecastContentExaminationDefinition -Id '4d24b932-5c57-4b57-b7b5-9c3f511e5c5c'

Removes a specific content examination definition.

.EXAMPLE
C:\> Get-MimecastContentExaminationDefinition -Description 'Old DLP Rules' |
    Remove-MimecastContentExaminationDefinition

Removes content examination definitions with the description 'Old DLP Rules'.

.EXAMPLE
C:\> Get-MimecastContentExaminationDefinition |
    Where-Object { $_.action -eq 'reject' } |
    Remove-MimecastContentExaminationDefinition -WhatIf

Shows what would happen if you removed all content examination definitions with 'reject' action.

.NOTES
- This operation cannot be undone
- Associated policies will no longer function
- Use -WhatIf to preview changes before executing
#>
function Remove-MimecastContentExaminationDefinition {
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]$Id
    )

    process {
        try {
            if ($PSCmdlet.ShouldProcess($Id, "Remove Mimecast content examination definition")) {
                # Build request body
                $body = @{
                    data = @(
                        @{
                            id = $Id
                        }
                    )
                }

                # Make API request
                $response = Invoke-ApiRequest -Method 'POST' -Path '/api/policy/content-examination/delete-definition' -Body $body
                $response.data
            }
        }
        catch {
            Write-Error -ErrorRecord $_
        }
    }
}

<#
.SYNOPSIS
Creates a new attachment management definition.

.DESCRIPTION
Creates a new attachment management definition in Mimecast. These definitions specify
how different types of attachments should be handled, including file type filtering,
size limits, and encrypted file handling. The definition can include lists of allowed
and blocked content types.

.PARAMETER Description
A description of the definition. This helps identify the definition's purpose.

.PARAMETER DefaultBlockAllow
The default action for attachments not explicitly allowed or blocked. Valid values:
- block: Block attachments by default
- allow: Allow attachments by default
Default is 'block'.

.PARAMETER PornographicImageSetting
How to handle suspected pornographic images. Valid values:
- block: Block messages containing such images
- allow: Allow messages containing such images
- scan: Scan but take no action
Default is 'block'.

.PARAMETER EncryptedArchives
How to handle encrypted archive files. Valid values:
- block: Block encrypted archives
- allow: Allow encrypted archives
Default is 'block'.

.PARAMETER EncryptedDocuments
How to handle encrypted documents. Valid values:
- block: Block encrypted documents
- allow: Allow encrypted documents
Default is 'block'.

.PARAMETER UnreadableArchives
How to handle archives that cannot be scanned. Valid values:
- block: Block unreadable archives
- allow: Allow unreadable archives
Default is 'block'.

.PARAMETER LargeFileThreshold
Optional. Size in bytes above which files are considered "large".
Default is 10485760 (10MB).

.PARAMETER EnableLargeFileSubstitution
Optional. Whether to replace large files with a link.
Default is $false.

.PARAMETER AllowedContent
Optional. Array of content types to allow. Each item should be a hashtable with:
- type: Content type (e.g., 'file_extension', 'mime_type')
- value: The actual type value (e.g., 'pdf', 'application/pdf')

.PARAMETER BlockedContent
Optional. Array of content types to block. Each item should be a hashtable with:
- type: Content type (e.g., 'file_extension', 'mime_type')
- value: The actual type value (e.g., 'exe', 'application/x-msdownload')

.PARAMETER NotificationGroupId
Optional. The notification group to use.

.PARAMETER NotifySender
Optional. Whether to notify the sender.

.PARAMETER NotifyRecipient
Optional. Whether to notify the recipient.

.PARAMETER NotifyAdministrator
Optional. Whether to notify administrators.

.PARAMETER CustomNotificationSubject
Optional. Custom notification subject template.

.PARAMETER CustomNotificationBody
Optional. Custom notification body template.

.INPUTS
None. You cannot pipe objects to New-MimecastAttachmentManagementDefinition.

.OUTPUTS
System.Object
Returns the newly created definition object.

.EXAMPLE
C:\> $allowedContent = @(
    @{
        type = 'file_extension'
        value = 'pdf'
    },
    @{
        type = 'file_extension'
        value = 'docx'
    }
)
C:\> $blockedContent = @(
    @{
        type = 'file_extension'
        value = 'exe'
    },
    @{
        type = 'file_extension'
        value = 'bat'
    }
)
C:\> New-MimecastAttachmentManagementDefinition `
    -Description 'Standard Attachment Policy' `
    -DefaultBlockAllow 'block' `
    -AllowedContent $allowedContent `
    -BlockedContent $blockedContent `
    -LargeFileThreshold 10485760 `
    -NotifyAdministrator $true

Creates an attachment management definition that allows PDF and DOCX files,
blocks executables, and notifies administrators.

.EXAMPLE
C:\> New-MimecastAttachmentManagementDefinition `
    -Description 'Strict Security Policy' `
    -DefaultBlockAllow 'block' `
    -EncryptedArchives 'block' `
    -EncryptedDocuments 'block' `
    -UnreadableArchives 'block' `
    -PornographicImageSetting 'block' `
    -NotifySender $true `
    -CustomNotificationSubject 'Blocked Attachment: {subject}' `
    -CustomNotificationBody 'Your message contained a blocked attachment type.'

Creates a strict attachment management definition that blocks all potentially
dangerous content types and notifies senders.

.NOTES
- Consider security implications when allowing file types
- Test with a limited scope first
- Use with New-MimecastPolicy to apply the definition
#>
function New-MimecastAttachmentManagementDefinition {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$Description,

        [Parameter()]
        [ValidateSet('block', 'allow')]
        [string]$DefaultBlockAllow = 'block',

        [Parameter()]
        [ValidateSet('block', 'allow', 'scan')]
        [string]$PornographicImageSetting = 'block',

        [Parameter()]
        [ValidateSet('block', 'allow')]
        [string]$EncryptedArchives = 'block',

        [Parameter()]
        [ValidateSet('block', 'allow')]
        [string]$EncryptedDocuments = 'block',

        [Parameter()]
        [ValidateSet('block', 'allow')]
        [string]$UnreadableArchives = 'block',

        [Parameter()]
        [int]$LargeFileThreshold = 10485760, # 10MB

        [Parameter()]
        [bool]$EnableLargeFileSubstitution = $false,

        [Parameter()]
        [object[]]$AllowedContent,

        [Parameter()]
        [object[]]$BlockedContent,

        # Common notification settings
        [Parameter()]
        [string]$NotificationGroupId,

        [Parameter()]
        [bool]$NotifySender,

        [Parameter()]
        [bool]$NotifyRecipient,

        [Parameter()]
        [bool]$NotifyAdministrator,

        [Parameter()]
        [string]$CustomNotificationSubject,

        [Parameter()]
        [string]$CustomNotificationBody
    )

    try {
        # Build request body
        $body = @{
            data = @(
                @{
                    description = $Description
                    defaultBlockAllow = $DefaultBlockAllow
                    pornographicImageSetting = $PornographicImageSetting
                    encryptedArchives = $EncryptedArchives
                    encryptedDocuments = $EncryptedDocuments
                    unreadableArchives = $UnreadableArchives
                    largeFileThreshold = $LargeFileThreshold
                    enableLargeFileSubstitution = $EnableLargeFileSubstitution
                }
            )
        }

        if ($AllowedContent) {
            $body.data[0].allowedContent = $AllowedContent
        }

        if ($BlockedContent) {
            $body.data[0].blockedContent = $BlockedContent
        }

        # Add notification settings if provided
        $notificationSettings = @{}
        $hasNotificationSettings = $false

        foreach ($param in $script:NotificationSettingsParams) {
            $value = $PSBoundParameters[$param.Name]
            if ($null -ne $value) {
                $notificationSettings[$param.JsonName] = $value
                $hasNotificationSettings = $true
            }
        }

        if ($hasNotificationSettings) {
            $body.data[0].notificationSettings = $notificationSettings
        }

        # Make API request
        $response = Invoke-ApiRequest -Method 'POST' -Path '/api/policy/attachment-management/create-definition' -Body $body
        $response.data
    }
    catch {
        Write-Error -ErrorRecord $_
    }
}

<#
.SYNOPSIS
Gets attachment management definitions based on optional filters.

.DESCRIPTION
Retrieves attachment management definitions from Mimecast. These definitions specify
how different types of attachments should be handled, including file type filtering,
size limits, and encrypted file handling.

.PARAMETER Id
Optional. The unique identifier of a specific definition to retrieve.

.PARAMETER Description
Optional. Filter definitions by description (exact match).

.INPUTS
None. You cannot pipe objects to Get-MimecastAttachmentManagementDefinition.

.OUTPUTS
System.Object[]
Returns an array of attachment management definition objects.

.EXAMPLE
C:\> Get-MimecastAttachmentManagementDefinition
Gets all attachment management definitions.

.EXAMPLE
C:\> Get-MimecastAttachmentManagementDefinition -Id '4d24b932-5c57-4b57-b7b5-9c3f511e5c5c'
Gets a specific attachment management definition by ID.

.EXAMPLE
C:\> Get-MimecastAttachmentManagementDefinition -Description 'Block Executables'
Gets attachment management definitions with the exact description 'Block Executables'.

.NOTES
- Definitions are used by attachment management policies
- Each definition can have allowed and blocked content lists
- Definitions control various attachment handling settings
#>
function Get-MimecastAttachmentManagementDefinition {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]$Id,

        [Parameter()]
        [string]$Description
    )

    try {
        # Build request body
        $body = @{
            data = @(
                @{}
            )
        }

        if ($Id) {
            $body.data[0].id = $Id
        }
        if ($Description) {
            $body.data[0].description = $Description
        }

        # Make API request
        $response = Invoke-ApiRequest -Method 'POST' -Path '/api/policy/attachment-management/get-definition' -Body $body
        $response.data
    }
    catch {
        Write-Error -ErrorRecord $_
    }
}

<#
.SYNOPSIS
Removes an attachment management definition.

.DESCRIPTION
Deletes an attachment management definition from Mimecast. This operation is permanent
and cannot be undone. The function supports pipeline input from Get-MimecastAttachmentManagementDefinition
and includes ShouldProcess support for confirmation.

.PARAMETER Id
The unique identifier of the definition to remove. This parameter accepts pipeline input
by property name from Get-MimecastAttachmentManagementDefinition.

.INPUTS
System.Object. You can pipe definition objects from Get-MimecastAttachmentManagementDefinition.

.OUTPUTS
System.Object
Returns the result of the removal operation.

.EXAMPLE
C:\> Remove-MimecastAttachmentManagementDefinition -Id '4d24b932-5c57-4b57-b7b5-9c3f511e5c5c'

Removes a specific attachment management definition.

.EXAMPLE
C:\> Get-MimecastAttachmentManagementDefinition -Description 'Old Rules' |
    Remove-MimecastAttachmentManagementDefinition

Removes attachment management definitions with the description 'Old Rules'.

.EXAMPLE
C:\> Get-MimecastAttachmentManagementDefinition |
    Where-Object { $_.defaultBlockAllow -eq 'allow' } |
    Remove-MimecastAttachmentManagementDefinition -WhatIf

Shows what would happen if you removed all attachment management definitions
that allow attachments by default.

.NOTES
- This operation cannot be undone
- Associated policies will no longer function
- Use -WhatIf to preview changes before executing
#>
function Remove-MimecastAttachmentManagementDefinition {
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]$Id
    )

    process {
        try {
            if ($PSCmdlet.ShouldProcess($Id, "Remove Mimecast attachment management definition")) {
                # Build request body
                $body = @{
                    data = @(
                        @{
                            id = $Id
                        }
                    )
                }

                # Make API request
                $response = Invoke-ApiRequest -Method 'POST' -Path '/api/policy/attachment-management/delete-definition' -Body $body
                $response.data
            }
        }
        catch {
            Write-Error -ErrorRecord $_
        }
    }
}

<#
.SYNOPSIS
Creates a new spam scanner definition.

.DESCRIPTION
Creates a new spam scanner definition in Mimecast. These definitions specify how
spam detection should be configured, including detection levels, actions to take,
and various protection features. The definition can also include notification
settings for when spam is detected.

.PARAMETER Description
A description of the definition. This helps identify the definition's purpose.

.PARAMETER SpamDetectionLevel
Optional. How aggressively to detect spam. Valid values:
- relaxed: More permissive, fewer false positives
- moderate: Balanced detection
- aggressive: Stricter detection, may have more false positives
Default is 'moderate'.

.PARAMETER ActionIfSpamDetected
Optional. What to do when spam is detected. Valid values:
- tag_headers: Add spam headers but deliver
- hold: Hold message for review
- reject: Block and reject the message
Default is 'tag_headers'.

.PARAMETER SpamScoreThreshold
Optional. Score above which messages are considered spam.
Default is 5.

.PARAMETER ActionIfSpamScoreHit
Optional. What to do when spam score threshold is hit. Valid values:
- tag_headers: Add spam headers but deliver
- hold: Hold message for review
- reject: Block and reject the message
Default is 'hold'.

.PARAMETER EnableSpamFeedback
Optional. Whether to enable spam feedback loop.
Default is $true.

.PARAMETER EnableGreylistingProtection
Optional. Whether to enable greylisting protection.
Default is $true.

.PARAMETER EnableIntelligentRejection
Optional. Whether to enable intelligent rejection.
Default is $true.

.PARAMETER NotificationGroupId
Optional. The notification group to use.

.PARAMETER NotifySender
Optional. Whether to notify the sender.

.PARAMETER NotifyRecipient
Optional. Whether to notify the recipient.

.PARAMETER NotifyAdministrator
Optional. Whether to notify administrators.

.PARAMETER CustomNotificationSubject
Optional. Custom notification subject template.

.PARAMETER CustomNotificationBody
Optional. Custom notification body template.

.INPUTS
None. You cannot pipe objects to New-MimecastSpamScannerDefinition.

.OUTPUTS
System.Object
Returns the newly created definition object.

.EXAMPLE
C:\> New-MimecastSpamScannerDefinition `
    -Description 'Aggressive Spam Policy' `
    -SpamDetectionLevel 'aggressive' `
    -ActionIfSpamDetected 'hold' `
    -SpamScoreThreshold 7 `
    -ActionIfSpamScoreHit 'hold' `
    -NotifyAdministrator $true

Creates an aggressive spam scanner definition that holds suspicious messages
and notifies administrators.

.EXAMPLE
C:\> New-MimecastSpamScannerDefinition `
    -Description 'Basic Spam Tagging' `
    -SpamDetectionLevel 'moderate' `
    -ActionIfSpamDetected 'tag_headers' `
    -EnableSpamFeedback $true `
    -EnableGreylistingProtection $true `
    -EnableIntelligentRejection $true

Creates a moderate spam scanner definition that tags messages and enables
all protection features.

.NOTES
- Consider false positive rates when setting detection level
- Test with a limited scope first
- Use with New-MimecastPolicy to apply the definition
#>
function New-MimecastSpamScannerDefinition {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$Description,

        [Parameter()]
        [ValidateSet('relaxed', 'moderate', 'aggressive')]
        [string]$SpamDetectionLevel = 'moderate',

        [Parameter()]
        [ValidateSet('tag_headers', 'hold', 'reject')]
        [string]$ActionIfSpamDetected = 'tag_headers',

        [Parameter()]
        [int]$SpamScoreThreshold = 5,

        [Parameter()]
        [ValidateSet('tag_headers', 'hold', 'reject')]
        [string]$ActionIfSpamScoreHit = 'hold',

        [Parameter()]
        [bool]$EnableSpamFeedback = $true,

        [Parameter()]
        [bool]$EnableGreylistingProtection = $true,

        [Parameter()]
        [bool]$EnableIntelligentRejection = $true,

        # Common notification settings
        [Parameter()]
        [string]$NotificationGroupId,

        [Parameter()]
        [bool]$NotifySender,

        [Parameter()]
        [bool]$NotifyRecipient,

        [Parameter()]
        [bool]$NotifyAdministrator,

        [Parameter()]
        [string]$CustomNotificationSubject,

        [Parameter()]
        [string]$CustomNotificationBody
    )

    try {
        # Build request body
        $body = @{
            data = @(
                @{
                    description = $Description
                    spamDetectionLevel = $SpamDetectionLevel
                    actionIfSpamDetected = $ActionIfSpamDetected
                    spamScoreThreshold = $SpamScoreThreshold
                    actionIfSpamScoreHit = $ActionIfSpamScoreHit
                    enableSpamFeedback = $EnableSpamFeedback
                    enableGreylistingProtection = $EnableGreylistingProtection
                    enableIntelligentRejection = $EnableIntelligentRejection
                }
            )
        }

        # Add notification settings if provided
        $notificationSettings = @{}
        $hasNotificationSettings = $false

        foreach ($param in $script:NotificationSettingsParams) {
            $value = $PSBoundParameters[$param.Name]
            if ($null -ne $value) {
                $notificationSettings[$param.JsonName] = $value
                $hasNotificationSettings = $true
            }
        }

        if ($hasNotificationSettings) {
            $body.data[0].notificationSettings = $notificationSettings
        }

        # Make API request
        $response = Invoke-ApiRequest -Method 'POST' -Path '/api/policy/spam-scanning/create-definition' -Body $body
        $response.data
    }
    catch {
        Write-Error -ErrorRecord $_
    }
}

<#
.SYNOPSIS
Gets spam scanner definitions based on optional filters.

.DESCRIPTION
Retrieves spam scanner definitions from Mimecast. These definitions specify how
spam detection should be configured, including detection levels, actions to take,
and various protection features.

.PARAMETER Id
Optional. The unique identifier of a specific definition to retrieve.

.PARAMETER Description
Optional. Filter definitions by description (exact match).

.INPUTS
None. You cannot pipe objects to Get-MimecastSpamScannerDefinition.

.OUTPUTS
System.Object[]
Returns an array of spam scanner definition objects.

.EXAMPLE
C:\> Get-MimecastSpamScannerDefinition
Gets all spam scanner definitions.

.EXAMPLE
C:\> Get-MimecastSpamScannerDefinition -Id '4d24b932-5c57-4b57-b7b5-9c3f511e5c5c'
Gets a specific spam scanner definition by ID.

.EXAMPLE
C:\> Get-MimecastSpamScannerDefinition -Description 'Aggressive Spam Policy'
Gets spam scanner definitions with the exact description 'Aggressive Spam Policy'.

.NOTES
- Definitions are used by spam scanning policies
- Each definition controls spam detection behavior
- Definitions can specify different actions based on spam scores
#>
function Get-MimecastSpamScannerDefinition {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]$Id,

        [Parameter()]
        [string]$Description
    )

    try {
        # Build request body
        $body = @{
            data = @(
                @{}
            )
        }

        if ($Id) {
            $body.data[0].id = $Id
        }
        if ($Description) {
            $body.data[0].description = $Description
        }

        # Make API request
        $response = Invoke-ApiRequest -Method 'POST' -Path '/api/policy/spam-scanning/get-definition' -Body $body
        $response.data
    }
    catch {
        Write-Error -ErrorRecord $_
    }
}

<#
.SYNOPSIS
Removes a spam scanner definition.

.DESCRIPTION
Deletes a spam scanner definition from Mimecast. This operation is permanent
and cannot be undone. The function supports pipeline input from Get-MimecastSpamScannerDefinition
and includes ShouldProcess support for confirmation.

.PARAMETER Id
The unique identifier of the definition to remove. This parameter accepts pipeline input
by property name from Get-MimecastSpamScannerDefinition.

.INPUTS
System.Object. You can pipe definition objects from Get-MimecastSpamScannerDefinition.

.OUTPUTS
System.Object
Returns the result of the removal operation.

.EXAMPLE
C:\> Remove-MimecastSpamScannerDefinition -Id '4d24b932-5c57-4b57-b7b5-9c3f511e5c5c'

Removes a specific spam scanner definition.

.EXAMPLE
C:\> Get-MimecastSpamScannerDefinition -Description 'Old Rules' |
    Remove-MimecastSpamScannerDefinition

Removes spam scanner definitions with the description 'Old Rules'.

.EXAMPLE
C:\> Get-MimecastSpamScannerDefinition |
    Where-Object { $_.spamDetectionLevel -eq 'aggressive' } |
    Remove-MimecastSpamScannerDefinition -WhatIf

Shows what would happen if you removed all spam scanner definitions
with aggressive detection level.

.NOTES
- This operation cannot be undone
- Associated policies will no longer function
- Use -WhatIf to preview changes before executing
#>
function Remove-MimecastSpamScannerDefinition {
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]$Id
    )

    process {
        try {
            if ($PSCmdlet.ShouldProcess($Id, "Remove Mimecast spam scanner definition")) {
                # Build request body
                $body = @{
                    data = @(
                        @{
                            id = $Id
                        }
                    )
                }

                # Make API request
                $response = Invoke-ApiRequest -Method 'POST' -Path '/api/policy/spam-scanning/delete-definition' -Body $body
                $response.data
            }
        }
        catch {
            Write-Error -ErrorRecord $_
        }
    }
}

<#
.SYNOPSIS
Creates a new notification set definition.

.DESCRIPTION
Creates a new notification set definition in Mimecast. These definitions specify
the templates and settings for various types of email notifications sent by Mimecast,
including delivery receipts, hold notifications, and policy-specific notifications.

.PARAMETER Description
A description of the definition. This helps identify the definition's purpose.

.PARAMETER BrandingId
Optional. The ID of the branding template to use.

.PARAMETER PreferredLanguage
Optional. The language to use for notifications. Valid values:
- en: English
- de: German
- es: Spanish
- fr: French
- ja: Japanese
- nl: Dutch
- pt-br: Brazilian Portuguese
- zh-cn: Simplified Chinese
Default is 'en'.

.PARAMETER UserMessage
Optional. A custom message to include in all notifications.

.PARAMETER DeliveryReceipts
Optional. Settings for delivery receipt notifications. Should be a hashtable with:
- subject: Template for the subject line
- body: Template for the message body

.PARAMETER HoldNotifications
Optional. Settings for hold notifications. Should be a hashtable with:
- subject: Template for the subject line
- body: Template for the message body

.PARAMETER RejectionNotifications
Optional. Settings for rejection notifications. Should be a hashtable with:
- subject: Template for the subject line
- body: Template for the message body

.PARAMETER PermittedSenderNotifications
Optional. Settings for permitted sender notifications. Should be a hashtable with:
- subject: Template for the subject line
- body: Template for the message body

.PARAMETER BlockedSenderNotifications
Optional. Settings for blocked sender notifications. Should be a hashtable with:
- subject: Template for the subject line
- body: Template for the message body

.PARAMETER ContentExaminationNotifications
Optional. Settings for content examination notifications. Should be a hashtable with:
- subject: Template for the subject line
- body: Template for the message body

.PARAMETER AttachmentManagementNotifications
Optional. Settings for attachment management notifications. Should be a hashtable with:
- subject: Template for the subject line
- body: Template for the message body

.PARAMETER SpamNotifications
Optional. Settings for spam notifications. Should be a hashtable with:
- subject: Template for the subject line
- body: Template for the message body

.INPUTS
None. You cannot pipe objects to New-MimecastNotificationSetDefinition.

.OUTPUTS
System.Object
Returns the newly created definition object.

.EXAMPLE
C:\> $deliveryReceipts = @{
    subject = 'Message Delivered: {subject}'
    body = @'
Your message "{subject}" was delivered successfully.
From: {from}
To: {to}
Sent: {sent}
'@
}
C:\> $holdNotifications = @{
    subject = 'Message Held: {subject}'
    body = @'
Your message "{subject}" was held for review.
From: {from}
To: {to}
Sent: {sent}
Reason: {reason}
'@
}
C:\> New-MimecastNotificationSetDefinition `
    -Description 'Standard Notification Set' `
    -PreferredLanguage 'en' `
    -UserMessage 'This is an automated notification from the email security system.' `
    -DeliveryReceipts $deliveryReceipts `
    -HoldNotifications $holdNotifications

Creates a notification set definition with custom templates for delivery receipts
and hold notifications.

.EXAMPLE
C:\> $spamNotifications = @{
    subject = 'Spam Detected: {subject}'
    body = @'
A message was identified as spam and {action}.
From: {from}
To: {to}
Subject: {subject}
Spam Score: {score}
'@
}
C:\> New-MimecastNotificationSetDefinition `
    -Description 'Spam Notifications' `
    -PreferredLanguage 'en' `
    -SpamNotifications $spamNotifications

Creates a notification set definition specifically for spam notifications.

.NOTES
- Templates support various placeholders (e.g., {subject}, {from}, {to})
- Consider branding consistency across notifications
- Test notifications in a limited scope first
#>
function New-MimecastNotificationSetDefinition {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$Description,

        [Parameter()]
        [string]$BrandingId,

        [Parameter()]
        [ValidateSet('en', 'de', 'es', 'fr', 'ja', 'nl', 'pt-br', 'zh-cn')]
        [string]$PreferredLanguage = 'en',

        [Parameter()]
        [string]$UserMessage,

        [Parameter()]
        [object]$DeliveryReceipts,

        [Parameter()]
        [object]$HoldNotifications,

        [Parameter()]
        [object]$RejectionNotifications,

        [Parameter()]
        [object]$PermittedSenderNotifications,

        [Parameter()]
        [object]$BlockedSenderNotifications,

        [Parameter()]
        [object]$ContentExaminationNotifications,

        [Parameter()]
        [object]$AttachmentManagementNotifications,

        [Parameter()]
        [object]$SpamNotifications
    )

    try {
        # Build request body
        $body = @{
            data = @(
                @{
                    description = $Description
                    preferredLanguage = $PreferredLanguage
                }
            )
        }

        if ($BrandingId) {
            $body.data[0].brandingId = $BrandingId
        }
        if ($UserMessage) {
            $body.data[0].userMessage = $UserMessage
        }
        if ($DeliveryReceipts) {
            $body.data[0].deliveryReceipts = $DeliveryReceipts
        }
        if ($HoldNotifications) {
            $body.data[0].holdNotifications = $HoldNotifications
        }
        if ($RejectionNotifications) {
            $body.data[0].rejectionNotifications = $RejectionNotifications
        }
        if ($PermittedSenderNotifications) {
            $body.data[0].permittedSenderNotifications = $PermittedSenderNotifications
        }
        if ($BlockedSenderNotifications) {
            $body.data[0].blockedSenderNotifications = $BlockedSenderNotifications
        }
        if ($ContentExaminationNotifications) {
            $body.data[0].contentExaminationNotifications = $ContentExaminationNotifications
        }
        if ($AttachmentManagementNotifications) {
            $body.data[0].attachmentManagementNotifications = $AttachmentManagementNotifications
        }
        if ($SpamNotifications) {
            $body.data[0].spamNotifications = $SpamNotifications
        }

        # Make API request
        $response = Invoke-ApiRequest -Method 'POST' -Path '/api/notification-set/create-definition' -Body $body
        $response.data
    }
    catch {
        Write-Error -ErrorRecord $_
    }
}

<#
.SYNOPSIS
Gets notification set definitions based on optional filters.

.DESCRIPTION
Retrieves notification set definitions from Mimecast. These definitions specify
the templates and settings for various types of email notifications sent by Mimecast,
including delivery receipts, hold notifications, and policy-specific notifications.

.PARAMETER Id
Optional. The unique identifier of a specific definition to retrieve.

.PARAMETER Description
Optional. Filter definitions by description (exact match).

.INPUTS
None. You cannot pipe objects to Get-MimecastNotificationSetDefinition.

.OUTPUTS
System.Object[]
Returns an array of notification set definition objects.

.EXAMPLE
C:\> Get-MimecastNotificationSetDefinition
Gets all notification set definitions.

.EXAMPLE
C:\> Get-MimecastNotificationSetDefinition -Id '4d24b932-5c57-4b57-b7b5-9c3f511e5c5c'
Gets a specific notification set definition by ID.

.EXAMPLE
C:\> Get-MimecastNotificationSetDefinition -Description 'Standard Notifications'
Gets notification set definitions with the exact description 'Standard Notifications'.

.NOTES
- Definitions contain templates for various notification types
- Each definition can have different language settings
- Templates support placeholders for dynamic content
#>
function Get-MimecastNotificationSetDefinition {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]$Id,

        [Parameter()]
        [string]$Description
    )

    try {
        # Build request body
        $body = @{
            data = @(
                @{}
            )
        }

        if ($Id) {
            $body.data[0].id = $Id
        }
        if ($Description) {
            $body.data[0].description = $Description
        }

        # Make API request
        $response = Invoke-ApiRequest -Method 'POST' -Path '/api/notification-set/get-definition' -Body $body
        $response.data
    }
    catch {
        Write-Error -ErrorRecord $_
    }
}

<#
.SYNOPSIS
Removes a notification set definition.

.DESCRIPTION
Deletes a notification set definition from Mimecast. This operation is permanent
and cannot be undone. The function supports pipeline input from Get-MimecastNotificationSetDefinition
and includes ShouldProcess support for confirmation.

.PARAMETER Id
The unique identifier of the definition to remove. This parameter accepts pipeline input
by property name from Get-MimecastNotificationSetDefinition.

.INPUTS
System.Object. You can pipe definition objects from Get-MimecastNotificationSetDefinition.

.OUTPUTS
System.Object
Returns the result of the removal operation.

.EXAMPLE
C:\> Remove-MimecastNotificationSetDefinition -Id '4d24b932-5c57-4b57-b7b5-9c3f511e5c5c'

Removes a specific notification set definition.

.EXAMPLE
C:\> Get-MimecastNotificationSetDefinition -Description 'Old Templates' |
    Remove-MimecastNotificationSetDefinition

Removes notification set definitions with the description 'Old Templates'.

.EXAMPLE
C:\> Get-MimecastNotificationSetDefinition |
    Where-Object { $_.preferredLanguage -eq 'en' } |
    Remove-MimecastNotificationSetDefinition -WhatIf

Shows what would happen if you removed all English notification set definitions.

.NOTES
- This operation cannot be undone
- Associated policies will no longer function
- Use -WhatIf to preview changes before executing
#>
function Remove-MimecastNotificationSetDefinition {
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]$Id
    )

    process {
        try {
            if ($PSCmdlet.ShouldProcess($Id, "Remove Mimecast notification set definition")) {
                # Build request body
                $body = @{
                    data = @(
                        @{
                            id = $Id
                        }
                    )
                }

                # Make API request
                $response = Invoke-ApiRequest -Method 'POST' -Path '/api/notification-set/delete-definition' -Body $body
                $response.data
            }
        }
        catch {
            Write-Error -ErrorRecord $_
        }
    }
}

function Remove-MimecastAddressAlterationDefinition {
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]$Id
    )

    process {
        try {
            if ($PSCmdlet.ShouldProcess($Id, "Remove Mimecast address alteration definition")) {
                # Build request body
                $body = @{
                    data = @(
                        @{
                            id = $Id
                        }
                    )
                }

                # Make API request
                $response = Invoke-ApiRequest -Method 'POST' -Path '/api/policy/address-alteration/delete-definition' -Body $body
                $response.data
            }
        }
        catch {
            Write-Error -ErrorRecord $_
        }
    }
}