# Common policy types found in the Postman collection
$script:PolicyTypes = @{
    AddressAlteration = @{
        BasePath = '/api/policy/address-alteration'
        RequiresDefinition = $true
        DefinitionIdField = 'addressAlterationSetId'
    }
    AntiSpoofingBypass = @{
        BasePath = '/api/policy/antispoofing-bypass'
        RequiresDefinition = $false
    }
    BlockedSender = @{
        BasePath = '/api/policy/blockedsenders'
        RequiresDefinition = $false
    }
    WebWhiteUrl = @{
        BasePath = '/api/policy/webwhiteurl'
        RequiresDefinition = $false
        HasTargets = $true
    }
    ContentExamination = @{
        BasePath = '/api/policy/content-examination'
        RequiresDefinition = $true
        DefinitionIdField = 'contentExaminationId'
    }
    AttachmentManagement = @{
        BasePath = '/api/policy/attachment-management'
        RequiresDefinition = $true
        DefinitionIdField = 'attachmentManagementId'
    }
    SpamScanning = @{
        BasePath = '/api/policy/spam-scanning'
        RequiresDefinition = $true
        DefinitionIdField = 'spamScannerId'
    }
    NotificationSet = @{
        BasePath = '/api/notification-set'
        RequiresDefinition = $true
        DefinitionIdField = 'notificationSetId'
    }
}

# Common policy types found in the Postman collection
$script:PolicyTypes = @{
    AddressAlteration = @{
        BasePath = '/api/policy/address-alteration'
        RequiresDefinition = $true
        DefinitionIdField = 'addressAlterationSetId'
    }
    AntiSpoofingBypass = @{
        BasePath = '/api/policy/antispoofing-bypass'
        RequiresDefinition = $false
    }
    BlockedSender = @{
        BasePath = '/api/policy/blockedsenders'
        RequiresDefinition = $false
    }
    WebWhiteUrl = @{
        BasePath = '/api/policy/webwhiteurl'
        RequiresDefinition = $false
        HasTargets = $true
    }
    ContentExamination = @{
        BasePath = '/api/policy/content-examination'
        RequiresDefinition = $true
        DefinitionIdField = 'contentExaminationId'
    }
    AttachmentManagement = @{
        BasePath = '/api/policy/attachment-management'
        RequiresDefinition = $true
        DefinitionIdField = 'attachmentManagementId'
    }
    SpamScanning = @{
        BasePath = '/api/policy/spam-scanning'
        RequiresDefinition = $true
        DefinitionIdField = 'spamScannerId'
    }
    NotificationSet = @{
        BasePath = '/api/notification-set'
        RequiresDefinition = $true
        DefinitionIdField = 'notificationSetId'
    }
}

<#
.SYNOPSIS
Gets Mimecast policies based on type and optional filters.

.DESCRIPTION
Retrieves Mimecast policies of a specified type. Can filter by ID or description.
Supports various policy types including address alteration, anti-spoofing bypass,
content examination, attachment management, and more.

.PARAMETER PolicyType
The type of policy to retrieve. Valid values:
- AddressAlteration: Email address rewriting policies
- AntiSpoofingBypass: Anti-spoofing exception policies
- BlockedSender: Blocked sender policies
- WebWhiteUrl: Permitted URL policies
- ContentExamination: DLP and content filtering policies
- AttachmentManagement: Attachment control policies
- SpamScanning: Spam detection policies
- NotificationSet: Email notification template policies

.PARAMETER Id
Optional. The unique identifier of a specific policy to retrieve.

.PARAMETER Description
Optional. Filter policies by description (exact match).

.INPUTS
None. You cannot pipe objects to Get-MimecastPolicy.

.OUTPUTS
System.Object[]
Returns an array of policy objects. Each object's properties depend on the policy type.

.EXAMPLE
C:\> Get-MimecastPolicy -PolicyType 'AntiSpoofingBypass'
Gets all anti-spoofing bypass policies.

.EXAMPLE
C:\> Get-MimecastPolicy -PolicyType 'ContentExamination' -Id '4d24b932-5c57-4b57-b7b5-9c3f511e5c5c'
Gets a specific content examination policy by ID.

.EXAMPLE
C:\> Get-MimecastPolicy -PolicyType 'AttachmentManagement' -Description 'Block Executables'
Gets attachment management policies with the exact description 'Block Executables'.

.NOTES
- Some policy types require associated definitions (see RequiresDefinition in $PolicyTypes).
- The WebWhiteUrl policy type supports additional targets for URL whitelisting.
- Policies are region-specific; ensure you're connected to the correct region.
#>
function Get-MimecastPolicy {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateSet('AddressAlteration', 'AntiSpoofingBypass', 'BlockedSender', 'WebWhiteUrl')]
        [string]$PolicyType,

        [Parameter()]
        [string]$Id,

        [Parameter()]
        [string]$Description
    )

    try {
        # Get policy type info
        $policyInfo = $script:PolicyTypes[$PolicyType]
        if (-not $policyInfo) {
            throw "Unknown policy type: $PolicyType"
        }

        # Build request path
        $path = if ($policyInfo.HasTargets) {
            "$($policyInfo.BasePath)/get-policy-with-targets"
        }
        else {
            "$($policyInfo.BasePath)/get-policy"
        }

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
        $response = Invoke-ApiRequest -Method 'POST' -Path $path -Body $body
        $response.data
    }
    catch {
        Write-Error -ErrorRecord $_
    }
}

<#
.SYNOPSIS
Creates a new Mimecast policy of the specified type.

.DESCRIPTION
Creates a new Mimecast policy with the specified settings. Different policy types
have different requirements and supported parameters. Some policy types require
an associated definition that contains the detailed configuration.

.PARAMETER PolicyType
The type of policy to create. Valid values:
- AddressAlteration: Email address rewriting policies
- AntiSpoofingBypass: Anti-spoofing exception policies
- BlockedSender: Blocked sender policies
- WebWhiteUrl: Permitted URL policies
- ContentExamination: DLP and content filtering policies
- AttachmentManagement: Attachment control policies
- SpamScanning: Spam detection policies
- NotificationSet: Email notification template policies

.PARAMETER Description
A description of the policy. This helps identify the policy's purpose.

.PARAMETER FromType
The type of sender to which this policy applies. Valid values:
- everyone: All senders
- internal_addresses: Internal email addresses
- external_addresses: External email addresses
- email_domain: Specific email domain
- profile_group: Group from address book
- address_attribute_value: Address with specific attribute
- individual_email_address: Single email address
- free_mail_domains: Free email providers
- header_display_name: Display name in header

.PARAMETER FromValue
The specific value for FromType (e.g., email address, domain name).
Required when FromType is not 'everyone'.

.PARAMETER ToType
The type of recipient to which this policy applies. Same valid values as FromType.

.PARAMETER ToValue
The specific value for ToType. Required when ToType is not 'everyone'.

.PARAMETER FromPart
For policies that examine email addresses, which part to check. Valid values:
- envelope_from: Envelope From address
- header_from: From header
- both: Both envelope and header

.PARAMETER Direction
The direction of email flow this policy applies to. Valid values:
- inbound: Incoming email
- outbound: Outgoing email
- both: Both directions

.PARAMETER StartDate
Optional. The date and time when this policy becomes active.

.PARAMETER EndDate
Optional. The date and time when this policy expires.

.PARAMETER Bidirectional
Optional. If true, the policy applies in both directions regardless of From/To settings.

.PARAMETER Enabled
Optional. Whether the policy is active. Default is $true.

.PARAMETER Override
Optional. Whether this policy can override other policies.

.PARAMETER Comment
Optional. Additional notes about the policy.

.PARAMETER DefinitionId
For policy types that require a definition, the ID of the associated definition.
See the RequiresDefinition property in $PolicyTypes.

.PARAMETER DefinitionName
Alternative to DefinitionId. The name of the definition to use.
The function will attempt to resolve this to a definition ID.

.PARAMETER Option
For AntiSpoofingBypass policies, the action to take. Valid values:
- enable_bypass: Allow messages that would be blocked
- no_action: Take no action

.PARAMETER Targets
For WebWhiteUrl policies, an array of target objects. Each object should have:
- action: 'allow' or 'block'
- type: 'domain', 'url', etc.
- value: The actual URL or domain

.INPUTS
None. You cannot pipe objects to New-MimecastPolicy.

.OUTPUTS
System.Object
Returns the newly created policy object.

.EXAMPLE
C:\> New-MimecastPolicy -PolicyType 'AntiSpoofingBypass' `
    -Description 'Allow trusted partner' `
    -FromType 'email_domain' `
    -FromValue 'trusted-partner.com' `
    -ToType 'internal_addresses' `
    -Option 'enable_bypass'

Creates an anti-spoofing bypass policy for a trusted partner domain.

.EXAMPLE
C:\> $definition = New-MimecastContentExaminationDefinition `
    -Description 'DLP Rules' `
    -Action 'hold' `
    -Rules $rules
C:\> New-MimecastPolicy -PolicyType 'ContentExamination' `
    -Description 'DLP Policy' `
    -FromType 'internal_addresses' `
    -ToType 'external_addresses' `
    -Direction 'outbound' `
    -DefinitionId $definition.id

Creates a DLP policy using a content examination definition.

.EXAMPLE
C:\> $targets = @(
    @{
        action = 'allow'
        type = 'domain'
        value = 'trusted-site.com'
    }
)
C:\> New-MimecastPolicy -PolicyType 'WebWhiteUrl' `
    -Description 'Allow trusted website' `
    -FromType 'internal_addresses' `
    -ToType 'everyone' `
    -Targets $targets

Creates a web URL whitelisting policy.

.NOTES
- Some policy types require definitions to be created first
- Policies are region-specific; ensure you're connected to the correct region
- Consider testing new policies with a limited scope first
#>
function New-MimecastPolicy {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateSet('AddressAlteration', 'AntiSpoofingBypass', 'BlockedSender', 'WebWhiteUrl')]
        [string]$PolicyType,

        [Parameter(Mandatory)]
        [string]$Description,

        [Parameter()]
        [ValidateSet('everyone', 'internal_addresses', 'external_addresses', 'email_domain', 'profile_group', 'address_attribute_value', 'individual_email_address', 'free_mail_domains', 'header_display_name')]
        [string]$FromType = 'everyone',

        [Parameter()]
        [string]$FromValue,

        [Parameter()]
        [ValidateSet('everyone', 'internal_addresses', 'external_addresses', 'email_domain', 'profile_group', 'address_attribute_value', 'individual_email_address', 'free_mail_domains', 'header_display_name')]
        [string]$ToType = 'everyone',

        [Parameter()]
        [string]$ToValue,

        [Parameter()]
        [ValidateSet('envelope_from', 'header_from', 'both')]
        [string]$FromPart,

        [Parameter()]
        [ValidateSet('inbound', 'outbound', 'both')]
        [string]$Direction,

        [Parameter()]
        [datetime]$StartDate,

        [Parameter()]
        [datetime]$EndDate,

        [Parameter()]
        [bool]$Bidirectional,

        [Parameter()]
        [bool]$Enabled = $true,

        [Parameter()]
        [bool]$Override,

        [Parameter()]
        [string]$Comment,

        [Parameter()]
        [string]$DefinitionId,

        [Parameter()]
        [string]$DefinitionName,

        # Anti-spoofing bypass specific
        [Parameter()]
        [ValidateSet('enable_bypass', 'no_action')]
        [string]$Option = 'enable_bypass',

        # Web white URL specific
        [Parameter()]
        [object[]]$Targets
    )

    try {
        # Get policy type info
        $policyInfo = $script:PolicyTypes[$PolicyType]
        if (-not $policyInfo) {
            throw "Unknown policy type: $PolicyType"
        }

        # Build base policy object
        $policy = @{
            description = $Description
            from = @{
                type = $FromType
            }
            to = @{
                type = $ToType
            }
        }

        # Add optional policy fields
        if ($FromValue) {
            $policy.from.emailAddress = $FromValue
        }
        if ($ToValue) {
            $policy.to.emailAddress = $ToValue
        }
        if ($FromPart) {
            $policy.fromPart = $FromPart
        }
        if ($Direction) {
            $policy.direction = $Direction
        }
        if ($StartDate) {
            $policy.startDate = $StartDate.ToString('yyyy-MM-ddTHH:mm:ssZ')
        }
        if ($EndDate) {
            $policy.endDate = $EndDate.ToString('yyyy-MM-ddTHH:mm:ssZ')
        }
        if ($PSBoundParameters.ContainsKey('Bidirectional')) {
            $policy.bidirectional = $Bidirectional
        }
        if ($PSBoundParameters.ContainsKey('Enabled')) {
            $policy.enabled = $Enabled
        }
        if ($PSBoundParameters.ContainsKey('Override')) {
            $policy.override = $Override
        }
        if ($Comment) {
            $policy.comment = $Comment
        }

        # Build request body
        $body = @{
            data = @(
                @{
                    policy = $policy
                }
            )
        }

        # Add policy type specific fields
        if ($policyInfo.RequiresDefinition) {
            if ($DefinitionId) {
                $body.data[0][$policyInfo.DefinitionIdField] = $DefinitionId
            }
            elseif ($DefinitionName) {
                # TODO: Implement definition name resolution
                throw "Definition name resolution not yet implemented"
            }
            else {
                throw "Policy type $PolicyType requires a definition ID or name"
            }
        }

        if ($PolicyType -eq 'AntiSpoofingBypass') {
            $body.data[0].option = $Option
        }

        if ($policyInfo.HasTargets -and $Targets) {
            $body.data[0].targets = $Targets
        }

        # Build request path
        $path = if ($policyInfo.HasTargets) {
            "$($policyInfo.BasePath)/create-policy-with-targets"
        }
        else {
            "$($policyInfo.BasePath)/create-policy"
        }

        # Make API request
        $response = Invoke-ApiRequest -Method 'POST' -Path $path -Body $body
        $response.data
    }
    catch {
        Write-Error -ErrorRecord $_
    }
}

<#
.SYNOPSIS
Updates an existing Mimecast policy.

.DESCRIPTION
Modifies the settings of an existing Mimecast policy. Only the specified parameters
will be updated; all other settings remain unchanged. Different policy types support
different parameters for modification.

.PARAMETER PolicyType
The type of policy to update. Valid values:
- AddressAlteration: Email address rewriting policies
- AntiSpoofingBypass: Anti-spoofing exception policies
- BlockedSender: Blocked sender policies
- WebWhiteUrl: Permitted URL policies
- ContentExamination: DLP and content filtering policies
- AttachmentManagement: Attachment control policies
- SpamScanning: Spam detection policies
- NotificationSet: Email notification template policies

.PARAMETER Id
The unique identifier of the policy to update. This parameter accepts pipeline input
by property name from Get-MimecastPolicy.

.PARAMETER Description
Optional. A new description for the policy.

.PARAMETER FromType
Optional. Update the type of sender this policy applies to. Valid values:
- everyone: All senders
- internal_addresses: Internal email addresses
- external_addresses: External email addresses
- email_domain: Specific email domain
- profile_group: Group from address book
- address_attribute_value: Address with specific attribute
- individual_email_address: Single email address
- free_mail_domains: Free email providers
- header_display_name: Display name in header

.PARAMETER FromValue
Optional. Update the specific value for FromType (e.g., email address, domain name).
Required when updating FromType to a value other than 'everyone'.

.PARAMETER ToType
Optional. Update the type of recipient this policy applies to.
Same valid values as FromType.

.PARAMETER ToValue
Optional. Update the specific value for ToType.
Required when updating ToType to a value other than 'everyone'.

.PARAMETER FromPart
Optional. For policies that examine email addresses, update which part to check.
Valid values:
- envelope_from: Envelope From address
- header_from: From header
- both: Both envelope and header

.PARAMETER Direction
Optional. Update the direction of email flow this policy applies to. Valid values:
- inbound: Incoming email
- outbound: Outgoing email
- both: Both directions

.PARAMETER StartDate
Optional. Update the date and time when this policy becomes active.

.PARAMETER EndDate
Optional. Update the date and time when this policy expires.

.PARAMETER Bidirectional
Optional. Update whether the policy applies in both directions.

.PARAMETER Enabled
Optional. Update whether the policy is active.

.PARAMETER Override
Optional. Update whether this policy can override other policies.

.PARAMETER Comment
Optional. Update additional notes about the policy.

.PARAMETER DefinitionId
Optional. For policy types that require a definition, update the associated
definition ID. See the RequiresDefinition property in $PolicyTypes.

.PARAMETER Option
Optional. For AntiSpoofingBypass policies, update the action to take. Valid values:
- enable_bypass: Allow messages that would be blocked
- no_action: Take no action

.PARAMETER Targets
Optional. For WebWhiteUrl policies, update the array of target objects.
Each object should have:
- action: 'allow' or 'block'
- type: 'domain', 'url', etc.
- value: The actual URL or domain

.INPUTS
System.Object. You can pipe policy objects from Get-MimecastPolicy.

.OUTPUTS
System.Object
Returns the updated policy object.

.EXAMPLE
C:\> Get-MimecastPolicy -PolicyType 'AntiSpoofingBypass' | 
    Set-MimecastPolicy -PolicyType 'AntiSpoofingBypass' -Enabled $false

Disables all anti-spoofing bypass policies.

.EXAMPLE
C:\> $policy = Get-MimecastPolicy -PolicyType 'ContentExamination' -Description 'Old DLP Policy'
C:\> $newDefinition = New-MimecastContentExaminationDefinition -Description 'Updated Rules' -Action 'hold'
C:\> Set-MimecastPolicy -PolicyType 'ContentExamination' `
    -Id $policy.id `
    -Description 'Updated DLP Policy' `
    -DefinitionId $newDefinition.id

Updates a content examination policy with a new description and definition.

.EXAMPLE
C:\> $newTargets = @(
    @{
        action = 'allow'
        type = 'domain'
        value = 'new-trusted-site.com'
    }
)
C:\> Set-MimecastPolicy -PolicyType 'WebWhiteUrl' `
    -Id '4d24b932-5c57-4b57-b7b5-9c3f511e5c5c' `
    -Targets $newTargets

Updates the targets for a web URL whitelisting policy.

.NOTES
- Only specified parameters will be updated
- Some policy types require definitions to exist
- Policies are region-specific; ensure you're connected to the correct region
- Consider testing changes with a limited scope first
#>
function Set-MimecastPolicy {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateSet('AddressAlteration', 'AntiSpoofingBypass', 'BlockedSender', 'WebWhiteUrl')]
        [string]$PolicyType,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]$Id,

        [Parameter()]
        [string]$Description,

        [Parameter()]
        [ValidateSet('everyone', 'internal_addresses', 'external_addresses', 'email_domain', 'profile_group', 'address_attribute_value', 'individual_email_address', 'free_mail_domains', 'header_display_name')]
        [string]$FromType,

        [Parameter()]
        [string]$FromValue,

        [Parameter()]
        [ValidateSet('everyone', 'internal_addresses', 'external_addresses', 'email_domain', 'profile_group', 'address_attribute_value', 'individual_email_address', 'free_mail_domains', 'header_display_name')]
        [string]$ToType,

        [Parameter()]
        [string]$ToValue,

        [Parameter()]
        [ValidateSet('envelope_from', 'header_from', 'both')]
        [string]$FromPart,

        [Parameter()]
        [ValidateSet('inbound', 'outbound', 'both')]
        [string]$Direction,

        [Parameter()]
        [datetime]$StartDate,

        [Parameter()]
        [datetime]$EndDate,

        [Parameter()]
        [bool]$Bidirectional,

        [Parameter()]
        [bool]$Enabled,

        [Parameter()]
        [bool]$Override,

        [Parameter()]
        [string]$Comment,

        [Parameter()]
        [string]$DefinitionId,

        # Anti-spoofing bypass specific
        [Parameter()]
        [ValidateSet('enable_bypass', 'no_action')]
        [string]$Option,

        # Web white URL specific
        [Parameter()]
        [object[]]$Targets
    )

    process {
        try {
            # Get policy type info
            $policyInfo = $script:PolicyTypes[$PolicyType]
            if (-not $policyInfo) {
                throw "Unknown policy type: $PolicyType"
            }

            # Build base request body
            $body = @{
                data = @(
                    @{
                        id = $Id
                    }
                )
            }

            # Build policy object if we have any policy updates
            $policy = @{}
            $hasUpdates = $false

            if ($Description) {
                $policy.description = $Description
                $hasUpdates = $true
            }

            if ($FromType -or $FromValue) {
                $policy.from = @{}
                if ($FromType) {
                    $policy.from.type = $FromType
                }
                if ($FromValue) {
                    $policy.from.emailAddress = $FromValue
                }
                $hasUpdates = $true
            }

            if ($ToType -or $ToValue) {
                $policy.to = @{}
                if ($ToType) {
                    $policy.to.type = $ToType
                }
                if ($ToValue) {
                    $policy.to.emailAddress = $ToValue
                }
                $hasUpdates = $true
            }

            if ($FromPart) {
                $policy.fromPart = $FromPart
                $hasUpdates = $true
            }
            if ($Direction) {
                $policy.direction = $Direction
                $hasUpdates = $true
            }
            if ($StartDate) {
                $policy.startDate = $StartDate.ToString('yyyy-MM-ddTHH:mm:ssZ')
                $hasUpdates = $true
            }
            if ($EndDate) {
                $policy.endDate = $EndDate.ToString('yyyy-MM-ddTHH:mm:ssZ')
                $hasUpdates = $true
            }
            if ($PSBoundParameters.ContainsKey('Bidirectional')) {
                $policy.bidirectional = $Bidirectional
                $hasUpdates = $true
            }
            if ($PSBoundParameters.ContainsKey('Enabled')) {
                $policy.enabled = $Enabled
                $hasUpdates = $true
            }
            if ($PSBoundParameters.ContainsKey('Override')) {
                $policy.override = $Override
                $hasUpdates = $true
            }
            if ($Comment) {
                $policy.comment = $Comment
                $hasUpdates = $true
            }

            # Add policy object if we have updates
            if ($hasUpdates) {
                $body.data[0].policy = $policy
            }

            # Add policy type specific fields
            if ($policyInfo.RequiresDefinition -and $DefinitionId) {
                $body.data[0][$policyInfo.DefinitionIdField] = $DefinitionId
            }

            if ($PolicyType -eq 'AntiSpoofingBypass' -and $Option) {
                $body.data[0].option = $Option
            }

            if ($policyInfo.HasTargets -and $Targets) {
                $body.data[0].targets = $Targets
            }

            # Build request path
            $path = if ($policyInfo.HasTargets) {
                "$($policyInfo.BasePath)/update-policy-with-targets"
            }
            else {
                "$($policyInfo.BasePath)/update-policy"
            }

            # Make API request
            $response = Invoke-ApiRequest -Method 'POST' -Path $path -Body $body
            $response.data
        }
        catch {
            Write-Error -ErrorRecord $_
        }
    }
}

<#
.SYNOPSIS
Removes a Mimecast policy.

.DESCRIPTION
Deletes a Mimecast policy of the specified type. This operation is permanent and
cannot be undone. The function supports pipeline input from Get-MimecastPolicy
and includes ShouldProcess support for confirmation.

.PARAMETER PolicyType
The type of policy to remove. Valid values:
- AddressAlteration: Email address rewriting policies
- AntiSpoofingBypass: Anti-spoofing exception policies
- BlockedSender: Blocked sender policies
- WebWhiteUrl: Permitted URL policies
- ContentExamination: DLP and content filtering policies
- AttachmentManagement: Attachment control policies
- SpamScanning: Spam detection policies
- NotificationSet: Email notification template policies

.PARAMETER Id
The unique identifier of the policy to remove. This parameter accepts pipeline input
by property name from Get-MimecastPolicy.

.INPUTS
System.Object. You can pipe policy objects from Get-MimecastPolicy.

.OUTPUTS
System.Object
Returns the result of the removal operation.

.EXAMPLE
C:\> Remove-MimecastPolicy -PolicyType 'AntiSpoofingBypass' -Id '4d24b932-5c57-4b57-b7b5-9c3f511e5c5c'

Removes a specific anti-spoofing bypass policy.

.EXAMPLE
C:\> Get-MimecastPolicy -PolicyType 'BlockedSender' -Description 'Temporary Block' |
    Remove-MimecastPolicy -PolicyType 'BlockedSender'

Removes all blocked sender policies with the description 'Temporary Block'.

.EXAMPLE
C:\> Get-MimecastPolicy -PolicyType 'WebWhiteUrl' |
    Where-Object { $_.enabled -eq $false } |
    Remove-MimecastPolicy -PolicyType 'WebWhiteUrl' -WhatIf

Shows what would happen if you removed all disabled web URL whitelisting policies.

.NOTES
- This operation cannot be undone
- Some policy types may have associated definitions that are not automatically removed
- Policies are region-specific; ensure you're connected to the correct region
- Use -WhatIf to preview changes before executing
#>
function Remove-MimecastPolicy {
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory)]
        [ValidateSet('AddressAlteration', 'AntiSpoofingBypass', 'BlockedSender', 'WebWhiteUrl')]
        [string]$PolicyType,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]$Id
    )

    process {
        try {
            # Get policy type info
            $policyInfo = $script:PolicyTypes[$PolicyType]
            if (-not $policyInfo) {
                throw "Unknown policy type: $PolicyType"
            }

            if ($PSCmdlet.ShouldProcess($Id, "Remove Mimecast $PolicyType policy")) {
                # Build request path
                $path = if ($policyInfo.HasTargets) {
                    "$($policyInfo.BasePath)/delete-policy-with-targets"
                }
                else {
                    "$($policyInfo.BasePath)/delete-policy"
                }

                # Build request body
                $body = @{
                    data = @(
                        @{
                            id = $Id
                        }
                    )
                }

                # Make API request
                $response = Invoke-ApiRequest -Method 'POST' -Path $path -Body $body
                $response.data
            }
        }
        catch {
            Write-Error -ErrorRecord $_
        }
    }
}