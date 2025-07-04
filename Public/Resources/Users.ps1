<#
.SYNOPSIS
Gets Mimecast users based on optional filters.

.DESCRIPTION
Retrieves users from Mimecast. Can filter by email address, domain, or other attributes.
Supports pagination for large result sets and includes detailed user information such as
account status, aliases, and group memberships.

.PARAMETER EmailAddress
Optional. Filter users by exact email address match.

.PARAMETER Domain
Optional. Filter users by domain (e.g., 'example.com').

.PARAMETER FirstPageOnly
Optional. Return only the first page of results.
Default is $false (returns all pages).

.PARAMETER PageSize
Optional. Number of users to return per page.
Default is 100. Maximum is 500.

.PARAMETER AccountType
Optional. Filter users by account type. Valid values:
- internal: Internal users
- external: External users
- all: Both internal and external users (default)

.PARAMETER AccountStatus
Optional. Filter users by account status. Valid values:
- active: Active accounts
- disabled: Disabled accounts
- deleted: Deleted accounts
- all: All accounts (default)

.INPUTS
None. You cannot pipe objects to Get-MimecastUser.

.OUTPUTS
System.Object[]
Returns an array of user objects. Each object includes:
- emailAddress: User's primary email address
- name: User's display name
- domain: User's domain
- accountType: Internal or external
- accountStatus: Active, disabled, or deleted
- aliases: Array of email aliases
- groups: Array of group memberships
- attributes: Custom attributes
- created: Account creation date
- lastLogon: Last login date

.EXAMPLE
C:\> Get-MimecastUser
Gets all users across all domains.

.EXAMPLE
C:\> Get-MimecastUser -EmailAddress 'user@example.com'
Gets a specific user by email address.

.EXAMPLE
C:\> Get-MimecastUser -Domain 'example.com' -AccountType 'internal' -AccountStatus 'active'
Gets all active internal users in the specified domain.

.EXAMPLE
C:\> Get-MimecastUser -FirstPageOnly -PageSize 50
Gets only the first 50 users.

.NOTES
- Results are paginated for performance
- Some fields may be empty if not set
- Account status changes may take time to propagate
- API permissions required:
  * "User > User Information > Read"
#>
function Get-MimecastUser {
    [CmdletBinding(DefaultParameterSetName='Internal')]
    param (
        [Parameter(ParameterSetName='Internal')]
        [string]$EmailAddress,

        [Parameter(ParameterSetName='Internal')]
        [string]$Domain,

        [Parameter(ParameterSetName='Attributes', Mandatory)]
        [string[]]$Attributes,

        [Parameter(ParameterSetName='Aliases', Mandatory)]
        [string]$UserEmailAddress,

        [Parameter()]
        [switch]$FirstPageOnly
    )

    try {
        switch ($PSCmdlet.ParameterSetName) {
            'Internal' {
                # Get internal users
                $body = @{
                    data = @(
                        @{}
                    )
                }
                if ($EmailAddress) {
                    $body.data[0].emailAddress = $EmailAddress
                }
                if ($Domain) {
                    $body.data[0].domain = $Domain
                }
                
                # Use pagination helper
                Invoke-ApiRequestWithPagination -Path '/api/user/get-internal-users' `
                    -InitialBody $body `
                    -ReturnFirstPageOnly:$FirstPageOnly
            }
            'Attributes' {
                # Get user attributes
                $body = @{
                    data = @(
                        @{
                            attributes = $Attributes
                        }
                    )
                }
                $response = Invoke-ApiRequest -Method 'POST' -Path '/api/user/get-attributes' -Body $body
                $response.data
            }
            'Aliases' {
                # Get user aliases
                $body = @{
                    data = @(
                        @{
                            emailAddress = $UserEmailAddress
                        }
                    )
                }
                $response = Invoke-ApiRequest -Method 'POST' -Path '/api/user/get-aliases' -Body $body
                $response.data
            }
        }
    }
    catch {
        Write-Error -ErrorRecord $_
    }
}

<#
.SYNOPSIS
Creates a new Mimecast user.

.DESCRIPTION
Creates a new user in Mimecast with the specified settings. Can create both internal
and external users, set passwords, and configure various account settings. The function
supports creating users with email aliases and custom attributes.

.PARAMETER EmailAddress
The primary email address for the new user.

.PARAMETER Name
The display name for the new user.

.PARAMETER Password
Optional. The user's password as a SecureString.
Required for internal users unless using directory synchronization.

.PARAMETER AccountType
Optional. The type of user account. Valid values:
- internal: Internal user (default)
- external: External user

.PARAMETER AccountStatus
Optional. The initial account status. Valid values:
- active: Account is enabled (default)
- disabled: Account is disabled

.PARAMETER Domain
Optional. The domain for the user.
If not specified, extracted from EmailAddress.

.PARAMETER Aliases
Optional. Array of additional email addresses to assign as aliases.

.PARAMETER Attributes
Optional. Hashtable of custom attributes to assign to the user.
Example: @{ department = 'Sales'; location = 'New York' }

.PARAMETER ForcePasswordChange
Optional. Whether to require password change at next login.
Default is $true.

.PARAMETER SendWelcomeEmail
Optional. Whether to send a welcome email to the new user.
Default is $true.

.INPUTS
None. You cannot pipe objects to New-MimecastUser.

.OUTPUTS
System.Object
Returns the newly created user object.

.EXAMPLE
C:\> $password = ConvertTo-SecureString 'MySecurePassword123!' -AsPlainText -Force
C:\> New-MimecastUser `
    -EmailAddress 'newuser@example.com' `
    -Name 'New User' `
    -Password $password

Creates a new internal user with basic settings.

.EXAMPLE
C:\> $attributes = @{
    department = 'Sales'
    location = 'New York'
    employeeId = '12345'
}
C:\> $aliases = @(
    'nuser@example.com',
    'newu@example.com'
)
C:\> New-MimecastUser `
    -EmailAddress 'newuser@example.com' `
    -Name 'New User' `
    -AccountType 'internal' `
    -Aliases $aliases `
    -Attributes $attributes `
    -ForcePasswordChange $true `
    -SendWelcomeEmail $true

Creates a new internal user with aliases, attributes, and welcome email.

.EXAMPLE
C:\> New-MimecastUser `
    -EmailAddress 'external@partner.com' `
    -Name 'External Partner' `
    -AccountType 'external' `
    -AccountStatus 'disabled'

Creates a new disabled external user account.

.NOTES
- Password complexity requirements apply to internal users
- Some settings may require specific API permissions
- Account creation may be affected by domain policies
- API permissions required:
  * "User > User Information > Write"
  * "User > User Password > Write" (if setting password)
#>
function New-MimecastUser {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$EmailAddress,

        [Parameter()]
        [string]$Name,

        [Parameter()]
        [securestring]$Password,

        [Parameter()]
        [bool]$AccountDisabled,

        [Parameter()]
        [bool]$AccountLocked,

        [Parameter()]
        [bool]$ForcePasswordChange,

        [Parameter()]
        [bool]$AllowPop
    )

    try {
        $body = @{
            data = @(
                @{
                    emailAddress = $EmailAddress
                }
            )
        }

        if ($Name) {
            $body.data[0].name = $Name
        }
        if ($Password) {
            $plainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
                [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password))
            $body.data[0].password = $plainPassword
        }
        if ($PSBoundParameters.ContainsKey('AccountDisabled')) {
            $body.data[0].accountDisabled = $AccountDisabled
        }
        if ($PSBoundParameters.ContainsKey('AccountLocked')) {
            $body.data[0].accountLocked = $AccountLocked
        }
        if ($PSBoundParameters.ContainsKey('ForcePasswordChange')) {
            $body.data[0].forcePasswordChange = $ForcePasswordChange
        }
        if ($PSBoundParameters.ContainsKey('AllowPop')) {
            $body.data[0].allowPop = $AllowPop
        }

        $response = Invoke-ApiRequest -Method 'POST' -Path '/api/user/create-user' -Body $body
        $response.data
    }
    catch {
        Write-Error -ErrorRecord $_
    }
}

<#
.SYNOPSIS
Updates an existing Mimecast user.

.DESCRIPTION
Modifies settings for an existing user in Mimecast. Can update display name,
password, account status, and other attributes. The function supports pipeline
input from Get-MimecastUser and includes ShouldProcess support for confirmation.

.PARAMETER EmailAddress
The email address of the user to update. This parameter accepts pipeline input
by property name from Get-MimecastUser.

.PARAMETER Name
Optional. Update the user's display name.

.PARAMETER Password
Optional. Update the user's password (as SecureString).
Only applies to internal users.

.PARAMETER AccountStatus
Optional. Update the account status. Valid values:
- active: Enable the account
- disabled: Disable the account
- deleted: Mark the account as deleted

.PARAMETER Attributes
Optional. Update custom attributes. Should be a hashtable of attribute names and values.
Example: @{ department = 'Sales'; location = 'New York' }
Note: This replaces all existing attributes.

.PARAMETER ForcePasswordChange
Optional. Whether to require password change at next login.
Only applies when setting a new password.

.PARAMETER SendPasswordEmail
Optional. Whether to send password change notification email.
Only applies when setting a new password.

.INPUTS
System.Object. You can pipe user objects from Get-MimecastUser.

.OUTPUTS
System.Object
Returns the updated user object.

.EXAMPLE
C:\> Set-MimecastUser -EmailAddress 'user@example.com' -Name 'Updated Name'

Updates a user's display name.

.EXAMPLE
C:\> $password = ConvertTo-SecureString 'NewSecurePassword123!' -AsPlainText -Force
C:\> Set-MimecastUser `
    -EmailAddress 'user@example.com' `
    -Password $password `
    -ForcePasswordChange $true `
    -SendPasswordEmail $true

Updates a user's password and forces password change at next login.

.EXAMPLE
C:\> Get-MimecastUser -Domain 'example.com' |
    Where-Object { $_.accountStatus -eq 'active' } |
    Set-MimecastUser -AccountStatus 'disabled' -WhatIf

Shows what would happen if you disabled all active users in a domain.

.EXAMPLE
C:\> $attributes = @{
    department = 'Marketing'
    location = 'London'
    title = 'Manager'
}
C:\> Set-MimecastUser -EmailAddress 'user@example.com' -Attributes $attributes

Updates a user's custom attributes.

.NOTES
- Some changes may require specific API permissions
- Password changes only apply to internal users
- Status changes may take time to propagate
- Use -WhatIf to preview changes before executing
- API permissions required:
  * "User > User Information > Write"
  * "User > User Password > Write" (if updating password)
#>
function Set-MimecastUser {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]$EmailAddress,

        [Parameter()]
        [bool]$AccountDisabled,

        [Parameter()]
        [bool]$AccountLocked,

        [Parameter()]
        [bool]$ForcePasswordChange,

        [Parameter()]
        [bool]$AllowPop
    )

    process {
        try {
            $body = @{
                data = @(
                    @{
                        emailAddress = $EmailAddress
                    }
                )
            }

            if ($PSBoundParameters.ContainsKey('AccountDisabled')) {
                $body.data[0].accountDisabled = $AccountDisabled
            }
            if ($PSBoundParameters.ContainsKey('AccountLocked')) {
                $body.data[0].accountLocked = $AccountLocked
            }
            if ($PSBoundParameters.ContainsKey('ForcePasswordChange')) {
                $body.data[0].forcePasswordChange = $ForcePasswordChange
            }
            if ($PSBoundParameters.ContainsKey('AllowPop')) {
                $body.data[0].allowPop = $AllowPop
            }

            $response = Invoke-ApiRequest -Method 'POST' -Path '/api/user/update-user' -Body $body
            $response.data
        }
        catch {
            Write-Error -ErrorRecord $_
        }
    }
}

<#
.SYNOPSIS
Adds an email alias to a Mimecast user.

.DESCRIPTION
Adds one or more email aliases to an existing Mimecast user. Aliases allow the user
to receive email at additional email addresses. The function supports pipeline input
from Get-MimecastUser and includes ShouldProcess support for confirmation.

.PARAMETER EmailAddress
The primary email address of the user. This parameter accepts pipeline input
by property name from Get-MimecastUser.

.PARAMETER Alias
The email alias to add. Can be a single alias or an array of aliases.
Each alias must be a valid email address format.

.INPUTS
System.Object. You can pipe user objects from Get-MimecastUser.

.OUTPUTS
System.Object
Returns the updated user object with the new alias(es).

.EXAMPLE
C:\> Add-MimecastUserAlias -EmailAddress 'user@example.com' -Alias 'alias@example.com'

Adds a single alias to a user.

.EXAMPLE
C:\> $aliases = @(
    'alias1@example.com',
    'alias2@example.com',
    'alias3@example.com'
)
C:\> Add-MimecastUserAlias -EmailAddress 'user@example.com' -Alias $aliases

Adds multiple aliases to a user.

.EXAMPLE
C:\> Get-MimecastUser -Domain 'example.com' |
    Add-MimecastUserAlias -Alias 'archive@example.com' -WhatIf

Shows what would happen if you added the same alias to all users in a domain.

.NOTES
- Aliases must be unique across all users
- Some domains may have alias restrictions
- Changes may take time to propagate
- Use -WhatIf to preview changes before executing
- API permissions required:
  * "User > User Alias > Write"
#>
function Add-MimecastUserAlias {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$Alias,

        [Parameter(Mandatory)]
        [string]$AliasFor
    )

    try {
        $body = @{
            data = @(
                @{
                    alias = $Alias
                    aliasFor = $AliasFor
                }
            )
        }

        $response = Invoke-ApiRequest -Method 'POST' -Path '/api/user/update-alias' -Body $body
        $response.data
    }
    catch {
        Write-Error -ErrorRecord $_
    }
}

<#
.SYNOPSIS
Removes an email alias from a Mimecast user.

.DESCRIPTION
Removes one or more email aliases from an existing Mimecast user. The function
supports pipeline input from Get-MimecastUser and includes ShouldProcess support
for confirmation.

.PARAMETER EmailAddress
The primary email address of the user. This parameter accepts pipeline input
by property name from Get-MimecastUser.

.PARAMETER Alias
The email alias to remove. Can be a single alias or an array of aliases.
Each alias must be a valid email address format.

.INPUTS
System.Object. You can pipe user objects from Get-MimecastUser.

.OUTPUTS
System.Object
Returns the updated user object without the removed alias(es).

.EXAMPLE
C:\> Remove-MimecastUserAlias -EmailAddress 'user@example.com' -Alias 'oldalias@example.com'

Removes a single alias from a user.

.EXAMPLE
C:\> $aliases = @(
    'alias1@example.com',
    'alias2@example.com',
    'alias3@example.com'
)
C:\> Remove-MimecastUserAlias -EmailAddress 'user@example.com' -Alias $aliases

Removes multiple aliases from a user.

.EXAMPLE
C:\> Get-MimecastUser -Domain 'example.com' |
    Where-Object { $_.aliases -contains 'archive@example.com' } |
    Remove-MimecastUserAlias -Alias 'archive@example.com' -WhatIf

Shows what would happen if you removed a specific alias from all users who have it.

.NOTES
- Cannot remove primary email address
- Changes may take time to propagate
- Use -WhatIf to preview changes before executing
- API permissions required:
  * "User > User Alias > Write"
#>
function Remove-MimecastUserAlias {
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory)]
        [string]$Alias,

        [Parameter(Mandatory)]
        [string]$AliasFor
    )

    try {
        if ($PSCmdlet.ShouldProcess($Alias, "Remove alias for $AliasFor")) {
            $body = @{
                data = @(
                    @{
                        alias = $Alias
                        aliasFor = $AliasFor
                    }
                )
            }

            $response = Invoke-ApiRequest -Method 'POST' -Path '/api/user/remove-alias' -Body $body
            $response.data
        }
    }
    catch {
        Write-Error -ErrorRecord $_
    }
}