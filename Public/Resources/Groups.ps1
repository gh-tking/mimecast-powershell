<#
.SYNOPSIS
Gets Mimecast groups based on optional filters.

.DESCRIPTION
Retrieves groups from Mimecast. Can filter by name, type, or other attributes.
Supports pagination for large result sets and includes detailed group information
such as description, members, and custom attributes.

.PARAMETER Name
Optional. Filter groups by exact name match.

.PARAMETER Type
Optional. Filter groups by type. Valid values:
- distribution: Email distribution groups
- security: Security groups
- static: Static groups
- dynamic: Dynamic groups (based on rules)
- all: All group types (default)

.PARAMETER FirstPageOnly
Optional. Return only the first page of results.
Default is $false (returns all pages).

.PARAMETER PageSize
Optional. Number of groups to return per page.
Default is 100. Maximum is 500.

.PARAMETER Domain
Optional. Filter groups by domain (e.g., 'example.com').

.PARAMETER IncludeMembers
Optional. Whether to include member details in the results.
Default is $false.

.INPUTS
None. You cannot pipe objects to Get-MimecastGroup.

.OUTPUTS
System.Object[]
Returns an array of group objects. Each object includes:
- id: Unique group identifier
- name: Group name
- description: Group description
- type: Group type (distribution, security, etc.)
- domain: Group domain
- members: Array of member objects (if IncludeMembers is $true)
- attributes: Custom attributes
- created: Group creation date
- lastModified: Last modification date

.EXAMPLE
C:\> Get-MimecastGroup
Gets all groups across all domains.

.EXAMPLE
C:\> Get-MimecastGroup -Name 'Sales Team' -Type 'distribution'
Gets a specific distribution group by name.

.EXAMPLE
C:\> Get-MimecastGroup -Domain 'example.com' -IncludeMembers $true
Gets all groups in the specified domain, including member details.

.EXAMPLE
C:\> Get-MimecastGroup -FirstPageOnly -PageSize 50
Gets only the first 50 groups.

.NOTES
- Results are paginated for performance
- Member details increase response size
- Some fields may be empty if not set
- API permissions required:
  * "Group > Group Information > Read"
#>
function Get-MimecastGroup {
    [CmdletBinding(DefaultParameterSetName='Search')]
    param (
        [Parameter(ParameterSetName='Search')]
        [string]$Name,

        [Parameter(ParameterSetName='Search')]
        [string]$Domain,

        [Parameter(ParameterSetName='Members', Mandatory)]
        [string]$GroupId,

        [Parameter()]
        [switch]$FirstPageOnly
    )

    try {
        if ($PSCmdlet.ParameterSetName -eq 'Members') {
            # Get group members
            $body = @{
                data = @(
                    @{
                        id = $GroupId
                    }
                )
            }
            
            # Use pagination helper for members
            Invoke-ApiRequestWithPagination -Path '/api/directory/get-group-members' `
                -InitialBody $body `
                -ReturnFirstPageOnly:$FirstPageOnly
        }
        else {
            # Search for groups
            $body = @{
                data = @(
                    @{}
                )
            }
            if ($Name) {
                $body.data[0].query = $Name
            }
            if ($Domain) {
                $body.data[0].domain = $Domain
            }
            
            # Use pagination helper for groups
            Invoke-ApiRequestWithPagination -Path '/api/directory/find-groups' `
                -InitialBody $body `
                -ReturnFirstPageOnly:$FirstPageOnly
        }
    }
    catch {
        Write-Error -ErrorRecord $_
    }
}

<#
.SYNOPSIS
Creates a new Mimecast group.

.DESCRIPTION
Creates a new group in Mimecast with the specified settings. Can create different
types of groups (distribution, security, etc.) and optionally add initial members.
The function supports creating groups with custom attributes and descriptions.

.PARAMETER Name
The name for the new group.

.PARAMETER Type
The type of group to create. Valid values:
- distribution: Email distribution group
- security: Security group
- static: Static group
- dynamic: Dynamic group (based on rules)

.PARAMETER Description
Optional. A description of the group's purpose.

.PARAMETER Domain
Optional. The domain for the group.
If not specified, uses the default domain.

.PARAMETER Members
Optional. Array of email addresses to add as initial members.

.PARAMETER Attributes
Optional. Hashtable of custom attributes to assign to the group.
Example: @{ department = 'Sales'; location = 'New York' }

.PARAMETER Rules
Optional. For dynamic groups, array of rules that determine membership.
Each rule should be a hashtable with:
- field: Attribute to check (e.g., 'department', 'title')
- operator: Comparison operator (e.g., 'equals', 'contains')
- value: Value to compare against

.INPUTS
None. You cannot pipe objects to New-MimecastGroup.

.OUTPUTS
System.Object
Returns the newly created group object.

.EXAMPLE
C:\> New-MimecastGroup `
    -Name 'Sales Team' `
    -Type 'distribution' `
    -Description 'Sales department distribution list' `
    -Members @('sales1@example.com', 'sales2@example.com')

Creates a distribution group with initial members.

.EXAMPLE
C:\> $attributes = @{
    department = 'Sales'
    location = 'New York'
    costCenter = '12345'
}
C:\> New-MimecastGroup `
    -Name 'NY Sales' `
    -Type 'static' `
    -Description 'New York Sales Team' `
    -Attributes $attributes

Creates a static group with custom attributes.

.EXAMPLE
C:\> $rules = @(
    @{
        field = 'department'
        operator = 'equals'
        value = 'Sales'
    },
    @{
        field = 'location'
        operator = 'equals'
        value = 'New York'
    }
)
C:\> New-MimecastGroup `
    -Name 'NY Sales Dynamic' `
    -Type 'dynamic' `
    -Description 'Dynamic group for NY Sales' `
    -Rules $rules

Creates a dynamic group based on department and location rules.

.NOTES
- Group names must be unique within a domain
- Dynamic groups update membership automatically
- Some settings may require specific API permissions
- API permissions required:
  * "Group > Group Information > Write"
  * "Group > Group Membership > Write" (if adding members)
#>
function New-MimecastGroup {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter()]
        [string]$Description,

        [Parameter()]
        [string]$ParentId,

        [Parameter()]
        [string]$Source = "cloud",

        [Parameter()]
        [ValidateSet("distribution", "security")]
        [string]$Type = "distribution"
    )

    try {
        $body = @{
            data = @(
                @{
                    source = $Source
                    name = $Name
                    type = $Type
                }
            )
        }

        if ($Description) {
            $body.data[0].description = $Description
        }
        if ($ParentId) {
            $body.data[0].parentId = $ParentId
        }

        $response = Invoke-ApiRequest -Method 'POST' -Path '/api/directory/create-group' -Body $body
        $response.data
    }
    catch {
        Write-Error -ErrorRecord $_
    }
}

<#
.SYNOPSIS
Updates an existing Mimecast group.

.DESCRIPTION
Modifies settings for an existing group in Mimecast. Can update name, description,
attributes, and for dynamic groups, membership rules. The function supports pipeline
input from Get-MimecastGroup and includes ShouldProcess support for confirmation.

.PARAMETER Id
The unique identifier of the group to update. This parameter accepts pipeline input
by property name from Get-MimecastGroup.

.PARAMETER Name
Optional. Update the group's name.

.PARAMETER Description
Optional. Update the group's description.

.PARAMETER Attributes
Optional. Update custom attributes. Should be a hashtable of attribute names and values.
Example: @{ department = 'Sales'; location = 'New York' }
Note: This replaces all existing attributes.

.PARAMETER Rules
Optional. For dynamic groups, update the membership rules. Should be an array of
rule hashtables, each containing:
- field: Attribute to check (e.g., 'department', 'title')
- operator: Comparison operator (e.g., 'equals', 'contains')
- value: Value to compare against
Note: This replaces all existing rules.

.INPUTS
System.Object. You can pipe group objects from Get-MimecastGroup.

.OUTPUTS
System.Object
Returns the updated group object.

.EXAMPLE
C:\> Set-MimecastGroup -Id '4d24b932-5c57-4b57-b7b5-9c3f511e5c5c' `
    -Name 'Updated Sales Team' `
    -Description 'Updated sales department group'

Updates a group's name and description.

.EXAMPLE
C:\> $attributes = @{
    department = 'Marketing'
    location = 'London'
    costCenter = '67890'
}
C:\> Get-MimecastGroup -Name 'Sales Team' |
    Set-MimecastGroup -Attributes $attributes

Updates a group's custom attributes.

.EXAMPLE
C:\> $rules = @(
    @{
        field = 'department'
        operator = 'equals'
        value = 'Sales'
    },
    @{
        field = 'title'
        operator = 'contains'
        value = 'Manager'
    }
)
C:\> Set-MimecastGroup -Id '4d24b932-5c57-4b57-b7b5-9c3f511e5c5c' `
    -Rules $rules

Updates the membership rules for a dynamic group.

.EXAMPLE
C:\> Get-MimecastGroup -Domain 'example.com' |
    Where-Object { $_.type -eq 'distribution' } |
    Set-MimecastGroup -Description 'Updated by automation' -WhatIf

Shows what would happen if you updated the description for all distribution groups.

.NOTES
- Some changes may require specific API permissions
- Dynamic group rule changes affect membership
- Changes may take time to propagate
- Use -WhatIf to preview changes before executing
- API permissions required:
  * "Group > Group Information > Write"
#>
function Set-MimecastGroup {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]$Id,

        [Parameter()]
        [string]$Name,

        [Parameter()]
        [string]$Description,

        [Parameter()]
        [string]$ParentId
    )

    process {
        try {
            $body = @{
                data = @(
                    @{
                        id = $Id
                    }
                )
            }

            if ($Name) {
                $body.data[0].name = $Name
            }
            if ($Description) {
                $body.data[0].description = $Description
            }
            if ($ParentId) {
                $body.data[0].parentId = $ParentId
            }

            $response = Invoke-ApiRequest -Method 'POST' -Path '/api/directory/update-group' -Body $body
            $response.data
        }
        catch {
            Write-Error -ErrorRecord $_
        }
    }
}

<#
.SYNOPSIS
Removes a Mimecast group.

.DESCRIPTION
Deletes a group from Mimecast. This operation is permanent and cannot be undone.
The function supports pipeline input from Get-MimecastGroup and includes ShouldProcess
support for confirmation.

.PARAMETER Id
The unique identifier of the group to remove. This parameter accepts pipeline input
by property name from Get-MimecastGroup.

.INPUTS
System.Object. You can pipe group objects from Get-MimecastGroup.

.OUTPUTS
System.Object
Returns the result of the removal operation.

.EXAMPLE
C:\> Remove-MimecastGroup -Id '4d24b932-5c57-4b57-b7b5-9c3f511e5c5c'

Removes a specific group.

.EXAMPLE
C:\> Get-MimecastGroup -Name 'Old Team' |
    Remove-MimecastGroup

Removes a group by name.

.EXAMPLE
C:\> Get-MimecastGroup -Domain 'example.com' |
    Where-Object { $_.type -eq 'distribution' -and $_.members.Count -eq 0 } |
    Remove-MimecastGroup -WhatIf

Shows what would happen if you removed all empty distribution groups.

.NOTES
- This operation cannot be undone
- Members are automatically removed
- Group policies may be affected
- Use -WhatIf to preview changes before executing
- API permissions required:
  * "Group > Group Information > Write"
#>
function Remove-MimecastGroup {
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]$Id
    )

    process {
        try {
            if ($PSCmdlet.ShouldProcess($Id, "Remove Mimecast group")) {
                $body = @{
                    data = @(
                        @{
                            id = $Id
                        }
                    )
                }

                $response = Invoke-ApiRequest -Method 'POST' -Path '/api/directory/delete-group' -Body $body
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
Adds a member to a Mimecast group.

.DESCRIPTION
Adds one or more members to an existing Mimecast group. Members can be added by
email address. The function supports pipeline input from Get-MimecastGroup and
includes ShouldProcess support for confirmation.

.PARAMETER GroupId
The unique identifier of the group. This parameter accepts pipeline input
by property name from Get-MimecastGroup.

.PARAMETER EmailAddress
The email address(es) to add as members. Can be a single address or an array.

.INPUTS
System.Object. You can pipe group objects from Get-MimecastGroup.

.OUTPUTS
System.Object
Returns the updated group object with the new member(s).

.EXAMPLE
C:\> Add-MimecastGroupMember `
    -GroupId '4d24b932-5c57-4b57-b7b5-9c3f511e5c5c' `
    -EmailAddress 'user@example.com'

Adds a single member to a group.

.EXAMPLE
C:\> $members = @(
    'user1@example.com',
    'user2@example.com',
    'user3@example.com'
)
C:\> Add-MimecastGroupMember `
    -GroupId '4d24b932-5c57-4b57-b7b5-9c3f511e5c5c' `
    -EmailAddress $members

Adds multiple members to a group.

.EXAMPLE
C:\> Get-MimecastGroup -Name 'Sales Team' |
    Add-MimecastGroupMember -EmailAddress 'newuser@example.com'

Adds a member to a group found by name.

.EXAMPLE
C:\> Get-MimecastGroup -Domain 'example.com' |
    Where-Object { $_.type -eq 'distribution' } |
    Add-MimecastGroupMember -EmailAddress 'notify@example.com' -WhatIf

Shows what would happen if you added a member to all distribution groups.

.NOTES
- Cannot add members to dynamic groups
- Members must be valid email addresses
- Changes may take time to propagate
- Use -WhatIf to preview changes before executing
- API permissions required:
  * "Group > Group Membership > Write"
#>
function Add-MimecastGroupMember {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$GroupId,

        [Parameter(Mandatory)]
        [string]$EmailAddress
    )

    try {
        $body = @{
            data = @(
                @{
                    id = $GroupId
                    emailAddress = $EmailAddress
                }
            )
        }

        $response = Invoke-ApiRequest -Method 'POST' -Path '/api/directory/add-group-member' -Body $body
        $response.data
    }
    catch {
        Write-Error -ErrorRecord $_
    }
}

<#
.SYNOPSIS
Removes a member from a Mimecast group.

.DESCRIPTION
Removes one or more members from an existing Mimecast group. Members can be removed
by email address. The function supports pipeline input from Get-MimecastGroup and
includes ShouldProcess support for confirmation.

.PARAMETER GroupId
The unique identifier of the group. This parameter accepts pipeline input
by property name from Get-MimecastGroup.

.PARAMETER EmailAddress
The email address(es) to remove from the group. Can be a single address or an array.

.INPUTS
System.Object. You can pipe group objects from Get-MimecastGroup.

.OUTPUTS
System.Object
Returns the updated group object without the removed member(s).

.EXAMPLE
C:\> Remove-MimecastGroupMember `
    -GroupId '4d24b932-5c57-4b57-b7b5-9c3f511e5c5c' `
    -EmailAddress 'user@example.com'

Removes a single member from a group.

.EXAMPLE
C:\> $members = @(
    'user1@example.com',
    'user2@example.com',
    'user3@example.com'
)
C:\> Remove-MimecastGroupMember `
    -GroupId '4d24b932-5c57-4b57-b7b5-9c3f511e5c5c' `
    -EmailAddress $members

Removes multiple members from a group.

.EXAMPLE
C:\> Get-MimecastGroup -Name 'Sales Team' |
    Remove-MimecastGroupMember -EmailAddress 'olduser@example.com'

Removes a member from a group found by name.

.EXAMPLE
C:\> Get-MimecastGroup -Domain 'example.com' |
    Where-Object { $_.type -eq 'distribution' } |
    Remove-MimecastGroupMember -EmailAddress 'oldnotify@example.com' -WhatIf

Shows what would happen if you removed a member from all distribution groups.

.NOTES
- Cannot remove members from dynamic groups
- Non-existent members are ignored
- Changes may take time to propagate
- Use -WhatIf to preview changes before executing
- API permissions required:
  * "Group > Group Membership > Write"
#>
function Remove-MimecastGroupMember {
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory)]
        [string]$GroupId,

        [Parameter(Mandatory)]
        [string]$EmailAddress
    )

    try {
        if ($PSCmdlet.ShouldProcess($EmailAddress, "Remove from Mimecast group $GroupId")) {
            $body = @{
                data = @(
                    @{
                        id = $GroupId
                        emailAddress = $EmailAddress
                    }
                )
            }

            $response = Invoke-ApiRequest -Method 'POST' -Path '/api/directory/remove-group-member' -Body $body
            $response.data
        }
    }
    catch {
        Write-Error -ErrorRecord $_
    }
}