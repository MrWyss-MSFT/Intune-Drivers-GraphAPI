#Requires -Modules Microsoft.Graph.Authentication
<#
.SYNOPSIS
    Checks whether an Intune Windows driver update policy is assigned to a device or a group.

.DESCRIPTION
    Reports whether one or more Intune Windows driver update policies
    (windowsDriverUpdateProfile) are assigned to a specified Entra group or device.

    Driver update policies are assigned to groups (or the built-in "All devices" /
    "All users" targets), not directly to devices. When a device is supplied, the script
    resolves the device's transitive group memberships and evaluates every policy
    assignment against those groups, honoring exclusion groups as well.

    Only the Microsoft.Graph.Authentication module is required; all Graph calls are made
    with Invoke-MgGraphRequest.

.PARAMETER Group
    Display name or Object ID (GUID) of the Entra group to check.

.PARAMETER Device
    Device name or ID (Intune managed device ID or Entra device object ID) to check.

.PARAMETER UseV1Endpoint
    Use the Microsoft Graph v1.0 endpoint. By default the script uses beta, because the
    windowsDriverUpdateProfiles resource is currently only available on the beta endpoint.

.PARAMETER TenantId
    Optional tenant ID to pass to Connect-MgGraph.

.PARAMETER PassThru
    Emit result objects to the pipeline in addition to the formatted summary table.

.EXAMPLE
    .\Check-DriverPolicyAssignment.ps1 -Group "All Pilot Devices"

.EXAMPLE
    .\Check-DriverPolicyAssignment.ps1 -Device "DESKTOP-1234"

.EXAMPLE
    .\Check-DriverPolicyAssignment.ps1 -Group 11111111-2222-3333-4444-555555555555 -PassThru

.NOTES
    Required delegated scopes:
        DeviceManagementConfiguration.Read.All
        Group.Read.All
        Device.Read.All
        GroupMember.Read.All
#>
[CmdletBinding(DefaultParameterSetName = 'Group')]
param(
    [Parameter(Mandatory, ParameterSetName = 'Group', Position = 0)]
    [string] $Group,

    [Parameter(Mandatory, ParameterSetName = 'Device', Position = 0)]
    [string] $Device,

    [Parameter()]
    [switch] $UseV1Endpoint,

    [Parameter()]
    [string] $TenantId,

    [Parameter()]
    [switch] $PassThru
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Helpers ----------------------------------------------------------------

# windowsDriverUpdateProfiles is a beta-only Graph resource, so beta is the default.
$script:GraphVersion = if ($UseV1Endpoint) { 'v1.0' } else { 'beta' }
$script:RequiredScopes = @(
    'DeviceManagementConfiguration.Read.All'
    'Group.Read.All'
    'Device.Read.All'
    'GroupMember.Read.All'
)

function Connect-Graph {
    [CmdletBinding()]
    param()

    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        throw "The 'Microsoft.Graph.Authentication' module is required. Install it with: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser"
    }
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

    $context = $null
    try { $context = Get-MgContext } catch { $context = $null }

    $hasScopes = $false
    if ($context -and $context.Scopes) {
        $missing = $script:RequiredScopes | Where-Object { $_ -notin $context.Scopes }
        $hasScopes = -not $missing
    }

    if (-not $context -or -not $hasScopes) {
        Write-Verbose 'Connecting to Microsoft Graph...'
        $connectParams = @{ Scopes = $script:RequiredScopes; NoWelcome = $true }
        if ($TenantId) { $connectParams['TenantId'] = $TenantId }
        Connect-MgGraph @connectParams | Out-Null
        $context = Get-MgContext
    }

    if (-not $context) {
        throw 'Failed to establish a Microsoft Graph connection.'
    }
    Write-Verbose ("Connected as '{0}' (tenant {1})." -f $context.Account, $context.TenantId)
}

function Invoke-GraphGet {
    <#
        Performs a GET against Graph and returns all pages of 'value' (or the single
        object when the response is not a collection). Accepts an absolute URL or a
        path relative to the configured Graph version root.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string] $Uri
    )

    if ($Uri -notmatch '^https?://') {
        $Uri = "https://graph.microsoft.com/$script:GraphVersion/$($Uri.TrimStart('/'))"
    }

    $results = [System.Collections.Generic.List[object]]::new()
    $next = $Uri
    do {
        $response = Invoke-MgGraphRequest -Method GET -Uri $next -OutputType PSObject
        if ($null -eq $response) { break }

        if ($response.PSObject.Properties.Name -contains 'value') {
            foreach ($item in $response.value) { $results.Add($item) }
            if ($response.PSObject.Properties.Name -contains '@odata.nextLink') {
                $next = $response.'@odata.nextLink'
            }
            else {
                $next = $null
            }
        }
        else {
            $results.Add($response)
            $next = $null
        }
    } while ($next)

    # Comma operator prevents PowerShell from unrolling the list (which would drop
    # the .Count property for single-item results under Set-StrictMode).
    return , $results
}

function ConvertTo-ODataFilterString {
    param([Parameter(Mandatory)][string] $Value)
    return $Value.Replace("'", "''")
}

function Test-IsGuid {
    param([string] $Value)
    [System.Guid]::TryParse($Value, [ref]([System.Guid]::Empty))
}

#endregion Helpers -------------------------------------------------------------

#region Resolution -------------------------------------------------------------

function Resolve-TargetGroup {
    <# Resolves the -Group input to an object with Id and DisplayName. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string] $InputValue)

    if (Test-IsGuid $InputValue) {
        $g = Invoke-GraphGet -Uri "groups/$InputValue?`$select=id,displayName" | Select-Object -First 1
        if (-not $g) { throw "No group found with Object ID '$InputValue'." }
        return [pscustomobject]@{ Id = $g.id; DisplayName = $g.displayName }
    }

    $filter = ConvertTo-ODataFilterString $InputValue
    $found = Invoke-GraphGet -Uri "groups?`$filter=displayName eq '$filter'&`$select=id,displayName"
    if (-not $found -or $found.Count -eq 0) {
        throw "No group found with display name '$InputValue'."
    }
    if ($found.Count -gt 1) {
        $ids = ($found | ForEach-Object { $_.id }) -join ', '
        throw "Multiple groups match display name '$InputValue' (ids: $ids). Re-run using the Object ID."
    }
    return [pscustomobject]@{ Id = $found[0].id; DisplayName = $found[0].displayName }
}

function Resolve-TargetDevice {
    <#
        Resolves the -Device input to an object with managed device info and the Entra
        device object ID (used to read transitive group memberships).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string] $InputValue)

    $managed = $null
    if (Test-IsGuid $InputValue) {
        # Could be a managed device id or an Entra device object id.
        try {
            $managed = Invoke-GraphGet -Uri "deviceManagement/managedDevices/$InputValue`?`$select=id,deviceName,azureADDeviceId" | Select-Object -First 1
        } catch { $managed = $null }
    }

    if (-not $managed) {
        $filter = ConvertTo-ODataFilterString $InputValue
        $found = Invoke-GraphGet -Uri "deviceManagement/managedDevices?`$filter=deviceName eq '$filter'&`$select=id,deviceName,azureADDeviceId"
        if ($found.Count -gt 1) {
            $names = ($found | ForEach-Object { "$($_.deviceName) [$($_.id)]" }) -join '; '
            throw "Multiple managed devices match '$InputValue': $names. Re-run using the managed device ID."
        }
        if ($found.Count -eq 1) { $managed = $found[0] }
    }

    # Determine the Entra device object id from the azureADDeviceId (the device's deviceId).
    $entraObjectId = $null
    $deviceName = $InputValue
    $azureAdDeviceId = $null

    if ($managed) {
        $deviceName = $managed.deviceName
        $azureAdDeviceId = $managed.azureADDeviceId
    }
    elseif (Test-IsGuid $InputValue) {
        # Treat the GUID as a possible Entra deviceId or object id directly.
        $azureAdDeviceId = $InputValue
    }

    if ($azureAdDeviceId -and (Test-IsGuid $azureAdDeviceId) -and $azureAdDeviceId -ne '00000000-0000-0000-0000-000000000000') {
        $filter = ConvertTo-ODataFilterString $azureAdDeviceId
        $dirDevice = Invoke-GraphGet -Uri "devices?`$filter=deviceId eq '$filter'&`$select=id,displayName,deviceId" | Select-Object -First 1
        if ($dirDevice) {
            $entraObjectId = $dirDevice.id
            if (-not $managed) { $deviceName = $dirDevice.displayName }
        }
    }

    if (-not $entraObjectId -and (Test-IsGuid $InputValue)) {
        # Last resort: the supplied GUID might itself be the Entra device object id.
        try {
            $dirDevice = Invoke-GraphGet -Uri "devices/$InputValue`?`$select=id,displayName,deviceId" | Select-Object -First 1
            if ($dirDevice) {
                $entraObjectId = $dirDevice.id
                if (-not $managed) { $deviceName = $dirDevice.displayName }
            }
        } catch { }
    }

    if (-not $entraObjectId) {
        throw "Could not resolve device '$InputValue' to an Entra device object. Verify the device name/ID and that you have Device.Read.All."
    }

    return [pscustomobject]@{
        DeviceName    = $deviceName
        ManagedId     = if ($managed) { $managed.id } else { $null }
        EntraObjectId = $entraObjectId
    }
}

function Get-DeviceGroupMembership {
    <# Returns hashtable of group object id -> display name for a device's transitive groups. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string] $EntraObjectId)

    $groups = @{}
    $members = Invoke-GraphGet -Uri "devices/$EntraObjectId/transitiveMemberOf?`$select=id,displayName&`$top=999"
    foreach ($m in $members) {
        $hasType = $m.PSObject.Properties.Name -contains '@odata.type'
        $isGroup = $hasType -and ($m.'@odata.type' -eq '#microsoft.graph.group')
        if (($isGroup -or ($m.PSObject.Properties.Name -contains 'displayName')) -and $m.id) {
            $groups[$m.id] = $m.displayName
        }
    }
    return $groups
}

#endregion Resolution ----------------------------------------------------------

#region Policies & Evaluation --------------------------------------------------

function Get-DriverUpdatePolicy {
    <#
        Returns all driver update profiles with their assignments.
        Uses $expand=assignments to retrieve everything in a single paged request
        (avoids an N+1 call per profile). Falls back to per-profile assignment
        requests if the tenant rejects the expand.
    #>
    [CmdletBinding()]
    param()

    try {
        $profiles = Invoke-GraphGet -Uri "deviceManagement/windowsDriverUpdateProfiles?`$select=id,displayName&`$expand=assignments&`$top=100"
        $needsFallback = $false
        foreach ($p in $profiles) {
            if ($p.PSObject.Properties.Name -notcontains 'assignments') {
                $needsFallback = $true
                break
            }
        }
        if (-not $needsFallback) {
            foreach ($p in $profiles) {
                if ($null -eq $p.assignments) {
                    Add-Member -InputObject $p -NotePropertyName 'assignments' -NotePropertyValue @() -Force
                }
            }
            return , $profiles
        }
    }
    catch {
        Write-Verbose "Expand on assignments failed ($($_.Exception.Message)); falling back to per-profile requests."
    }

    $profiles = Invoke-GraphGet -Uri "deviceManagement/windowsDriverUpdateProfiles?`$select=id,displayName&`$top=100"
    foreach ($p in $profiles) {
        $assignments = Invoke-GraphGet -Uri "deviceManagement/windowsDriverUpdateProfiles/$($p.id)/assignments"
        Add-Member -InputObject $p -NotePropertyName 'assignments' -NotePropertyValue $assignments -Force
    }
    return , $profiles
}

function Get-AssignmentTargetInfo {
    <# Normalizes a single assignment target into type / groupId. #>
    param([Parameter(Mandatory)] $Assignment)

    $target = $Assignment.target
    $type = $target.'@odata.type'
    $groupId = if ($target.PSObject.Properties.Name -contains 'groupId') { $target.groupId } else { $null }

    [pscustomobject]@{
        Type    = $type
        GroupId = $groupId
        IsAll   = ($type -eq '#microsoft.graph.allDevicesAssignmentTarget' -or $type -eq '#microsoft.graph.allLicensedUsersAssignmentTarget')
        IsExcl  = ($type -eq '#microsoft.graph.exclusionGroupAssignmentTarget')
        IsIncl  = ($type -eq '#microsoft.graph.groupAssignmentTarget')
    }
}

function Resolve-PolicyAssignment {
    <#
        Evaluates one policy against a set of group ids (the device's groups, or the
        single target group). Returns a result object describing the effective verdict.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $Policy,
        [Parameter(Mandatory)][hashtable] $GroupIdMap   # id -> displayName
    )

    $includedVia = [System.Collections.Generic.List[string]]::new()
    $excludedVia = [System.Collections.Generic.List[string]]::new()
    $allTarget = $false

    foreach ($a in $Policy.assignments) {
        $t = Get-AssignmentTargetInfo -Assignment $a
        if ($t.IsAll) {
            $allTarget = $true
            $includedVia.Add('All devices/users')
        }
        elseif ($t.IsIncl -and $t.GroupId -and $GroupIdMap.ContainsKey($t.GroupId)) {
            $includedVia.Add($GroupIdMap[$t.GroupId])
        }
        elseif ($t.IsExcl -and $t.GroupId -and $GroupIdMap.ContainsKey($t.GroupId)) {
            $excludedVia.Add($GroupIdMap[$t.GroupId])
        }
    }

    $isIncluded = ($includedVia.Count -gt 0)
    $isExcluded = ($excludedVia.Count -gt 0)

    $status =
        if ($isExcluded -and $isIncluded) { 'Excluded (overrides include)' }
        elseif ($isExcluded) { 'Excluded' }
        elseif ($isIncluded) { 'Assigned' }
        else { 'Not assigned' }

    # Effective = applies only when included and not excluded.
    $effective = $isIncluded -and -not $isExcluded

    [pscustomobject]@{
        PolicyName  = $Policy.displayName
        PolicyId    = $Policy.id
        Status      = $status
        Effective   = $effective
        IncludedVia = ($includedVia | Select-Object -Unique) -join ', '
        ExcludedVia = ($excludedVia | Select-Object -Unique) -join ', '
        AllTarget   = $allTarget
    }
}

#endregion Policies & Evaluation -----------------------------------------------

#region Main -------------------------------------------------------------------

Connect-Graph

Write-Verbose 'Loading driver update policies...'
$policies = Get-DriverUpdatePolicy
if (-not $policies -or $policies.Count -eq 0) {
    Write-Warning 'No Windows driver update policies were found in this tenant.'
    return
}

# Build the group-id map to evaluate against, and a friendly target description.
$groupIdMap = @{}
$targetDescription = ''

if ($PSCmdlet.ParameterSetName -eq 'Group') {
    $g = Resolve-TargetGroup -InputValue $Group
    $groupIdMap[$g.Id] = $g.DisplayName
    $targetDescription = "Group '$($g.DisplayName)' ($($g.Id))"
}
else {
    $d = Resolve-TargetDevice -InputValue $Device
    $targetDescription = "Device '$($d.DeviceName)' (Entra object $($d.EntraObjectId))"
    Write-Verbose 'Resolving device group memberships...'
    $groupIdMap = Get-DeviceGroupMembership -EntraObjectId $d.EntraObjectId
    Write-Verbose ("Device is a member of {0} group(s)." -f $groupIdMap.Count)
}

$results = foreach ($p in $policies) {
    Resolve-PolicyAssignment -Policy $p -GroupIdMap $groupIdMap
}

$applicable = $results | Where-Object { $_.Status -ne 'Not assigned' }

Write-Host ''
Write-Host "Driver update policy assignment for: $targetDescription" -ForegroundColor Cyan
Write-Host ''

if (-not $applicable) {
    Write-Host 'No driver update policy is assigned.' -ForegroundColor Yellow
}
else {
    $applicable |
        Sort-Object @{ Expression = 'Effective'; Descending = $true }, PolicyName |
        Format-Table PolicyName, Status, Effective, IncludedVia, ExcludedVia -AutoSize |
        Out-Host

    $effectiveCount = ($applicable | Where-Object { $_.Effective }).Count
    if ($effectiveCount -gt 0) {
        Write-Host ("Result: {0} driver update policy/policies effectively apply." -f $effectiveCount) -ForegroundColor Green
    }
    else {
        Write-Host 'Result: matching policies exist but all are excluded; none effectively apply.' -ForegroundColor Yellow
    }
}

if ($PassThru) {
    $results
}

#endregion Main ----------------------------------------------------------------
