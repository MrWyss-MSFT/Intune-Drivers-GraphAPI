# Intune Drivers – Graph API

PowerShell tooling to inspect Intune **Windows driver update policies**
(`windowsDriverUpdateProfile`) via Microsoft Graph, using only the
`Microsoft.Graph.Authentication` module.

> [!WARNING]
> **Disclaimer — this was "vibe coded".** This project was generated quickly with the help of
> an AI assistant and is provided **as-is**, with no warranty of any kind. Treat it as an
> **example / source of inspiration**, not as production-ready software. **Do not run it in any
> environment without first reading and understanding the code yourself**, validating it against
> your own requirements, and testing it in a safe/non-production tenant. You are responsible for
> any actions taken in your environment.

## Check-DriverPolicyAssignment.ps1

Reports whether one or more Windows driver update policies are assigned to a **group** or a
**device**.

Driver update policies are assigned to Entra groups (or the built-in *All devices* / *All users*
targets), not directly to devices. When you pass a device, the script resolves the device's
**transitive group memberships** and evaluates every policy assignment against those groups,
honoring **exclusion** groups as well.

### Prerequisites

- PowerShell 5.1+ or PowerShell 7+
- The `Microsoft.Graph.Authentication` module:

  ```powershell
  Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
  ```

- An account with permission to consent to / use these **delegated** scopes:
  - `DeviceManagementConfiguration.Read.All`
  - `Group.Read.All`
  - `Device.Read.All`
  - `GroupMember.Read.All`
- An active Intune license in the tenant.

The script signs in interactively with `Connect-MgGraph` and requests the scopes above. If you
are already connected with the required scopes, the existing session is reused.

### Usage

Check a group by display name:

```powershell
.\Check-DriverPolicyAssignment.ps1 -Group "All Pilot Devices"
```

Check a group by Object ID:

```powershell
.\Check-DriverPolicyAssignment.ps1 -Group 11111111-2222-3333-4444-555555555555
```

Check a device by name:

```powershell
.\Check-DriverPolicyAssignment.ps1 -Device "DESKTOP-1234"
```

Return result objects to the pipeline (in addition to the on-screen table):

```powershell
$results = .\Check-DriverPolicyAssignment.ps1 -Device "DESKTOP-1234" -PassThru
$results | Where-Object Effective
```

Use a specific tenant (the script uses the Graph beta endpoint by default):

```powershell
.\Check-DriverPolicyAssignment.ps1 -Group "Drivers - Ring 1" -TenantId contoso.onmicrosoft.com
```

### Parameters

| Parameter      | Description                                                                 |
| -------------- | --------------------------------------------------------------------------- |
| `-Group`       | Group display name or Object ID (GUID) to check.                            |
| `-Device`      | Device name, Intune managed device ID, or Entra device object ID to check.  |
| `-UseV1Endpoint` | Use the Microsoft Graph `v1.0` endpoint. Default is `beta`, since `windowsDriverUpdateProfiles` is currently beta-only. |
| `-TenantId`    | Optional tenant ID/domain passed to `Connect-MgGraph`.                      |
| `-PassThru`    | Emit result objects to the pipeline in addition to the summary table.       |

### Output

For each matching policy the script reports:

| Field         | Meaning                                                                          |
| ------------- | -------------------------------------------------------------------------------- |
| `PolicyName`  | Driver update policy display name.                                               |
| `Status`      | `Assigned`, `Excluded`, `Excluded (overrides include)`, or `Not assigned`.       |
| `Effective`   | `True` when the policy is included **and** not excluded for the target.          |
| `IncludedVia` | Group(s) or *All devices/users* target that caused the include.                  |
| `ExcludedVia` | Group(s) that caused the exclusion.                                              |

An exclusion overrides an include, so a policy can match yet not effectively apply.

### Performance in large environments

- **Paging** is handled automatically for every Graph collection (follows `@odata.nextLink`),
  so results are complete regardless of tenant size.
- **Policies + assignments** are retrieved in a single `$expand=assignments` request rather than
  one request per policy (with an automatic fallback if a tenant rejects the expand).
- **Throttling (HTTP 429)** is retried automatically with `Retry-After` by the underlying
  `Invoke-MgGraphRequest` (Microsoft Graph SDK).
- **Device checks** read the device's transitive group memberships in one paged call
  (`$top=999`). A device in thousands of groups simply means a few extra pages.
- Typical runtime is dominated by interactive sign-in; the data calls are a handful of requests
  (number of driver policies is usually small — tens, not thousands).

### Required Graph permissions reference

See Microsoft's guidance on driver and firmware update programmatic controls:
<https://learn.microsoft.com/en-us/windows/deployment/windows-autopatch/manage/windows-autopatch-driver-and-firmware-update-programmatic-controls>
