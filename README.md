# M365-TenantOps

PowerShell tools for Microsoft 365 administration across independent tenant sessions, published by **N7 Technologies**.

> **Release status:** Community release. This is an independent project, not an official Microsoft tool. Test in a non-production environment and verify the intended tenant before administrative changes.

## Available tool

### Ad-hoc Microsoft 365 connection

`AdHoc/Connect-M365Adhoc.ps1` starts a tenant-specific PowerShell 7 session for occasional administration. It connects to:

- Microsoft Graph
- Exchange Online
- Microsoft Teams

It intentionally does **not** connect to Azure subscriptions, SharePoint Online, or Purview. Use separate tools for those services.

The script uses **device-code authentication** for all three services. Graph authentication uses the Microsoft identity platform OAuth device-code endpoints directly and supplies an access token to the Microsoft Graph PowerShell SDK. This approach was developed to work around an observed SDK `DeviceCodeCredential` failure. The script keeps Graph access and refresh tokens in its PowerShell process memory and refreshes the Graph token when needed.

After successful verification, the PowerShell prompt resembles:

```text
[YOUR TENANT] [GRAPH:OK] [EXO:OK] [TEAMS:VERIFIED]
PS C:\Work>
```

The Teams indicator reflects the most recently verified session, not a live remote check at every prompt. Prompt indicators are informational; run an explicit assertion before sensitive operations.

## Requirements

- Windows with **PowerShell 7** and a fresh terminal per client/tenant.
- PowerShell modules: `Microsoft.Graph.Authentication`, `Microsoft.Graph.Identity.DirectoryManagement`, `ExchangeOnlineManagement`, and `MicrosoftTeams`.
- The current script explicitly imports Graph Authentication and DirectoryManagement **version 2.41.2**. Other installed versions have not been validated for this release.
- An administrative account with appropriate roles and consent in the target tenant.
- Access to the Microsoft device sign-in page.

Check installed modules:

```powershell
Get-Module -ListAvailable Microsoft.Graph.Authentication,Microsoft.Graph.Identity.DirectoryManagement,ExchangeOnlineManagement,MicrosoftTeams | Select-Object Name,Version
```

## Connect

From the repository root, open a **new PowerShell 7** terminal and run:

```powershell
& ".\AdHoc\Connect-M365Adhoc.ps1"
```

Complete each device-code sign-in as prompted. Use a separate terminal for each tenant; do not reuse a connected session for a different client.

The script automatically identifies the Graph tenant and account and then verifies the Exchange Online and Teams tenant IDs against the original Graph tenant. If any service fails, review the reported status before taking action.

## Session safety and repair

Run these commands in **the same terminal** used to connect:

| Command | Purpose |
| --- | --- |
| `Assert-AdhocClientContext` | Validate Graph, Exchange, and Teams against the initial tenant. |
| `Assert-AdhocClientContext -Service Graph` | Validate Graph only. |
| `Assert-AdhocClientContext -Service Exchange` | Validate Exchange only. |
| `Assert-AdhocClientContext -Service Teams` | Validate Teams only. |
| `Repair-AdhocClientSession -Service All` | Check and repair all three connections. |
| `Repair-AdhocClientSession -Service Graph -Force` | Force Graph token renewal and verify tenant identity. |
| `Repair-AdhocClientSession -Service Exchange` | Restore Exchange connection if necessary. |
| `Repair-AdhocClientSession -Service Teams` | Restore Teams connection if necessary. |

**Important:** Assertions do not automatically prevent unrelated administrator commands from running. The administrator remains responsible for verifying the intended tenant before changes.

## Microsoft Graph permissions

The current production baseline requests these delegated scopes at startup:

| Scope | Use |
| --- | --- |
| `Organization.Read.All` | Discover the organization's display name and initial domain. |
| `Reports.Read.All` | Read supported reporting data. |
| `User.Read.All` | Read directory user data. |
| `Group.Read.All` | Read group data. |
| `offline_access` | Request a refresh token for session renewal. |

Some permissions require administrator consent. Authentication succeeds only when the target tenant permits the application and requested scopes. The scopes are broader than the minimum needed merely to sign in; review them before use in environments with strict least-privilege policies.

To request more scopes for the *current tenant session*:

```powershell
Add-AdhocGraphScopes -Scopes "Application.Read.All"
```

The function requests the combined scope list through device authentication and verifies the tenant/account before accepting the new context. Added permissions apply only when the target tenant grants consent.

## Security considerations

- **Do not** commit client configuration files, tenant exports, credentials, access tokens, refresh tokens, device codes, or authentication transcripts.
- Tokens stay in process memory in this version. Avoid dumping `$global:AdhocM365`, logging script variables, or exporting session state. Close the terminal when finished.
- Each service authenticates independently. A valid Graph context does not guarantee a valid Exchange or Teams session.
- Use a separate PowerShell 7 terminal per tenant and run `Assert-AdhocClientContext` before sensitive administrative commands.
- **Microsoft Graph application identity:** The Graph device-code flow uses the public Microsoft Graph Command Line Tools client ID (`14d82eec-204b-4c2f-b7e8-296a70dab67e`). This identifier is public, not a client secret. Microsoft identifies it as a Microsoft-owned application ID, and the toolkit is **not affiliated with or endorsed by Microsoft**. This script exchanges device codes directly with Microsoft identity endpoints rather than relying on the Graph SDK's built-in device credential handler. Microsoft has not explicitly guaranteed compatibility or support for this independent implementation; authentication behavior and tenant consent policies may change. The script never grants access to a tenant merely because someone knows this client ID.
- The script is provided without any client-specific information. Do not publish real tenant profiles or production logs.

## Project layout

```text
M365-TenantOps/
  AdHoc/
    Connect-M365Adhoc.ps1
  MultiTenant/          # Planned separately
  Docs/
  Tests/
  .github/workflows/
  .gitignore
  README.md
```

The named VS Code multitenant profiles and additional modules are planned for future releases, after sanitization and testing.

## License and support

Licensed under the **MIT License**. See [LICENSE](LICENSE). Copyright (c) 2026 N7 Technologies.

This is an independent N7 Technologies project, not an official Microsoft product. Microsoft 365, Microsoft Graph, Exchange Online, PowerShell, and Teams are Microsoft products and trademarks.
