# Microsoft 365 Group Creation Script

## Overview

This PowerShell script automates the creation of Microsoft 365 Groups with optional Teams integration. It ensures the necessary modules are installed and connected, handles potential errors, and logs any issues that occur during execution.

## Features

- Creates a Microsoft 365 Group with specified attributes.
- Optionally creates a corresponding Microsoft Team.
- Includes error handling and logging mechanisms.
- Checks and installs required Microsoft Exchange Online and Microsoft Teams modules.

## Requirements

- PowerShell 5.1 or later.
- Administrative permissions for Microsoft 365.
- Internet connectivity.

## Parameters

- **Display**: The display name of the group. (Mandatory)
- **EmailAddress**: The primary email address for the group. (Mandatory)
- **Owners**: The owners of the group. (Mandatory)
- **Members**: The members of the group. (Mandatory)
- **AccessType**: The access type of the group (Public or Private). (Mandatory)
- **Description**: The description of the group. (Mandatory)
- **AllowEmailExternals**: Specifies whether to allow external email addresses (Yes or No). (Mandatory)
- **AutoSubNewMembers**: Specifies whether to auto-subscribe new members to calendar events (Yes or No). (Mandatory)
- **CreateTeams**: Specifies whether to create a Microsoft Team for the group (Yes or No). (Mandatory)

## Usage

1. Ensure you have the necessary permissions and prerequisites.
2. Open PowerShell with administrative privileges.
3. Run the script with the appropriate parameters.

### Example:

```powershell
.\Create-M365Group.ps1 -Display "Project Team" -EmailAddress "projectteam@contoso.com" -Owners "owner1@contoso.com","owner2@contoso.com" -Members "member1@contoso.com","member2@contoso.com" -AccessType "Private" -Description "Group for Project Team" -AllowEmailExternals "No" -AutoSubNewMembers "Yes" -CreateTeams "Yes"
```
