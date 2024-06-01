# AssignMailboxPermissions.ps1

## Description

`AssignMailboxPermissions.ps1` is a PowerShell script designed to assign specified permissions (FullAccess, SendAs, SendOnBehalf) to users or security groups for a shared mailbox in an Exchange environment. The script validates the existence of the provided email addresses and the shared mailbox before assigning the permissions, ensuring that only valid entities are processed.

## Prerequisites

- Exchange Online Management Module or Exchange Server Management Tools installed.
- Appropriate permissions to run the script and modify mailbox permissions in Exchange.
- PowerShell 5.1 or later.

## Parameters

### `AccessEmail` (Mandatory)

The email addresses of the users or security groups to be granted permissions. This parameter accepts an array of strings.

- Type: `string[]`
- Example: `user1@domain.com, user2@domain.com`

### `SharedMailboxEmail` (Mandatory)

The email address of the shared mailbox to which permissions are being granted.

- Type: `string`
- Example: `shared@domain.com`

### `AccessType` (Mandatory)

The type of permissions to be granted. Valid values are 'SendAs', 'SendOnBehalf', and 'FullAccess'. This parameter accepts an array of strings.

- Type: `string[]`
- Example: `FullAccess, SendAs`

## Usage

### Example 1

Grant FullAccess and SendAs permissions to `user1@domain.com` and `user2@domain.com` for the shared mailbox `shared@domain.com`.

```powershell
.\AssignMailboxPermissions.ps1 -AccessEmail user1@domain.com, user2@domain.com -SharedMailboxEmail shared@domain.com -AccessType FullAccess, SendAs
```
### Example 2
Grant SendOnBehalf permissions to group@domain.com for the shared mailbox shared@domain.com

```powershell
.\AssignMailboxPermissions.ps1 -AccessEmail group@domain.com -SharedMailboxEmail shared@domain.com -AccessType SendOnBehalf
```
