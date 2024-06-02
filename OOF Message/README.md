# PowerShell Out-of-Office Management Script

This PowerShell script allows you to manage out-of-office settings for Exchange mailboxes. It provides functionalities to enable, disable, or schedule out-of-office messages with various options.

## Features

- **Enable:** Set out-of-office messages for a specified mailbox.
- **Disable:** Turn off out-of-office messages for a specified mailbox.
- **Schedule:** Set scheduled out-of-office messages with start and end times.

## Requirements

- PowerShell version 3.0 or later.
- Exchange Management Shell module installed.
- Appropriate permissions to manage Exchange mailboxes.

## Usage

1. Clone or download the script from this repository.
2. Open PowerShell.
3. Navigate to the directory containing the script.
4. Run the script with appropriate parameters:

    ```powershell
    .\Manage-OutOfOffice.ps1 -MailboxName <MailboxEmailAddress> -Mode <Enable|Disable|Schedule> [-Message <OutOfOfficeMessage>] [-StartTime <StartTime>] [-EndTime <EndTime>]
    ```

    - `MailboxName`: Email address of the mailbox to manage.
    - `Mode`: Action to perform (`Enable`, `Disable`, or `Schedule`).
    - `Message`: (Optional) Out-of-office message to set.
    - `StartTime`: (Optional) Start time for scheduled out-of-office.
    - `EndTime`: (Optional) End time for scheduled out-of-office.

5. Follow the prompts and confirmations provided by the script.

## Example

```powershell
.\Manage-OutOfOffice.ps1 -MailboxName user@example.com -Mode Enable -Message "I am currently out of the office."
```
