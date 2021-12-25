# Enable-Disabled-Scheduled-Out-of-Office-message-with-a-Single-PowerShell-Script #

This script will provide you the capability to enable or schedule or disable Auto-reply message or Out of Office message.

Run Powershell as Elevated User.
To run the PowerShell window with elevated permissions just click Start then type PowerShell then Right-Click on PowerShell icon and select Run as Administrator.
As soon as you will run the script, it will check for ExchangeOnline Module. If module is not installed it will Install module.
Next, it will prompt you to enter your credentials.

After successful login, Script will ask you Enter Mailbox Name. Please enter Mailbox Email ID. 
Next it will ask you to enter Out of Office message. If you want to Disable it, please keep it blank.
Next, script will ask you to choose whether you want Enable, Disable or Schedule Out of Office Message.
Enter your selection. 
For Enable you can type Enable or E, Disable you can type Disable or D and Schedule you can type Schedule or S
If you choose Schedule option, script will ask you enter Enter Out-of-Office Message Start Date & Time also End Date & Time in 'MM/DD/YYYY HH:MM:SS' Format.
Microsoft uses Coordinated Universal Time (UTC) format. Enter date and time accordingly.
How to convert UTC time to local time: https://support.microsoft.com/en-us/topic/how-to-convert-utc-time-to-local-time-0569c45d-5fb8-a516-814c-75374b44830a
Within few seconds Script will display message based on your selection.

If you have any question or suggestion or issue with this script please feel free to leave comments.
