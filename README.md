# Enable-Disabled-Scheduled-Out-of-Office-message-with-a-Single-PowerShell-Script
Before run this PowerShell, please check if ExchangeOnline Module is installed or not.
To verify run the command, Get-Module Get-Module ExchangeOnlineManagement
If Exchange Online Module is not Installed on your machine, you can install it running command Install-Module -Name ExchangeOnlineManagement
After Install Exchange Online Module, import the module running command Import-Module ExchangeOnlineManagement
If Exchange Online Module is already present or if you have installed Exchange Online Module. Connect to Exchange Online Module.
To connect with Exchange Online Module run Connect-ExchangeOnline
After connecting to Exchange Online Module, run this PowerShell Script.
Script will ask you Enter Mailbox Name. Please enter Mailbox Email ID. 
Next it will ask you to enter Out of Office message. If you want to Disable it, please keep it blank.
Next, script will ask you to choose whether you want Enable, Disable or Schedule Out of Office Message.
Enter your selection. 
For Enable you can type Enable or E, Disable you can type Disable or D and Schedule you can type Schedule or S
If you choose Schedule option, script will ask you enter Enter Out-of-Office Message Start Date & Time also End Date & Time in 'MM/DD/YYYY HH:MM:SS' Format.
Microsoft uses Coordinated Universal Time (UTC) format. Enter date and time accordingly.
How to convert UTC time to local time: https://support.microsoft.com/en-us/topic/how-to-convert-utc-time-to-local-time-0569c45d-5fb8-a516-814c-75374b44830a
Within few seconds Script will display message based on your selection.
If you like my script or if you have any advise to upgrade this script, please feel free to leave any comments.
