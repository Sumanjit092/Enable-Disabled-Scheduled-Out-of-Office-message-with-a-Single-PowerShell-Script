# Microsoft 365 Group Creation Script
# Overview
This PowerShell script automates the creation of Microsoft 365 Groups with optional Teams integration. It ensures the necessary modules are installed and connected, handles potential errors, and logs any issues that occur during execution.

# Features
Creates a Microsoft 365 Group with specified attributes.
Optionally creates a corresponding Microsoft Team.
Includes error handling and logging mechanisms.
Checks and installs required Microsoft Exchange Online and Microsoft Teams modules.
Requirements
PowerShell 5.1 or later.
Administrative permissions for Microsoft 365.
Internet connectivity.
Parameters
Display: The display name of the group. (Mandatory)
EmailAddress: The primary email address for the group. (Mandatory)
Owners: The owners of the group. (Mandatory)
Members: The members of the group. (Mandatory)
AccessType: The access type of the group (Public or Private). (Mandatory)
Description: The description of the group. (Mandatory)
AllowEmailExternals: Specifies whether to allow external email addresses (Yes or No). (Mandatory)
AutoSubNewMembers: Specifies whether to auto-subscribe new members to calendar events (Yes or No). (Mandatory)
CreateTeams: Specifies whether to create a Microsoft Team for the group (Yes or No). (Mandatory)

# Usage
Ensure you have the necessary permissions and prerequisites.
Open PowerShell with administrative privileges.
Run the script with the appropriate parameters.

# Example:

.\Create-M365Group.ps1 -Display "Project Team" -EmailAddress "projectteam@contoso.com" -Owners "owner1@contoso.com","owner2@contoso.com" -Members "member1@contoso.com","member2@contoso.com" -AccessType "Private" -Description "Group for Project Team" -AllowEmailExternals "No" -AutoSubNewMembers "Yes" -CreateTeams "Yes"

# Script Breakdown
CheckInternet: Verifies internet connectivity.
CheckMSExchange: Ensures the Microsoft Exchange Online Management module is installed and imported.
CheckMSTeams: Ensures the Microsoft Teams module is installed and imported.
Connect-Modules: Connects to Exchange Online and Microsoft Teams using a service account.
Main Logic:
Checks if the group or email address already exists.
Creates the Microsoft 365 Group.
Adds members and owners to the group.
Removes the service account from the group.
Optionally creates a Microsoft Team for the group.
Error Handling: Logs any errors and attempts to clean up partially created groups.
Cleanup: Removes PowerShell sessions.
Error Logging
Errors encountered during execution are logged to C:\Logs\GroupCreationErrorLog.txt with timestamps for troubleshooting.

# Important Notes
Creating Microsoft 365 Groups and Teams affects your organization's collaboration environment. Ensure you have the necessary permissions and understanding before running this script.
Review the group and team settings after creation to ensure they align with your organization's policies and requirements.
Use this script at your own risk. It comes with no warranties, whether express or implied.
