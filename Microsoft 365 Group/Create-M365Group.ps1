<#
.SYNOPSIS
    Automates the creation of Microsoft 365 Groups with optional Teams integration.

.DESCRIPTION
    This script creates a Microsoft 365 Group with specified attributes. It can also create a corresponding Microsoft Team if requested.
    The script includes error handling and logging mechanisms to capture any issues during execution.

.PARAMETER DisplayName
    The display name of the group.

.PARAMETER EmailAddress
    The primary email address for the group.

.PARAMETER Owners
    The owners of the group.

.PARAMETER Members
    The members of the group.

.PARAMETER AccessType
    The access type of the group (Public or Private).

.PARAMETER Description
    The description of the group.

.PARAMETER AllowEmailExternals
    Specifies whether to allow external email addresses (Yes or No).

.PARAMETER AutoSubNewMembers
    Specifies whether to auto-subscribe new members to calendar events (Yes or No).

.PARAMETER CreateTeams
    Specifies whether to create a Microsoft Team for the group (Yes or No).

.AUTHOR
    Sumanjit Pan

.VERSION
    1.0
    1.1 (Install and Import Module Function)
    1.2 (Optional Teams Integration, Validation Set in Parameter, & Cleanup Partially Created Group)

.FIRST PUBLISH DATE
    24th Jan, 2021

#>

Param(
    [Parameter(Mandatory=$true)]
    [string]$DisplayName,
    [Parameter(Mandatory=$true)]
    [string]$EmailAddress,
    [Parameter(Mandatory=$true)]
    [string[]]$Owners,
    [Parameter(Mandatory=$true)]
    [string[]]$Members,
    [Parameter(Mandatory=$true)]
    [ValidateSet('Public', 'Private')]
    [string]$AccessType,
    [Parameter(Mandatory=$true)]
    [string]$Description,
    [Parameter(Mandatory=$true)]
    [ValidateSet('Yes', 'No')]
    [string]$AllowEmailExternals,
    [Parameter(Mandatory=$true)]
    [ValidateSet('Yes', 'No')]
    [string]$AutoSubNewMembers,
    [Parameter(Mandatory=$true)]
    [ValidateSet('Yes', 'No')]
    [string]$CreateTeams
)

# Function to check internet connectivity
Function CheckInternet {
    Write-Host "Checking internet connectivity..." -ForegroundColor Yellow
    $statuscode = (Invoke-WebRequest -Uri https://adminwebservice.microsoftonline.com/ProvisioningService.svc -UseBasicParsing).StatusCode
    if ($statuscode -ne 200) {
        Write-Host "Operation aborted. Unable to connect to Microsoft Graph, please check your internet connection." -ForegroundColor Red
        exit
    }
    Write-Host "Internet connectivity check passed." -ForegroundColor Green
}

# Function to ensure Microsoft Exchange and Teams module is installed and imported
Function CheckMSExchange {
    Write-Host "Checking Microsoft Exchange Online Module..." -ForegroundColor Yellow
    if (Get-Module -ListAvailable | Where-Object {$_.Name -like "ExchangeOnlineManagement"}) {
        Write-Host "Microsoft ExchangeOnline Module is installed." -ForegroundColor Green
        Import-Module -Name "ExchangeOnlineManagement"
        Write-Host "Microsoft Exchange Online Module has been imported." -ForegroundColor Cyan
    } else {
        Write-Host "Microsoft Exchange Online is not installed." -ForegroundColor Red
        Write-Host "Installing Microsoft Exchange Online Module..." -ForegroundColor Yellow
        Install-Module -Name "ExchangeOnlineManagement" -Force -ErrorAction Stop
        Write-Host "Microsoft Exchange Online Module is installed." -ForegroundColor Green
        Import-Module -Name "ExchangeOnlineManagement"
        Write-Host "Microsoft Exchange Online Module has been imported." -ForegroundColor Cyan
    }
}

Function CheckMSTeams {
    Write-Host "Checking Microsoft Teams Module..." -ForegroundColor Yellow
    if (Get-Module -ListAvailable | Where-Object {$_.Name -like "MicrosoftTeams"}) {
        Write-Host "Microsoft Teams Module is installed." -ForegroundColor Green
        Import-Module -Name "MicrosoftTeams"
        Write-Host "Microsoft Teams Module has been imported." -ForegroundColor Cyan
    } else {
        Write-Host "Microsoft Teams is not installed." -ForegroundColor Red
        Write-Host "Installing Microsoft Teams Module..." -ForegroundColor Yellow
        Install-Module -Name "MicrosoftTeams" -Force -ErrorAction Stop
        Write-Host "Microsoft Teams Module is installed." -ForegroundColor Green
        Import-Module -Name "MicrosoftTeams"
        Write-Host "Microsoft Teams Module has been imported." -ForegroundColor Cyan
    }
}
#App based authentication not supported for Group Creation
$ServiceAccount = 'serviceaccount@contoso.com'

function Connect-Modules {

    try {
        # Connect to Exchange Online
        Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
        Connect-ExchangeOnline -UserPrincipalName $ServiceAccount -ShowBanner:$false -ErrorAction Stop
        Write-Host "Connected to Exchange Online successfully." -ForegroundColor Green

        # Connect to Microsoft Teams
        Write-Host "Connecting to Microsoft Teams..." -ForegroundColor Yellow
        Connect-MicrosoftTeams -ErrorAction Stop
        Write-Host "Connected to Microsoft Teams successfully." -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to connect to Exchange Online or Microsoft Teams." -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        exit
    }
}

Cls

'===================================================================================================='
Write-Host '                        Microsoft 365 Group Creation with Optional Teams Integration              ' -ForegroundColor Green 
'===================================================================================================='
''                    
Write-Host "                                          IMPORTANT NOTES                                           " -ForegroundColor red 
Write-Host "===================================================================================================="
Write-Host "This script automates the creation of Microsoft 365 Groups with optional Teams integration." -ForegroundColor Yellow 
Write-Host "It comes with no warranties, whether express or implied. Use it at your own risk." -ForegroundColor Yellow 
''
Write-Host "Creating Microsoft 365 Groups and Teams should be done with caution, as they affect your organization's collaboration environment." -ForegroundColor Yellow 
Write-Host "Ensure you have the necessary permissions and understanding of the implications before running this script." -ForegroundColor Yellow 
''
Write-Host "Teams integration can enhance collaboration, but it also increases complexity. Only enable it if necessary." -ForegroundColor Yellow 
Write-Host "Consider the management overhead and user training required for Teams usage." -ForegroundColor Yellow 
''
Write-Host "After creation, review the group and team settings to ensure they align with your organization's policies and requirements." -ForegroundColor Yellow 
''
Write-Host "===================================================================================================="
''

CheckInternet

CheckMSExchange

CheckMSTeams

Connect-Modules

# Convert parameters to boolean values
$AllowEmailExternal = if ($AllowEmailExternals -eq 'Yes') { $false } else { $true }
$AutoSubNewMember = if ($AutoSubNewMembers -eq 'Yes') { $true } else { $false }
$CreateTeam = if ($CreateTeams -eq 'Yes') { $true } else { $false }

# Function to log errors
function Log-Error {
    param(
        [string]$ErrorMessage
    )
    $ErrorLogPath = "C:\Logs\GroupCreationErrorLog.txt"
    if (-not (Test-Path -Path (Split-Path $ErrorLogPath))) {
        New-Item -Path (Split-Path $ErrorLogPath) -ItemType Directory -Force
    }
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $ErrorEntry = "$Timestamp - $ErrorMessage"
    Add-Content -Path $ErrorLogPath -Value $ErrorEntry
}

try {
    Write-Host "Inside Try block..." -ForegroundColor Green

    Write-Host "Checking if the group or email address already exists..." -ForegroundColor Yellow
    $Group = Get-UnifiedGroup -Identity $DisplayName -ErrorAction SilentlyContinue
    $Email = Get-UnifiedGroup -Identity $EmailAddress -ErrorAction SilentlyContinue

    if (-not $Group -and -not $Email) {
        Write-Host "Group and email address do not exist. Proceeding with creation..." -ForegroundColor Green
        Write-Host "Creating Microsoft 365 Group..." -ForegroundColor Yellow
        #Microsoft Limitation as Multiple Owner can't be added, adding owners as members.
        New-UnifiedGroup -DisplayName $DisplayName -AccessType $AccessType -AlwaysSubscribeMembersToCalendarEvents:$true `
            -AutoSubscribeNewMembers:$AutoSubNewMember -Confirm:$false -PrimarySMTPAddress $EmailAddress `
            -ExoErrorAsWarning -Language "en-US" -Members $Owners -Notes $Description `
            -RequireSenderAuthenticationEnabled:$AllowEmailExternal
        Write-Host "Microsoft 365 Group created successfully." -ForegroundColor Green

        Write-Host "Adding members and owners to the group..." -ForegroundColor Yellow
        Add-UnifiedGroupLinks -Identity $EmailAddress -LinkType Members -Links $Members
        #Promoting as Owner, as owners were added as member during Group creation.
        Add-UnifiedGroupLinks -Identity $EmailAddress -LinkType Owner -Links $Owners
        Write-Host "Members and owners added to the group." -ForegroundColor Green

        Write-Host "Waiting for changes to propagate..." -ForegroundColor Yellow
        Start-Sleep -Seconds 60

        Write-Host "Removing service account from group..." -ForegroundColor Yellow
        #Microsoft Limitation as Service Account Became Group Owner & Member
        Remove-UnifiedGroupLinks -Identity $EmailAddress -LinkType Owner -Links $ServiceAccount -Confirm:$false
        Remove-UnifiedGroupLinks -Identity $EmailAddress -LinkType Members -Links $ServiceAccount -Confirm:$false
        Write-Host "Service account removed from group." -ForegroundColor Green

        if ($CreateTeam) {
            Write-Host "Creating Microsoft Team for the group..." -ForegroundColor Yellow
            $GroupId = (Get-UnifiedGroup $EmailAddress).ExternalDirectoryObjectId
            New-Team -GroupId $GroupId
            Write-Host "Microsoft 365 Group and Teams Group have been created successfully." -ForegroundColor Green
        } else {
            Write-Host "Microsoft 365 Group has been successfully created without a Teams Group." -ForegroundColor Green
        }
    }
    elseif ($Group -and $Email) {
        Write-Host "Microsoft 365 Group and Email Address already exist. Kindly contact your Administrator for further assistance." -ForegroundColor Yellow
        exit
    }
    elseif (-not $Group -and $Email) {
        Write-Host "Email Address already exists. Kindly raise a new request with a unique Email Address." -ForegroundColor Yellow
        exit
    }
    elseif ($Group -and -not $Email) {
        Write-Host "Microsoft 365 Group already exists. Kindly raise a new request with a unique Group Display Name." -ForegroundColor Yellow
        exit
    }
}
catch {
    $ErrorMessage = $_.Exception.Message
    Write-Host "An error occurred: $ErrorMessage" -ForegroundColor Red
    Log-Error -ErrorMessage $ErrorMessage
    if ($Group) {
        Write-Host "Attempting to clean up partially created group..." -ForegroundColor Yellow
        try {
            Remove-UnifiedGroup -Identity $EmailAddress -Confirm:$false
            Write-Host "Partially created group has been removed due to an error." -ForegroundColor Yellow
        } catch {
            Write-Host "Failed to remove the partially created group." -ForegroundColor Red
            Log-Error -ErrorMessage "Failed to remove the partially created group: $($_.Exception.Message)"
        }
    }
}
finally {
    Write-Host "Cleaning up sessions..." -ForegroundColor Yellow
    Get-ConnectionInformation | Disconnect-ExchangeOnline -Confirm:$false
    Disconnect-MicrosoftTeams -Confirm:$false
    Write-Host "Session cleanup completed." -ForegroundColor Green
}