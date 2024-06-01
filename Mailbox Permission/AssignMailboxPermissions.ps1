<#
.SYNOPSIS
    Assigns specified permissions to users or security groups for a shared mailbox in Exchange.

.DESCRIPTION
    This script validates and assigns permissions such as FullAccess, SendAs, and SendOnBehalf
    to users or security groups for a specified shared mailbox. It ensures that the provided email
    addresses exist as either user mailboxes or security groups and that the shared mailbox exists.
    It handles errors gracefully and provides feedback on the actions performed.

.PARAMETER AccessEmail
    The email addresses of the users or security groups to be granted permissions.
    This parameter is mandatory.

.PARAMETER SharedMailboxEmail
    The email address of the shared mailbox to which permissions are being granted.
    This parameter is mandatory.

.PARAMETER AccessType
    The type of permissions to be granted. Valid values are 'SendAs', 'SendOnBehalf', and 'FullAccess'.
    This parameter is mandatory and can accept multiple values.

.EXAMPLE
    .\AssignMailboxPermissions.ps1 -AccessEmail user1@domain.com,user2@domain.com -SharedMailboxEmail shared@domain.com -AccessType FullAccess,SendAs

    This example grants FullAccess and SendAs permissions to user1@domain.com and user2@domain.com
    for the shared mailbox shared@domain.com.

.EXAMPLE
    .\AssignMailboxPermissions.ps1 -AccessEmail group@domain.com -SharedMailboxEmail shared@domain.com -AccessType SendOnBehalf

    This example grants SendOnBehalf permissions to group@domain.com for the shared mailbox shared@domain.com.

.AUTHOR
    Sumanjit Pan

.VERSION
    1.0
    1.1 (Parameter Added: SendOnBehalf Permission)
    1.2 (Access Email validation between UserMailbox and MailEnabled Security Group)


.FIRST PUBLISH DATE
    15th Feb, 2021

#>

# Function to check internet connectivity

param (
    [Parameter(Mandatory=$true)]
    [string[]]$AccessEmail,

    [Parameter(Mandatory=$true)]
    [string]$SharedMailboxEmail,

    [Parameter(Mandatory=$true)]
    [ValidateSet("FullAccess", "SendAs", "SendOnBehalf")]
    [string[]]$AccessType
)
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

function Connect-Modules {

    try {
        # Connect to Exchange Online
        Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
        Connect-ExchangeOnline -AppId "App_id" -CertificateThumbprint "Thumbprint string of certificate" -Organization "contoso.onmicrosoft.com"
        Write-Host "Connected to Exchange Online successfully." -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to connect with Exchange Online." -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        exit
    }
}

cls

'===================================================================================================='
Write-Host '                                  Mailbox Permission Management                                               ' -ForegroundColor Green
'===================================================================================================='

''                    
Write-Host "                                          IMPORTANT NOTES                                           " -ForegroundColor Red 
Write-Host "===================================================================================================="
Write-Host "This PowerShell script is provided as freeware and is offered on an 'as is' basis without any warranties," -ForegroundColor Yellow 
Write-Host "whether express or implied, including but not limited to warranties of merchantability, fitness for a particular" -ForegroundColor Yellow 
Write-Host "purpose, or non-infringement. The entire risk arising out of the use or performance of the script remains with" -ForegroundColor Yellow 
Write-Host "the user." -ForegroundColor yellow 
''
Write-Host "The script manages mailbox permissions in an Exchange environment. It validates the existence of provided email" -ForegroundColor Yellow 
Write-Host "addresses and shared mailboxes before granting specified permissions such as FullAccess, SendAs, and SendOnBehalf." -ForegroundColor Yellow
''
Write-Host "For more information, please refer to the documentation and links below:" -ForegroundColor yellow 
Write-Host "Exchange Online PowerShell V2: https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2" -ForegroundColor yellow
Write-Host "Exchange Server PowerShell Documentation: https://docs.microsoft.com/en-us/powershell/exchange/exchange-server-powershell" -ForegroundColor Yellow
Write-Host "Managing Mailbox Permissions in Exchange: https://docs.microsoft.com/en-us/exchange/mailbox-permissions" -ForegroundColor Yellow
Write-Host "===================================================================================================="
''
CheckInternet

CheckMSExchange

Connect-Modules

try {
    # Validate AccessEmail addresses
    foreach ($Email in $AccessEmail) {
        $SecurityGroup = Get-DistributionGroup $Email -ErrorAction SilentlyContinue | Where-Object {$_.RecipientTypeDetails -eq "MailUniversalSecurityGroup"}
        if (-not $SecurityGroup) {
            $UserMailbox = Get-Mailbox $Email -ErrorAction SilentlyContinue | Where-Object {$_.RecipientTypeDetails -eq "UserMailbox"}
        }

        if (-not $SecurityGroup -and -not $UserMailbox) {
            Write-Host "Neither Security Group nor User Mailbox exists for email: $Email." -ForegroundColor Red
            return
        }
    }

    # Validate SharedMailboxEmail
    $SharedMailbox = Get-Mailbox -Identity $SharedMailboxEmail -ErrorAction Stop | Where-Object {$_.RecipientTypeDetails -eq "SharedMailbox"}
    if (-not $SharedMailbox) {
        Write-Host "Shared Mailbox does not exist for email: $SharedMailboxEmail" -ForegroundColor Red
        return
    }

    # Assign permissions based on AccessType
    foreach ($Email in $AccessEmail) {
        foreach ($Access in $AccessType) {
            switch ($Access) {
                "FullAccess" {
                    $ExistingFullAccess = Get-MailboxPermission -Identity $SharedMailboxEmail -User $Email -ErrorAction SilentlyContinue | Where-Object { $_.AccessRights -contains "FullAccess" }
                    if ($ExistingFullAccess) {
                        Write-Host "$Email already has FullAccess permission to the Shared Mailbox" -ForegroundColor Yellow
                    } else {
                        Add-MailboxPermission -Identity $SharedMailboxEmail -User $Email -AccessRights FullAccess -AutoMapping:$true -InheritanceType All
                        Write-Host "$Email has been granted Full Access to the Shared Mailbox" -ForegroundColor Green
                    }
                }
                "SendAs" {
                    $ExistingSendAs = Get-RecipientPermission -Identity $SharedMailboxEmail -Trustee $Email -ErrorAction SilentlyContinue | Where-Object { $_.AccessRights -contains "SendAs" }
                    if ($ExistingSendAs) {
                        Write-Host "$Email already has SendAs permission to the Shared Mailbox" -ForegroundColor Yellow
                    } else {
                        Add-RecipientPermission -Identity $SharedMailboxEmail -Trustee $Email -AccessRights SendAs -Confirm:$false
                        Write-Host "$Email has been granted SendAs access to the Shared Mailbox" -ForegroundColor Green
                    }
                }
                "SendOnBehalf" {
                    $ExistingSendOnBehalf = (Get-Mailbox -Identity $SharedMailboxEmail).GrantSendOnBehalfTo -contains ($Email -split "@")[0]
                    if ($ExistingSendOnBehalf) {
                        Write-Host "$Email already has SendOnBehalf permission to the Shared Mailbox" -ForegroundColor Yellow
                    } else {
                        Set-Mailbox $SharedMailboxEmail -GrantSendOnBehalfTo @{Add=$Email}
                        Write-Host "$Email has been granted SendOnBehalf access to the Shared Mailbox" -ForegroundColor Green
                    }
                }
            }
        }
    }
}
catch {
    Write-Host "Error occurred: $_" -ForegroundColor Red
    return
}