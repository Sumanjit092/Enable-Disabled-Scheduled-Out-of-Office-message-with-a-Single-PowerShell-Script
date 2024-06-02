<#
.SYNOPSIS
    Configures Out-of-Office (OOF) or auto-reply messages for Exchange mailboxes.

.DESCRIPTION
    This PowerShell script automates the setup of Out-of-Office (OOF) or auto-reply messages for Exchange mailboxes. 
    It allows users to specify the target mailbox, the content of the out-of-office message, and the action mode (enable, disable, or schedule).
    The script validates the email address format, formats the message content, and executes the chosen action accordingly.
    Results, including the auto-reply state, scheduling details (if applicable), and message content, are displayed for verification.

.PARAMETER MBName
    Specifies the email address of the mailbox for which the out-of-office message will be configured.

.PARAMETER Message
    Specifies the content of the out-of-office message.

.PARAMETER Mode
    Specifies the action mode: 
        - "Enable": Enables the out-of-office message.
        - "Disable": Disables the out-of-office message.
        - "Schedule": Schedules the out-of-office message for a specified time period.

.EXAMPLE
    .\Set-MailboxAutoReply.ps1 -MailboxName user@example.com -Mode Enable -Message "I am currently out of the office."

    This example enables the out-of-office message for the specified mailbox with the provided message.

.AUTHOR
    Sumanjit Pan

.VERSION
    1.0
    1.1 (Install, Import Module and Connect-Modules Function)
    1.2 (Parameter with Validation Set, Date and Time format Validation and Improved Display Result)

.FIRST PUBLISH DATE
    15th April, 2021
#>

param (
    [Parameter(Mandatory=$true)]
    [string]$MailboxName,

    [Parameter(Mandatory=$false)]
    [string]$Message,

    [Parameter(Mandatory=$true)]
    [ValidateSet("Enable", "Disable", "Schedule")]
    [string]$Mode,

    [Parameter(Mandatory=$false)]
    [string]$StartTime,

    [Parameter(Mandatory=$false)]
    [string]$EndTime
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

Write-Host "===================================================================================================="
Write-Host "                              Out-of-Office (OOF) Message Configuration                                  " -ForegroundColor Green
Write-Host "===================================================================================================="
Write-Host "                                          IMPORTANT NOTES                                           " -ForegroundColor Red 
Write-Host "===================================================================================================="
Write-Host "This script automates the configuration of Out-of-Office (OOF) or auto-reply messages for Exchange mailboxes." -ForegroundColor Yellow
Write-Host "It comes with no warranties, whether express or implied. Use it at your own risk." -ForegroundColor Yellow

Write-Host "Managing Out-of-Office messages affects email communication within your organization. Exercise caution and ensure proper testing before deploying changes." -ForegroundColor Yellow

Write-Host "Consider the impact on workflow and communication when enabling, disabling, or scheduling Out-of-Office messages." -ForegroundColor Yellow

Write-Host "Review the configured Out-of-Office message content to ensure it aligns with organizational policies and provides clear communication to senders." -ForegroundColor Yellow

Write-Host "===================================================================================================="

CheckInternet
CheckMSExchange
Connect-Modules

# Function to validate date and time format
function Validate-DateTimeFormat {
    param (
        [string]$DateTime
    )

    if ($DateTime -notmatch '^\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2}$') {
        Write-Host "Please provide Start and End Date & Time in 'MM/DD/YYYY HH:MM:SS' Format!" -ForegroundColor Red
        exit
    }
}

# Validate the email address format
if ($MailboxName -notmatch "^[\w\.\-]+@[\w\-]+\.[\w]{2,4}$") {
    Write-Host "Please provide a valid Email Address!" -ForegroundColor Red
    exit
}

# If no message is provided and mode is not Disable, prompt for a message
if ($Mode -ne "Disable" -and [string]::IsNullOrEmpty($Message)) {
    Write-Host "Please provide an Out-of-Office Message!" -ForegroundColor Red
    exit
}

# Place Out-of-Office Message Inside HTML Tags to preserve Formatting
if ($Message) {
    $OOOtxt = "<pre>$Message</pre>"
} else {
    $OOOtxt = ""
}

# Enable mode
if ($Mode -eq "Enable") {
    Write-Host "Enabling Out-of-Office for $MailboxName..." -ForegroundColor Yellow
    Set-MailboxAutoReplyConfiguration -Identity $MailboxName -AutoReplyState Enabled -InternalMessage $OOOtxt -ExternalMessage $OOOtxt
    Write-Host "Out-of-Office enabled successfully for $MailboxName!" -ForegroundColor Green
}

# Disable mode
elseif ($Mode -eq "Disable") {
    Write-Host "Disabling Out-of-Office for $MailboxName..." -ForegroundColor Yellow
    Set-MailboxAutoReplyConfiguration -Identity $MailboxName -AutoReplyState Disabled
    Write-Host "Out-of-Office disabled successfully for $MailboxName!" -ForegroundColor Green
}

# Schedule mode
elseif ($Mode -eq "Schedule") {
    Validate-DateTimeFormat -DateTime $StartTime
    Validate-DateTimeFormat -DateTime $EndTime

    Write-Host "Scheduling Out-of-Office for $MailboxName..." -ForegroundColor Yellow
    Set-MailboxAutoReplyConfiguration -Identity $MailboxName -AutoReplyState Scheduled -StartTime $StartTime -EndTime $EndTime -InternalMessage $OOOtxt -ExternalMessage $OOOtxt
    Write-Host "Out-of-Office scheduled successfully for $MailboxName!" -ForegroundColor Green
}

# Display results
Write-Host "----------------------------------------------------" -ForegroundColor Cyan
Write-Host "Out-of-Office Settings for Mailbox: $MailboxName" -ForegroundColor Cyan
Write-Host "----------------------------------------------------" -ForegroundColor Cyan
$AutoReplyConfig = Get-MailboxAutoReplyConfiguration -Identity $MailboxName

if ($AutoReplyConfig) {
    Write-Host "Auto-Reply State     : $($AutoReplyConfig.AutoReplyState)" -ForegroundColor Yellow
    Write-Host "Start Time           : $($AutoReplyConfig.StartTime)" -ForegroundColor Yellow
    Write-Host "End Time             : $($AutoReplyConfig.EndTime)" -ForegroundColor Yellow
    Write-Host "Internal Message     : $($AutoReplyConfig.InternalMessage)" -ForegroundColor Yellow
    Write-Host "External Message     : $($AutoReplyConfig.ExternalMessage)" -ForegroundColor Yellow
} else {
    Write-Host "No Out-of-Office settings found for $MailboxName." -ForegroundColor Yellow
}
Write-Host "----------------------------------------------------" -ForegroundColor Cyan
