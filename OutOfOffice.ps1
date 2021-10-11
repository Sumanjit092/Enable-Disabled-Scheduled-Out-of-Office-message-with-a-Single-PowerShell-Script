####################################################################################################################
####################################################################################################################
###                      PowerShell Script for Out of Office or Auto Reply Message                               ###
###                                             Version: 1.0                                                     ###
###                                        Date: 15th April, 2021                                                ###
###                                        Script by: Sumanjit Pan                                               ###
####################################################################################################################
####################################################################################################################

# Mailbox Name to set Out-of-Office Message #
[string] $MBName = Read-Host -Prompt "Enter Mailbox to Set Out-of-Office"

While ($MBName -notlike “*@*” -or $MBName -notlike “*.*”){

Write-Host -ForeGroundColor Red “Please provide a valid Email Address!”

[string] $MBName = Read-Host -Prompt "Enter Mailbox to Set Out-of-Office"

}

# Out-of-Office Message #
Write-Host “”

# If Out of Office contain multiple line, you may use string "<br> <br>" start of every new line #

[string] $Message = Read-Host -Prompt "Paste Out-of-Office Message Here – Leave blank if you plan to Disable"

# Place Out-of-Office Message Inside HTML Tags to preserve Formatting #
$OOOtxt = ‘<pre>’ + $Message + ‘</pre>’

# Actions #
Write-Host “”

[string] $Mode = Read-Host -Prompt "(E)nable (D)isable or (S)chedule"

# (E)nable #
If ($Mode -match “E”) {

Set-MailboxAutoReplyConfiguration -Identity $MBName -AutoReplyState Enabled -InternalMessage $OOOtxt -ExternalMessage $OOOtxt

}

# (D)isable
If ($Mode -match “D”) {

Set-MailboxAutoReplyConfiguration -Identity $MBName -AutoReplyState Disabled

}

# (S)chedule
If ($Mode -match “S”) {

[string]$StartTime = Read-Host -Prompt "Enter Out-of-Office Message Start Date & Time in 'MM/DD/YYYY HH:MM:SS' Format. For Example: 05/01/2021 01:00:00"

[string]$EndTime = Read-Host -Prompt "Enter Out-of-Office Message End Date & Time in 'MM/DD/YYYY HH:MM:SS' Format. For Example: 05/01/2021 23:59:00"

Set-MailboxAutoReplyConfiguration -Identity $MBName -AutoReplyState Scheduled -StartTime $StartTime -EndTime $EndTime -InternalMessage $OOOtxt -ExternalMessage $OOOtxt

}

# Display Results
Write-Host “————————————————————————–”

Write-Host -ForegroundColor Green “The following Out-of-Office Settings have been applied to Mailbox:" $MBName

$OutPut = Get-MailboxAutoReplyConfiguration -Identity $MBName | FL -Property AutoReplyState, StartTime , EndTime, InternalMessage, ExternalMessage
Write-Output $OutPut
