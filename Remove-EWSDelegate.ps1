function Global:Remove-EWSDelegate {

<#
.SYNOPSIS

Removes specified delegate from target mailbox

.DESCRIPTION

Takes the identifier of the target mailbox and removes the specified delegate from that mailbox.

This script must be run in an Exchange Management Shell session. 

Full Access permissions must have been granted on the mailbox prior to running the script. Also, the credentials for the Full Access accunt must be placed in a variable called $EWScred via the Get-Credential command, specifying the username in UserPrincipalName format.

.PARAMETER Identity
 
 Identifier of the target mailbox

.PARAMETER Delegate

 Identifier of delegate to be removed

.EXAMPLE

 Remove-EWSDelegate.ps1 -Id joe.user@mydomain.com -Delegate my.delegate@mydomain.com
#>


param (
 [parameter(
 Mandatory=$true,
 ValueFromPipeline=$True,
 ValueFromPipelineByPropertyName=$True,
 Position=0,
 HelpMessage="Identifier of target mailbox")]
 [Alias("Id")]
 [STRING]$Identity,

 [parameter(
 Mandatory=$true,
 Position=1,
 HelpMessage='Identifier of delegate to be removed')]
 [STRING]$Delegate
) 
 

 $DelegateToBeRemoved=$Null
 $DelegateToBeRemoved = Get-Recipient $Delegate

if ($DelegateToBeRemoved.PrimarySmtpAddress -eq $null) {
 Write-Host "Specified delegate $($Delegate) does not appear to be valid!" -ForegroundColor Red
 return $error[0] }

 $Delegate = $DelegateToBeRemoved.PrimarySmtpAddress

 Write-Verbose "Loading EWS"
 Load-EWS

 Write-Verbose "Retrieving current list of delegates."
 $EWSdelegates = $service.GetDelegates($rec.PrimarySmtpAddress,$true)

 $removed=$false

 Write-Verbose "Searching current delegate list for delegate to be removed."
 foreach($CurrentDelegate in $EWSdelegates.DelegateUserResponses){ 
 
 if($CurrentDelegate.DelegateUser.UserId.PrimarySmtpAddress.ToLower() -eq $Delegate.ToLower()){ 
 Write-Host "Removing Delegate $($Delegate)!"
 $service.RemoveDelegates($rec.PrimarySmtpAddress,$Delegate) 
 $removed=$true
 } # End of handler for matching delegate.
} # End iterating through list of delegates.
 if ($removed -eq $false) {Write-Host "Delegate $($Delegate) not found! Verify address." -ForegroundColor Yellow }
} # End of global function Remove-EWSDelegate()
