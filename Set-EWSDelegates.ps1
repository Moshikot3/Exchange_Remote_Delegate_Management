function Global:Set-EWSDelegates {

<#
.SYNOPSIS

Modifies global delegate settings on target mailbox.

.DESCRIPTION

Sets global delegate settings on target mailbox. This corresponds to the settings in Outlook's delegate window specifying whether meeting requests and responses are delivered to delegate only but with a copy of requests and responses delivered to delegator, delegates only, or both.

This script must be run in an Exchange Management Shell session. 

Full Access permissions must have been granted on the mailbox prior to running the script. Also, the credentials for the Full Access accunt must be placed in a variable called $EWScred via the Get-Credential command, specifying the username in UserPrincipalName format.

.PARAMETER Identifier
 
 Identifier of target mailbox.

.PARAMETER DeliveryScope

 Determines to whom meeting requests and responses are sent.

 This parameter can take the following values:
 * DelegatesOnly
 * DelegatesAndMe
 * DelegatesAndSendInformationToMe
 * NoForward

.EXAMPLE

 Set-EWSDelegates.ps1 -UPN mbx@myDomain.com -DeliveryScope DelegatesAndMe
#>

param (
 [parameter(
 Mandatory=$true,
 Position=0,
 HelpMessage="Identifier of target mailbox")]
 [Alias("Id")]
 [STRING]$Identity,

 [parameter(
 Mandatory=$True,
 Position=1,
 HelpMessage="Who should receive notifications")]
 [ValidateSet('DelegatesOnly','DelegatesAndMe','DelegatesAndSendInformationToMe','NoForward')]
 [STRING]$DeliveryScope

) 
 Write-Verbose "Loading EWS."
 Load-EWS
 Write-Verbose "Retrieving current delegate settings."
 $EWSdelegates = $service.GetDelegates($rec.PrimarySmtpAddress,$true)
 Write-Verbose "Writing back delegate object with new scope."
 $EWSupdatedDelegates = $service.UpdateDelegates($rec.PrimarySmtpAddress,$DeliveryScope,$EWSdelegates.DelegateUserResponses[0].DelegateUser)

}
