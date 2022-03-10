function Global:Add-EWSDelegate {

<#
.SYNOPSIS

Adds delegate to target mailbox or modifies existing delegate.

.DESCRIPTION

Takes the identity of the user and adds a delegate with the specified settings, or modifies the delegate settings if it is already present.

This script must be run in an Exchange Management Shell session. 

Full Access permissions must have been granted on the mailbox prior to running the script. Also, the credentials for the Full Access accunt must be placed in a variable called $EWScred via the Get-Credential command, specifying the username in UserPrincipalName format.

.PARAMETER Identity
 
 Identity of target mailbox

.PARAMETER Delegate

 Identity of the delegate being added

.PARAMETER GetsMeetingRequests

 Boolean value specifying whether or not meeting requests are forwarded to this delegate

.PARAMETER CanSeePrivate

 Boolean value specifying whether or not this delegate can see the delegator's private items

.PARAMETER CalendarRights

 Delegate permissions on the Calendar folder. Allowed values include None, Editor, Reviewer, Author

.PARAMETER TasksRights

 Delegate permissions on the Tasks folder. Allowed values include None, Editor, Reviewer, Author

.PARAMETER InboxRights

 Delegate permissions on the Inbox folder. Allowed values include None, Editor, Reviewer, Author

.PARAMETER ContactsRights

 Delegate permissions on the Contacts folder. Allowed values include None, Editor, Reviewer, Author

.PARAMETER NotesRights

 Delegate permissions on the Notes folder. Allowed values include None, Editor, Reviewer, Author

.EXAMPLE

 Add-EWSDelegate.ps1 -Identity userA@yoyodyne.com -Delegate userB@yoyodyne.com -CalendarRights Editor -GetsMeetingRequests 
  Add-EWSDelegate.ps1 -Identity userA@yoyodyne.com -Delegate userB@yoyodyne.com -CalendarRights Editor -InboxRights Editor -ContactsRights Editor -GetsMeetingRequests 
#>

param (
 [parameter(
 Mandatory=$true,
 ValueFromPipeline=$True,
 ValueFromPipelineByPropertyName=$True,
 Position=0,
 HelpMessage="Identity of target mailbox")]
 [Alias("Id")]
 [STRING]$Identity, 

 [parameter(
 Mandatory=$true,
 Position=1,
 HelpMessage='Identity of delegate')]
 [STRING]$Delegate,

 [BOOLEAN]$GetsMeetingRequests,

 [BOOLEAN]$CanSeePrivate,

 [ValidateSet('None','Editor','Reviewer','Author')]
 [STRING]$CalendarRights,

 [ValidateSet('None','Editor','Reviewer','Author')]
 [STRING]$TasksRights,

 [ValidateSet('None','Editor','Reviewer','Author')]
 [STRING]$InboxRights,

 [ValidateSet('None','Editor','Reviewer','Author')]
 [STRING]$ContactsRights,

 [ValidateSet('None','Editor','Reviewer','Author')]
 [STRING]$NotesRights

) # End parameters

Write-Verbose "Loading EWS"
# loading the ews.dll
Load-EWS

<# Let's get the primarySmtpAddress of the new delegate so that we can properly compare with what is on the list, just in case what has been provided is a secondary address. #>
$NewDelegate=$Null
$NewDelegate = Get-Recipient $Delegate
$Delegate = $NewDelegate.PrimarySmtpAddress
if ($NewDelegate -eq $null) {
 Write-Host "Requested delegate $($Delegate) does not appear to be valid!" -ForegroundColor Red
 return $error[0] }
 
 <# Now, we need to make sure the delegate is in the same place as the target mailbox. Delegation does not work across orgs. For example, an on-premise user cannot be made a delegate for a cloud mailbox. #>
 if ($NewDelegate.RecipientType -eq "MailUser") {$DelegateMailboxType = "Remote Mailbox" }
 elseif ($NewDelegate.RecipientType -eq "UserMailbox") {$DelegateMailboxType = "Local Mailbox" }
 elseif ($NewDelegate.RecipientType -eq "MailUniversalSecurityGroup") {
 Write-Host "Groups are not supported for delegation via EWS." -Foregroundcolor Red # More about this in a bit.
 Return
 }
 else {
 Write-Host "Mailbox type $($NewDelegate.RecipientType), $($NewDelegate.RecipientTypeDetails) unsupported for delegation in EWS." -ForegroundColor Red
 return 
 }
 
 if (($DelegateMailboxType -ne "Group") -and ($DelegateMailboxType -ne $MailboxType)) {
 Write-Host "Delegate and Delegator are not in the same place!" -ForegroundColor Red
 Write-Host "Mailbox $($rec.PrimarySmtpAddress) is a $($MailboxType), but $($Delegate) is a $($DelegateMailboxType)" -ForegroundColor Red
 return
 }

Write-Verbose "Good news, everyone! $($MailboxType) $($rec.PrimarySmtpAddress) and $($DelegateMailboxType) $($Delegate) are in the same place!" 

# Our proposed delegate is legit. Let's get on with it.
# We'll retrieve a copy of the existing delegates.
 $EWSdelegates = $service.GetDelegates($rec.PrimarySmtpAddress,$true)

 # Let's see if the delegate is already present. If so, we'll just modify the settings. Otherwise, add....
 $exists = $false

 Write-Verbose "Checking to see if delegate $($Delegate) is already in the list of $($EWSdelegates.Count) delegates"
 foreach ($ExistingDelegate in $EWSdelegates.DelegateUserResponses) {
 if ($ExistingDelegate.DelegateUser.UserId.PrimarySmtpAddress.ToLower() -eq $Delegate.ToLower()) {
 Write-Host "Delegate $($Delegate) is already present in delegates list. Updating permissions."
 $exists = $True
 
# Here we modify the permissions and settings for this delegate.
 if ($CalendarRights -notlike $null) {
 $ExistingDelegate.DelegateUser.Permissions.CalendarFolderPermissionLevel = [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::$CalendarRights }
 if ($TasksRights -notlike $null) {
 $ExistingDelegate.DelegateUser.Permissions.TasksFolderPermissionLevel = [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::$TasksRights }
 if ($InboxRights -notlike $null) {
 $ExistingDelegate.DelegateUser.Permissions.InboxFolderPermissionLevel = [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::$InboxRights }
 if ($ContactsRights -notlike $null) {
 $ExistingDelegate.DelegateUser.Permissions.ContactsFolderPermissionLevel = [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::$ContactsRights }
 if ($NotesRights -notlike $null) {
 $ExistingDelegate.DelegateUser.Permissions.NotesFolderPermissionLevel = [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::$NotesRights }
 if ($GetsMeetingRequests -notlike $null) {
 $ExistingDelegate.DelegateUser.ReceiveCopiesOfMeetingMessages = $GetsMeetingRequests}
 if ($CanSeePrivate -notlike $null) {
 $ExistingDelegate.DelegateUser.ViewPrivateItems = $CanSeePrivate }

 try { $update= $service.UpdateDelegates($rec.PrimarySmtpAddress,$EWSdelegates.MeetingRequestsDeliveryScope,$ExistingDelegate.DelegateUser)}
 catch {write-host "Update delegates failed!" -ForegroundColor Red
 return $error[0] }

 
 } # End of section for modifying existing delegateee 
 
 } # end foreach existingdelegate

# Here we add the delegate if it isn't already on the list.
 if ($exists -eq $false) {
 Write-Host "Adding delegate $($Delegate)"

 $dgUser = new-object Microsoft.Exchange.WebServices.Data.DelegateUser($Delegate) 
# Use standard defaults on the following if they are not specified in parameters passed to the function.
# For ViewPrivateItems, the default is false. For ReceiveCopiesOfMeetingMessages, the default is true.

if ($CanSeePrivate -notlike $null) {
 $dgUser.ViewPrivateItems = $CanSeePrivate }
else { $dgUser.ViewPrivateItems = $null } 

if ($GetsMeetingRequests -notlike $null) {
 $dgUser.ReceiveCopiesOfMeetingMessages = $GetsMeetingRequests }
else { $dgUser.ReceiveCopiesOfMeetingMessages = $true } 

if ($CalendarRights -notlike $null) 
 {$dgUser.Permissions.CalendarFolderPermissionLevel = [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::$CalendarRights }
else {$dgUser.Permissions.CalendarFolderPermissionLevel = [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::None }
 
if ($TasksRights -notlike $null) 
 {$dgUser.Permissions.TasksFolderPermissionLevel = [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::$TasksRights }
else {$dgUser.Permissions.TasksFolderPermissionLevel = [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::None }
 
if ($InboxRights -notlike $null) 
 { $dgUser.Permissions.InboxFolderPermissionLevel = [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::$InboxRights }
else { $dgUser.Permissions.InboxFolderPermissionLevel = [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::None } 
 
 if ($ContactsRights -notlike $null)
 { $dgUser.Permissions.ContactsFolderPermissionLevel = [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::$ContactsRights } 
else { $dgUser.Permissions.ContactsFolderPermissionLevel = [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::None }
 
 if ($NotesRights -notlike $null) 
 { $dgUser.Permissions.NotesFolderPermissionLevel = [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::$NotesRights }
else { $dgUser.Permissions.NotesFolderPermissionLevel = [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::None }
 
 $dgArray = new-object Microsoft.Exchange.WebServices.Data.DelegateUser[] 1 
 $dgArray[0] = $dgUser 
 $service.AddDelegates($rec.PrimarySmtpAddress,$EWSdelegates.MeetingRequestsDeliveryScope,$dgArray) 

 } # End of section adding new delegate
 
} # End global function Add-EWSdelegate
