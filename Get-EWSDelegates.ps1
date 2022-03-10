function Global:Get-EWSDelegates  {
param (
 [parameter(
 Mandatory=$true,
 ValueFromPipeline=$True,
 ValueFromPipelineByPropertyName=$True,
 Position=0,
 HelpMessage="Identity of target mailbox")]
 [Alias("Id")]
 [STRING]$Identity ) 

 Load-EWS
Write-Verbose "Using CAS URI $($Service.url) for $($MailboxType) $($rec.PrimarySmtpAddress)"

 #Get Delegates for Mailbox 
Write-Verbose "Getting delegates for $($rec.PrimarySmtpAddress):"
$EWSdelegates = $service.getdelegates($rec.PrimarySmtpAddress,$true)
# Parse returned object and display results
Write-Host "MeetingRequestsDeliveryScope: $($EWSdelegates.MeetingRequestsDeliveryScope)"
#  Enumerate Delegates and their settings
foreach($Delegate in $EWSdelegates.DelegateUserResponses){ 
    write-host "`n`r`tDelegate: $($Delegate.DelegateUser.UserId.DisplayName) <$($Delegate.DelegateUser.UserId.PrimarySmtpAddress)>"
    write-host "`tReceives copies of meeting messages: $($Delegate.DelegateUser.ReceiveCopiesOfMeetingMessages)"
    write-host "`tDelegate can view private items: $($Delegate.DelegateUser.ViewPrivateItems)"
    [array]$Delegate.DelegateUser.Permissions | select CalendarFolderPermissionLevel,TasksFolderPermissionLevel,InboxFolderPermissionLevel,ContactsFolderPermissionLevel,NotesFolderPermissionLevel | ft
    if ($Delegate.Result -like "Error") {
        Write-Host "Result: $($Delegate.Result)" -ForegroundColor Red
        Write-Host "ErrorCode: $($Delegate.ErrorCode)" -ForegroundColor Red
        Write-Host "ErrorMessage: $($Delegate.ErrorMessage)" -ForegroundColor Red
        Write-Error "Error while refreshing data, Please try to refresh the data again, New unfinished delegation may cause this issue so if this problem still exist please wait few moments and refresh the data again, Or contact System Administrator."
    } # End if error 
 } # End foreach Delegate in $EWSDelegates

return $EWSDelegates # Return the object to the EMS command line.
} # End function Get-EWSDelegates()
