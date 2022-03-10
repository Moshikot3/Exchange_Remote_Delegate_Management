#Written by Moshe
."\\s6540001\ntcidsccm\powershellscripts\Exchange\Delegate_Access\V3\Load-EWS.ps1"
."\\s6540001\ntcidsccm\powershellscripts\Exchange\Delegate_Access\V3\Add-EWSDelegate.ps1"
."\\s6540001\ntcidsccm\powershellscripts\Exchange\Delegate_Access\V3\Get-EWSDelegates.ps1"
."\\s6540001\ntcidsccm\powershellscripts\Exchange\Delegate_Access\V3\Remove-EWSDelegate.ps1"
."\\s6540001\ntcidsccm\powershellscripts\Exchange\Delegate_Access\V3\Set-EWSDelegates.ps1"

$credentials=get-credential


function ExchangeConnect($commandname)
{

if (!(Get-PSSession | Where { $_.ConfigurationName -eq "Microsoft.Exchange" })) { 

$Session=New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://MAILSERVERHERE/PowerShell/ -Authentication Default -Credential $Credentials -AllowRedirection 
Set-executionPolicy RemoteSigned 

if ($commandname -eq $null)
{
Import-PSSession $Session -CommandName *
}
else {
    
    Import-PSSession $Session -CommandName *$commandname*
}
}
return $Session
}
ExchangeConnect get-recipient

try{
Get-Recipient $credentials.UserName
}catch{

Write-Error "Get-Recipient Check failed, User or Password incorrect, please relaunch the script. If issue still exist Please contact System Administrator."
msg $env:username $error[0]
break;

}


Add-Type -AssemblyName PresentationFramework
#$Credentials=get-credential
$EWSCred = $credentials

function Load-Xaml {
	[xml]$xaml = Get-Content -Path "Z:\V3\XAML\Delegate_AccessV3.xaml"
	$manager = New-Object System.Xml.XmlNamespaceManager -ArgumentList $xaml.NameTable
	$manager.AddNamespace("x","http://schemas.microsoft.com/winfx/2006/xaml");
	$xamlReader = New-Object System.Xml.XmlNodeReader $xaml
	[Windows.Markup.XamlReader]::Load($xamlReader)

}




#Buttons functions
function CheckUserDetails {
    $error.clear()
    $LogTextBox.text = " "
    $script:GetDelegate = $null
    try{
        $OwnerUsername = Get-ADUser -Identity $UserNameTB.Text -ErrorAction Continue

        $GoodBadOwner.Foreground = "#FF1AA01A"
		$GoodBadOwner.Content = "✔ $($OwnerUsername.Name)"
    }catch{
        $GoodBadOwner.Foreground = "#FFE41F1F"
		$GoodBadOwner.Content = "❌"
    }

    try{
        $GiverUsername = Get-ADUser -Identity $GiveAccessToTB.Text -ErrorAction Continue

		$GoodBadClient.Foreground = "#FF1AA01A"
		$GoodBadClient.Content = "✔ $($GiverUsername.Name)"
    }catch{
		$GoodBadClient.Foreground = "#FFE41F1F"
		$GoodBadClient.Content = "❌"
        
    }
    $PermissionsLB.Items.Clear()

    $script:GetDelegates = Get-EWSDelegates $UserNameTB.Text

    foreach($DelegatedMember in ($script:GetDelegates).DelegateUserResponses.Delegateuser.Userid.Displayname)
	{

    Write-Host $DelegatedMember -ForegroundColor Green

		$PermissionsLB.Items.Add($DelegatedMember)
	}

if(($script:GetDelegates.MeetingRequestsDeliveryScope) -eq "NoForward"){
		$DelegatesAndInfoToOwnerCB.IsEnabled = $false
		$DelegatesOnlyCB.IsEnabled = $false
		$DelegatesAndOwnerCB.IsEnabled = $false
		$DelegatesAndInfoToOwnerCB.IsChecked = $false
		$DelegatesOnlyCB.IsChecked = $false
		$DelegatesAndOwnerCB.IsChecked = $false
}else{
		$DelegatesAndInfoToOwnerCB.IsEnabled = $true
		$DelegatesOnlyCB.IsEnabled = $true
		$DelegatesAndOwnerCB.IsEnabled = $true

    switch ($script:GetDelegates.MeetingRequestsDeliveryScope)
    {

    'DelegatesOnly'{$DelegatesOnlyCB.IsChecked = $True}
    'DelegatesAndMe'{$DelegatesAndOwnerCB.IsChecked = $true}
    'DelegatesAndSendInformationToMe'{$DelegatesAndInfoToOwnerCB.IsChecked = $True}


    }
    RadioChecking

}




$CopiesCB.IsEnabled = $false
$CopiesCB.IsChecked = $false
$DelegatesAndInfoToOwnerCB.IsEnabled = $false
$DelegatesOnlyCB.IsEnabled = $false
$DelegatesAndOwnerCB.IsEnabled = $false
$CalendarCB.SelectedItem = "None"
$InboxCB.SelectedItem = "None"
$ContactsCB.SelectedItem = "None"


}


function SetBtn {

     try{
        Get-ADUser -Identity $UserNameTB.Text -ErrorAction Continue

    }catch{
        Write-Error "Error: Owner username might be wrong"
        $LogTextBox.text = $error[0]
        return
    }

    $error.clear()

    if($CopiesCB.IsChecked -and $script:GetDelegates.MeetingRequestsDeliveryScope -ne $script:MeetingRequestsDeliveryScope){
     Set-EWSDelegates -Identity $UserNameTB.Text -DeliveryScope $script:MeetingRequestsDeliveryScope
    }
    Add-EWSDelegate -Identity $UserNameTB.Text -Delegate $PermissionsLB.SelectedItem -CalendarRights $CalendarCB.SelectedItem -GetsMeetingRequests $CopiesCB.IsChecked -InboxRights $InboxCB.SelectedItem -ContactsRights $ContactsCB.SelectedItem 
    $LogTextBox.text = $error[0]
    CheckUserDetails

}



function AddNew {
    try{
        Get-ADUser -Identity $UserNameTB.Text -ErrorAction Continue
        Get-ADUser -Identity $GiveAccessToTB.Text -ErrorAction Continue

    }catch{
        Write-Error "Error: One of the Usernames might be wrong"
        $LogTextBox.text = $error[0]
        return
    }
    $error.clear()
    if($CopiesCB.IsChecked -and $script:GetDelegates.MeetingRequestsDeliveryScope -ne $script:MeetingRequestsDeliveryScope){
     Set-EWSDelegates -Identity $UserNameTB.Text -DeliveryScope $script:MeetingRequestsDeliveryScope
    }

    Add-EWSDelegate -Identity $UserNameTB.Text -Delegate $GiveAccessToTB.Text -CalendarRights $CalendarCB.SelectedItem -InboxRights $InboxCB.SelectedItem -ContactsRights $ContactsCB.SelectedItem -GetsMeetingRequests $CopiesCB.IsChecked
    $LogTextBox.text = $error[0]
    CheckUserDetails

}

function RemoveBtn {
    if($PermissionsLB.SelectedItem -eq $null -or $PermissionsLB.SelectedItem -eq ""){
        Write-Error "Error: No delegate as been selected"
        $LogTextBox.text = $error[0]

        return
    }
	 $error.clear()
     Remove-EWSDelegate -Id $UserNameTB.Text -Delegate $PermissionsLB.SelectedItem
     $LogTextBox.text = $error[0]
     CheckUserDetails

}

function ListPermissions {
#Reset some variables
$script:GetDelegatePermissions = $null
$CopiesCB.IsEnabled = $false
$CopiesCB.IsChecked = $false
$DelegatesAndInfoToOwnerCB.IsEnabled = $false
$DelegatesOnlyCB.IsEnabled = $false
$DelegatesAndOwnerCB.IsEnabled = $false
$CalendarCB.SelectedItem = "None"
$InboxCB.SelectedItem = "None"
$ContactsCB.SelectedItem = "None"

##Some testing

$LogTextBox.text = (([Array]($script:GetDelegates.DelegateUserResponses.Delegateuser | ?{$_.userid.displayname -eq $PermissionsLB.SelectedItem}).Permissions) | out-string).trim()
##end testing

$script:GetDelegatePermissions = ($script:GetDelegates.DelegateUserResponses.Delegateuser | ?{$_.userid.displayname -eq $PermissionsLB.SelectedItem})
$CalendarCB.SelectedItem = [String]$script:GetDelegatePermissions.Permissions.CalendarFolderPermissionLevel
$CopiesCB.isChecked = $script:GetDelegatePermissions.ReceiveCopiesOfMeetingMessages
if($CalendarCB.SelectedItem -eq "Editor"){
            $DelegatesAndInfoToOwnerCB.IsEnabled = $true
            $DelegatesOnlyCB.IsEnabled = $true
            $DelegatesAndOwnerCB.IsEnabled = $true
            $CopiesCB.IsEnabled = $true
}else{
			$CopiesCB.IsEnabled = $false
            $CopiesCB.IsChecked = $false
}
$InboxCB.SelectedItem = [String]$script:GetDelegatePermissions.Permissions.InboxFolderPermissionLevel
$ContactsCB.SelectedItem = [String]$script:GetDelegatePermissions.Permissions.ContactsFolderPermissionLevel

	
}

function CalendarEditor {


    if($CalendarCB.SelectedItem -eq "Editor" -and $PermissionsLB.SelectedItem -eq ""){
    		$CopiesCB.IsEnabled = $false
        	$CopiesCB.isChecked = $False
            $DelegatesAndInfoToOwnerCB.IsEnabled = $false
            $DelegatesOnlyCB.IsEnabled = $false
            $DelegatesAndOwnerCB.IsEnabled = $false

    }elseif($CalendarCB.SelectedItem -eq "Editor" -and $PermissionsLB.SelectedItem -ne ""){
        $CopiesCB.IsEnabled = $true
        $CopiesCB.isChecked = [Bool]$script:GetDelegatePermissions.ReceiveCopiesOfMeetingMessages
        $DelegatesAndInfoToOwnerCB.IsEnabled = $true
        $DelegatesOnlyCB.IsEnabled = $true
        $DelegatesAndOwnerCB.IsEnabled = $true

     switch ($script:GetDelegates.MeetingRequestsDeliveryScope)
        {

        'DelegatesOnly'{$DelegatesOnlyCB.IsChecked = $True}
        'DelegatesAndMe'{$DelegatesAndOwnerCB.IsChecked = $true}
        'DelegatesAndSendInformationToMe'{$DelegatesAndInfoToOwnerCB.IsChecked = $True}


       }


    
    }else{
        	$CopiesCB.IsEnabled = $false
        	$CopiesCB.isChecked = $false

    }

}

function RadioChecking {
    $script:MeetingRequestsDeliveryScope = ""
    switch ($true)
    {

        ($DelegatesOnlyCB.IsChecked){$script:MeetingRequestsDeliveryScope = 'DelegatesOnly'}
        ($DelegatesAndOwnerCB.IsChecked){$script:MeetingRequestsDeliveryScope = 'DelegatesAndMe'}
        ($DelegatesAndInfoToOwnerCB.IsChecked){$script:MeetingRequestsDeliveryScope = 'DelegatesAndSendInformationToMe'}


    }

}



$window = Load-Xaml
$UserNameTB = $window.FindName("UserNameTB")
$GiveAccessToTB = $window.FindName("GiveAccessToTB")
$GoodBadOwner = $window.FindName("GoodBadOwner")
$GoodBadClient = $window.FindName("GoodBadClient")
$CheckBtn = $window.FindName("CheckBtn")
$CheckBtn.Add_Click({CheckUserDetails;$LogTextBox.text = $error})
$Giveaccesstotext = $window.FindName("giveaccesstotext")
$AddNewBtn = $window.FindName("AddNewBtn")
$AddNewBtn.Add_Click({ AddNew })
$PermissionsLB = $window.FindName("PermissionsLB")
$PermissionsLB.add_selectionChanged({ ListPermissions })
#CalendarCombobox
$CalendarCB = $window.FindName("CalendarCB")
$CalendarCB.add_selectionChanged({ CalendarEditor })
$CalendarCB.Items.Add("None")
$CalendarCB.Items.Add("Reviewer")
$CalendarCB.Items.Add("Author")
$CalendarCB.Items.Add("Editor")
$CalendarCB.SelectedItem = "None"
#InboxCombobox
$InboxCB = $window.FindName("InboxCB")
$InboxCB.Items.Add("None")
$InboxCB.Items.Add("Reviewer")
$InboxCB.Items.Add("Author")
$InboxCB.Items.Add("Editor")
$InboxCB.SelectedItem = "None"
#ContactsCombobox
$ContactsCB = $window.FindName("ContactsCB")
$ContactsCB.Items.Add("None")
$ContactsCB.Items.Add("Reviewer")
$ContactsCB.Items.Add("Author")
$ContactsCB.Items.Add("Editor")
$ContactsCB.SelectedItem = "None"
#RemoveBtn
$RemoveBtn = $window.FindName("RemoveBtn")
$RemoveBtn.Add_Click({ RemoveBtn })
#SetBtn
$SetBtn = $window.FindName("SetBtn")
$SetBtn.Add_Click({ SetBtn })
#Exit Button
$ExitBtn = $window.FindName("ExitBtn")
$ExitBtn.Add_Click({ exit })
#Indication Label
$IndicationLabel = $window.FindName("IndicationLabel")
#Receive copies Checkbox
$CopiesCB = $window.FindName("CopiesCB")
$CopiesCB.IsEnabled = $false


#Delegate Scopes Checkboxes
$DelegatesAndInfoToOwnerCB = $window.FindName("DelegatesAndInfoToOwnerCB")
$DelegatesOnlyCB = $window.FindName("DelegatesOnlyCB")
$DelegatesAndOwnerCB = $window.FindName("DelegatesAndOwnerCB")
$DelegatesAndInfoToOwnerCB.IsEnabled = $false
$DelegatesAndInfoToOwnerCB.Add_Click({ RadioChecking })
$DelegatesOnlyCB.IsEnabled = $false
$DelegatesOnlyCB.Add_Click({ RadioChecking })
$DelegatesAndOwnerCB.IsEnabled = $false
$DelegatesAndOwnerCB.Add_Click({ RadioChecking })

#Log Textbox

$LogTextBox = $window.FindName("Outputlog")


##Code
$window.ShowDialog()
