function Load-EWS {
## Load Managed API dll
<#First, check for the EWS MANAGED API. If it is present, import the highest version. Otherwise, exit. #> 
$EWSDLL = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services'|Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Install Directory') + "Microsoft.Exchange.WebServices.dll") 
if (Test-Path $EWSDLL) 
 { 
 Write-Verbose " Loading EWS Managed API..."
 Import-Module $EWSDLL 
 } 
else 
 { 
 "This script requires the EWS Managed API." 
 "Please download and install the current version of the EWS Managed API from" 
 "https://github.com/OfficeDev/ews-managed-api" 
 "" 
 "Exiting Script." 
 exit 
 }
 # Set the Exchange Version variable
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1

$script:service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)

# Tell the service object to use credentials stored in $EWScred variable, defined before invocation by running $EWScred=get-credential

#$service.UseDefaultCredentials = $true
$service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials -ArgumentList $EWScred.UserName, $EWScred.GetNetworkCredential().Password
##CAS URL Option 1 Autodiscover
#$Service.EnableSCPLookup = $False #Skip SCP lookups to make autodiscover of cloud mailboxes work.
#$service.AutodiscoverUrl($Identity,{$True})
#The above line will only work if $Identity holds the SMTP address of the target mailbox.

#CAS URL Option 2 Hardcoded 
$localuri=[system.URI] "https://mailserverhere/ews/exchange.asmx"
$clouduri=[system.URI] "https://outlook.office365.com/ews/exchange.asmx"

# Here is the reason for running this in an EMS session.
 $script:rec = get-recipient $Identity
 $script:CloudUser = $False
 $Script:MailboxType = "Local Mailbox" 
 $service.Url = $localuri

} # End of utility function Load-EWS()
