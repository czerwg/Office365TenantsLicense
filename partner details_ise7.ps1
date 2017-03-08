#Define CSS Style
# Variables
# where to save report
$savefile = "$env:USERPROFILE\Desktop\Office365Licenses.html"
# Where to store the admin password
$credstore ="$env:APPDATA\admincred.txt"
# Details of global admin 
$glbadmname = "gczerw@ers.ie"
# Build CSS 
$head = @’
<style>
body { background-color:#dddddd;
       font-family:Tahoma;
       font-size:12pt; }
td, th { border:1px solid black;
         border-collapse:collapse; }
th { color:white;
     background-color:black; }
table, tr, td, th { padding: 2px; margin: 0px ;
     }
tr:hover {background-color: #ffff00}
table { margin-left:50px; }

h1 {
    color: red;
     text-align: center;
}
h2 {
    color: blue;
     text-align: center;
}
h3 {
    color: blue;
     text-align: center;
     background-color: Red;
}
.due {background-color: Red }
</style>
‘@
# Define hashtable to replace SKUs to human readable. 
 $SKUs = @{
                "AAD_BASIC"                     = "Azure Active Directory Basic";`
                "AAD_PREMIUM"                   = "Azure Active Directory Premium";`
                "ADALLOM_STANDALONE"            = "Cloud App Security";`
                "ATP_ENTERPRISE"                = "Exchange Online ATP";`
                "BI_AZURE_P1"                   = "Power BI Reporting and Analytics";`
                "CRMIUR"                        = "Dynamics CRM Online Pro IUR"
                "CRMPLAN1"                      = "Dynamics CRM Online Essential";`
                "CRMPLAN2"                      = "Dynamics CRM Online Basic" ;`
                "CRMSTANDARD"                   = "Dynamics CRM Online Pro";`
                "DESKLESSPACK"                  = "O365 Enterprise K1";`
                "DESKLESSPACK_YAMMER"           = "Office 365 Enterprise K1 With Yammer";`
                "EMS"                           = "Enterprise Mobility And Security Suite";`
                "ENTERPRISEPACK"                = "O365 Enterprise E3";`
                "ENTERPRISEPREMIUM"             = "O365 Enterprise E5";`
                "ENTERPRISEPREMIUM_NOPSTNCONF"  = "O365 Enterprise E5 w/o PSTN Conf";`
                "ENTERPRISEWITHSCAL"            = "O365 Enterprise E4";`
                "EOP_ENTERPRISE"                = "Exchange Online Protection";`
                "EQUIVIO_ANALYTICS"             = "O365 Advanced eDiscovery";`
                "ERP_INSTANCE"                  = "Microsoft Power BI for Office 365";`
                "EXCHANGEARCHIVE"               = "EOA for Exchange Server";`
                "EXCHANGEARCHIVE_ADDON"         = "EOA for Exchange Online";`
                "EXCHANGEDESKLESS"              = "Exchange Online Kiosk";`
                "EXCHANGEENTERPRISE"            = "Exchange Online (Plan 2)";`
                "EXCHANGESTANDARD"              = "Exchange Online (Plan 1)";`
                "EXCHANGE_ANALYTICS"            = "Delve Analytics";`
                "INTUNE_A"                      = "Intune";`
                "INTUNE_STORAGE"                = "Intune Extra Storage";`
                "LITEPACK"                      = "O365 Small Business";`
                "LITEPACK_P2"                   = "O365 Small Business Premium";`
                "LOCKBOX"                       = "Customer Lockbox";`
                "MCOEV"                         = "SfB Cloud PBX";`
                "MCOIMP"                        = "SfB Online (Plan 1)";`
                "MCOMEETADV"                    = "SfB PSTN Conferencing";`
                "MCOPLUSCAL"                    = "SfB Plus CAL";`
                "MCOPSTN1"                      = "SfB PSTN Dom. Calling";`
                "MCOPSTN2"                      = "SfB PSTN Dom. and Int. Calling";`
                "MCOSTANDARD"                   = "SfB Online (Plan 2)";`
                "O365_BUSINESS"                 = "O365 Business";`
                "O365_BUSINESS_ESSENTIALS"      = "O365 Business Essentials";`
                "O365_BUSINESS_PREMIUM"         = "O365 Business Premium";`
                "OFFICESUBSCRIPTION"            = "O365 ProPlus";`
                "PLANNERSTANDALONE"             = "Office 365 Planner";`
                "POWERAPPS_INDIVIDUAL_USER"     = "Microsoft PowerApps and logical Streams";`
                "POWER_BI_ADDON"                = "Power BI Add-on";`
                "POWER_BI_INDIVIDUAL_USE"       = "Power BI Individual User";`
                "POWER_BI_PRO"                  = "Power BI (Pro)";`
                "POWER_BI_STANDALONE"           = "Power BI Stand Alone";`
                "POWER_BI_STANDARD"             = "Power BI (free)";`
                "PROJECTCLIENT"                 = "Project Pro for O365";`
                "PROJECTESSENTIALS"             = "Project Lite";`
                "PROJECTONLINE_PLAN_1"          = "Project Online";`
                "PROJECTONLINE_PLAN_2"          = "Project Online and Pro";`
                "RIGHTSMANAGEMENT"              = "Azure Rights Management Premium";`
                "RIGHTSMANAGEMENT_ADHOC"        = "Windows Azure Rights Management";`
                "SHAREPOINTENTERPRISE"          = "SharePoint Online (Plan 2)";`
                "SHAREPOINTSTANDARD"            = "SharePoint Online (Plan 1)";`
                "SHAREPOINTSTORAGE"             = "O365 Extra File Storage";`
                "SMB_BUSINESS"                  = "O365 Business";`
                "SMB_BUSINESS_ESSENTIALS"       = "O365 Business Essentials";`
                "SMB_BUSINESS_PREMIUM"          = "O365 Business Premium";`
                "STANDARDPACK"                  = "O365 Enterprise E1";`
                "STANDARDWOFFPACK"              = "O365 Enterprise E2 (Nonprofit E1)";`
                "STREAM"                        = "Microsoft Stream"; `
                "VISIOCLIENT"                   = "Visio Pro for O365";`
                "WACONEDRIVEENTERPRISE"         = "OneDrive for Business (Plan 2)";`
                "WACONEDRIVESTANDARD"           = "OneDrive for Business (Plan 1)";`
                "YAMMER_ENTERPRISE_STANDALONE"  = "Yammer Enterprise";`
                
        }

if (!(Get-Module -ListAvailable -Name MsOnline)) {
     Write-Warning "Install office 365 admin powershell module"
     break
}

#Requires -Modules MsOnline


if (!(Test-Path $credstore)) 
{
read-host -assecurestring -Prompt "Enter Office 365 $glbadmname credentials or `nchange variable /glbadmname/ on the top of script" | convertfrom-securestring | out-file $credstore
}

try{
$password = get-content $credstore | convertto-securestring -ErrorAction Stop
$cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $glbadmname,$password -ErrorAction Stop
}
catch
{
Write-Warning "Something is wrong with stored credentials try delete file $credstore and save credentials again"
}


#$cred = Get-Credential -Message "Provide your Office 365 username and password"
try
{
    Connect-MsolService -Credential $cred -ErrorAction Stop
}
catch
{
    Write-Warning "Bad username or password, ask $glbadmname if password changed or`nSomething is wrong with stored credentials try delete file $credstore and save credentials again"
    Read-Host "Press ENTER to exit"
    break
}

if (Test-Path $savefile){Write-host "Report already exist it will be overwritten"; Read-Host "Press ENTER to continue"}
Write-Warning "Please Wait building the report" 

$partner = Get-MsolPartnerContract
$Subscriptions= $partner | %{
$ClientName=$_.name
Get-MsolSubscription -TenantId $_.tenantid| ? NextLifecycleDate -gt 0 |
select @{l="Company";e={$ClientName}},Skupartnumber,@{l="RenewalDate";e={(get-date ($_.NextlifecycleDate) -Format d) }},TotalLicenses,status
}
# Licensing in warning state
$warn =$Subscriptions | ? Status -Like 'Warning' |Sort-Object {[System.DateTime]::ParseExact($_.RenewalDate, "dd/MM/yyyy", $null)}| ConvertTo-Html -Fragment -PreContent "<H2>Licenses in warning state = renewal passed/limited functionality</H2>" |Out-String
# Licensing in warning state
$susp =$Subscriptions | ? Status -Like 'Suspended' | Sort-Object {[System.DateTime]::ParseExact($_.RenewalDate, "dd/MM/yyyy", $null)}| ConvertTo-Html -Fragment -PreContent "<H2>Licenses in Suspended state = canceled(data deleted)</H2>" |Out-String
# Licensing enabled state - sorted by renewal date, 
#$enabled =$Subscriptions | ? {$_.Status -NotMatch 'Warning|Suspended'} | Sort-Object {[System.DateTime]::ParseExact($_.RenewalDate, "dd/MM/yyyy", $null)} | ConvertTo-Html -Fragment -PreContent "<H2>Licenses in state enabled sorted by renewal date</H2>" |Out-String
$enabled =$Subscriptions | ? {$_.Status -NotMatch 'Warning|Suspended'} | Sort-Object Company | ConvertTo-Html -Fragment |Out-String

## https://www.petri.com/creating-colorful-emails-with-powershell
[xml]$en =$enabled
1..($en.table.tr.Count-1) |foreach {
$td= $en.table.tr[$_]
#$td.childnodes.item(2).'#text'
$class=$en.CreateAttribute("class")
if ( ([System.DateTime]::ParseExact($td.childnodes.item(2)."#text", "dd/MM/yyyy", $null)).adddays(-30) -lt (Get-Date))
{
$class.Value = "due"}
#$td.childnodes.item(2).attributes.append($class) | Out-Null
$en.table.tr[$_].attributes.append($class)  | Out-Null
}
$enabled1= "<H2>Licenses in state enabled</H2> <H3>Red color renewal in next 30 days.</H3>" + $en.InnerXml  |Out-String
# Assembly all parts:
$complete = ConvertTo-Html -head $head -PostContent $susp,$warn,$enabled1 -PreContent "<H1>Licenses status Office365 - created on $(Get-Date -Format d) </H2>" 
# Convert the internal SKUs to human readable.
 foreach ($key in $SKUs.Keys.GetEnumerator()) {
 $complete= $complete.Replace($key,$SKUs.$key)
 }
#Save the file
$complete| Out-File $savefile