## Auth using the desired Azure Automation account or service principal
## First create an automation account (Run As Account), import Modules Az.Accounts, Az.Compute, Az.Network, Az.Profile, Az.Resources, run this
## Can change the ENV:TEMP$ to a storage account.
## This script is under contruction use it carefully. Any contribution are welcome. @Rafael Egito


$cred = "AzureRunAsConnection"
try {
    # Get the connection "AzureRunAsConnection"
    $servicePrincipalConnection = Get-AutomationConnection -Name $cred        

    Connect-AzAccount `
        -ServicePrincipal `
        -TenantId $servicePrincipalConnection.TenantId `
        -ApplicationId $servicePrincipalConnection.ApplicationId `
        -CertificateThumbprint $servicePrincipalConnection.CertificateThumbprint
}
catch {
 if (!$servicePrincipalConnection)
 {
 $ErrorMessage = "Connection $cred not found."
 throw $ErrorMessage
 } else{
 Write-Error -Message $_.Exception
 throw $_.Exception
 }
}

# web service root URL
$ws = "https://endpoints.office.com"
# path where output files will be stored
$versionpath = $Env:TEMP + "\O365_endpoints_latestversion.txt"
$datapath = $Env:TEMP + "\O365_endpoints_data.txt"

# fetch client ID and version if version file exists; otherwise create new file and client ID
if (Test-Path $versionpath) {
    $content = Get-Content $versionpath
    $clientRequestId = $content[0]
    $lastVersion = $content[1]
    Write-Output ("Version file exists! Current version: " + $lastVersion)
}
else {
    $clientRequestId = [GUID]::NewGuid().Guid
    $lastVersion = "0000000000"
    @($clientRequestId, $lastVersion) | Out-File $versionpath
}

# call version method to check the latest version, and pull new data if version number is different
$version = Invoke-RestMethod -Uri ($ws + "/version/Worldwide?clientRequestId=" + $clientRequestId)
if ($version.latest -gt $lastVersion) {

#Get the latest endpoints information from Microsoft
$office365IPs = Invoke-webrequest -Uri ($ws + "/endpoints/Worldwide?clientRequestId=" + $clientRequestId) | ConvertFrom-Json

$servicearea=($office365IPs.ServiceArea | sort | select -Unique)

Foreach ($area in $servicearea)

{
#Write-Output $area
#($office365IPs | Where-Object {$_.ServiceArea -eq $area }).tcpPorts | sort | select -Unique 

($office365IPs | Where-Object {$_.ServiceArea -eq $area }).ips | sort | select -Unique  |  Out-File -FilePath ($Env:TEMP + "\"+$area+"IPs.txt")

($office365IPs | Where-Object {$_.ServiceArea -eq $area }).urls | sort | select -Unique |Out-File -FilePath ($Env:TEMP + "\"+$area+"URLs.txt")

} 


$CommonURLs = Get-Content -Path ($Env:TEMP + "\CommonURLs.txt")
$ExchangeURLs = ((Get-Content -Path ($Env:TEMP +  "\ExchangeURLs.txt") ) -replace "autodiscover.*.onmicrosoft.com", "*.onmicrosoft.com" )
$SkypeURLs = Get-Content -Path ($Env:TEMP +  "\SkypeURLs.txt")
$ShaepointURLs = Get-Content -Path ($Env:TEMP + "\SharePointURLs.txt")

$commonIPs = ($IPs = Get-Content -Path ($Env:TEMP + "\CommonIPs.txt") |  Select-String -Pattern "([0-9]{1,3}\.){3}[0-9]{1,3}(\/([0-9][0-9]|[0-9]))" -AllMatches).Matches.Value
$ExchangeIPs = ($IPs = Get-Content -Path ($Env:TEMP + "\ExchangeIPs.txt") |  Select-String -Pattern "([0-9]{1,3}\.){3}[0-9]{1,3}(\/([0-9][0-9]|[0-9]))" -AllMatches).Matches.Value
$SkypeIPs = ($IPs = Get-Content -Path ($Env:TEMP + "\SkypeIPs.txt") |  Select-String -Pattern "([0-9]{1,3}\.){3}[0-9]{1,3}(\/([0-9][0-9]|[0-9]))" -AllMatches).Matches.Value
$SharepointIPs = ($IPs = Get-Content -Path ($Env:TEMP + "\SharePointIPs.txt") |  Select-String -Pattern "([0-9]{1,3}\.){3}[0-9]{1,3}(\/([0-9][0-9]|[0-9]))" -AllMatches).Matches.Value


$RG = "rg-networking-us"


#Add a rule to allow URLs
$Azfw = Get-AzFirewall -ResourceGroupName $RG 
$Rule1 = New-AzFirewallApplicationRule -Name rule_common -Protocol "http:80","https:443" -TargetFqdn  $CommonURLs
$Rule2 = New-AzFirewallApplicationRule -Name rule_exchange -Protocol "http:80","https:443" -TargetFqdn  $ExchangeURLs
$Rule3 = New-AzFirewallApplicationRule -Name rule_skype -Protocol "http:80","https:443" -TargetFqdn $SkypeURLs
$Rule4 = New-AzFirewallApplicationRule -Name rule_sharepoint -Protocol "http:80","https:443" -TargetFqdn $ShaepointURLs

$RuleCollection = New-AzFirewallApplicationRuleCollection -Name RC_O365 -Priority 100 -Rule $Rule1 -ActionType "Allow"

$RuleCollection.AddRule($Rule2)
$RuleCollection.AddRule($Rule3)
$RuleCollection.AddRule($Rule4)

$IPrule1 = New-AzFirewallNetworkRule -Name "CommonIPs" -Description "O365 common all traffic" -Protocol "Any" -SourceAddress "*" -DestinationAddress $commonIPs -DestinationPort "443", "80"
$IPrule2 = New-AzFirewallNetworkRule -Name "ExchangeIPs" -Description "Exhachange all traffic" -Protocol "Any" -SourceAddress "*" -DestinationAddress $ExchangeIPs -DestinationPort "143", "993", "25", "443", "587", "80", "995"
$IPrule3 = New-AzFirewallNetworkRule -Name "SkypeIPs" -Description "O365 Skype all traffic" -Protocol "Any" -SourceAddress "*" -DestinationAddress $SkypeIPs -DestinationPort "443", "80"
$ÌPrule4 = New-AzFirewallNetworkRule -Name "SharePointIPs" -Description "O365 Sharepoint all traffic" -Protocol "Any" -SourceAddress "*" -DestinationAddress $SharepointIPs -DestinationPort "443", "80"

$NetworkRuleCollection = New-AzFirewallNetworkRuleCollection -Name RC_O365_network -Priority 100 -Rule $IPrule1 -ActionType "Allow"
$NetworkRuleCollection.AddRule($IPrule2)
$NetworkRuleCollection.AddRule($IPrule3)
#$NetworkRuleCollection.AddRule($IPrule4)

$Azfw.ApplicationRuleCollections = $RuleCollection
$Azfw.NetworkRuleCollections = $NetworkRuleCollection
Set-AzFirewall -AzureFirewall $Azfw

$version.latest |Out-File $versionpath


}
else {
    Write-Host "Office 365 worldwide commercial service instance endpoints are up-to-date."
}

