Connect-AzAccount 
Set-AzContext -SubscriptionName 'Subscription Name'

$P_Common = "443, 80"
$P_Exchange = "143, 993, 25 443, 587, 80, 995"
$P_SharePoint = "443, 80"
$P_Skype = "443,80"


#Get the latest endpoints information from Microsoft
$office365IPs=Invoke-webrequest -Uri https://endpoints.office.com/endpoints/worldwide?clientrequestid=b10c5ed1-bad1-445f-b386-b919946339a7 | ConvertFrom-Json

$servicearea=($office365IPs.ServiceArea | sort | select -Unique)

Foreach ($area in $servicearea)

{
#Write-Output $area
#($office365IPs | Where-Object {$_.ServiceArea -eq $area }).tcpPorts | sort | select -Unique 

($office365IPs | Where-Object {$_.ServiceArea -eq $area }).ips | sort | select -Unique  |  Out-File -FilePath C:\temp\"$area"IPs.txt

($office365IPs | Where-Object {$_.ServiceArea -eq $area }).urls | sort | select -Unique |Out-File -FilePath C:\temp\"$area"URLs.txt

} 


$CommonURLs = Get-Content -Path C:\temp\CommonURLs.txt
$ExchangeURLs = ((Get-Content -Path C:\temp\ExchangeURLs.txt ) -replace "autodiscover.*.onmicrosoft.com", "*.onmicrosoft.com" )
$SkypeURLs = Get-Content -Path C:\temp\SkypeURLs.txt
$ShaepointURLs = Get-Content -Path C:\temp\SharePointURLs.txt

$commonIPs = ($IPs = Get-Content -Path C:\temp\CommonIPs.txt|  Select-String -Pattern "([0-9]{1,3}\.){3}[0-9]{1,3}(\/([0-9][0-9]|[0-9]))" -AllMatches).Matches.Value
$ExchangeIPs = ($IPs = Get-Content -Path C:\temp\ExchangeIPs.txt |  Select-String -Pattern "([0-9]{1,3}\.){3}[0-9]{1,3}(\/([0-9][0-9]|[0-9]))" -AllMatches).Matches.Value
$SkypeIPs = ($IPs = Get-Content -Path C:\temp\SkypeIPs.txt |  Select-String -Pattern "([0-9]{1,3}\.){3}[0-9]{1,3}(\/([0-9][0-9]|[0-9]))" -AllMatches).Matches.Value
$SharepointIPs = ($IPs = Get-Content -Path C:\temp\SharePointIPs.txt |  Select-String -Pattern "([0-9]{1,3}\.){3}[0-9]{1,3}(\/([0-9][0-9]|[0-9]))" -AllMatches).Matches.Value


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
$NetworkRuleCollection.AddRule($IPrule4)

$Azfw.ApplicationRuleCollections = $RuleCollection
$Azfw.NetworkRuleCollections = $NetworkRuleCollection
Set-AzFirewall -AzureFirewall $Azfw

