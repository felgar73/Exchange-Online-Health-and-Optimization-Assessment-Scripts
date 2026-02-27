# For collecting EXO Hybrid Configuration settings. 
# Requires EXO Powershell module and Graph Powershell module for collection of Entra ID info related to OAuth configuration.

#Check for 'run as admin':
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
    {Write-host -ForegroundColor Yellow "Please close window and re-run powershell 'as administrator'."
    exit
    #Write-Host -ForegroundColor Cyan "& "C:\Users\$env:username\Desktop\Microsoft Exchange Online Powershell Module.appref-ms""
        } else {Write-Host -ForegroundColor Cyan "Checking execution policy..."
    }

#Execution Policy Check:
$execPol = Get-ExecutionPolicy
if ($execPol -ne 'Unrestricted'){
    Write-Host -ForegroundColor Cyan "Execution policy is" $execPol
    Write-Host -ForegroundColor Cyan "Changing policy to 'Unrestricted'..."
    Set-ExecutionPolicy Unrestricted -Force
}else {Write-Host -ForegroundColor Cyan "Execution policy is already '$execPol', continuing..."}

#Prompt for Collection Folder location:
Write-Host -ForegroundColor Yellow "Enter a root location for the collection folder (Ex: c:\temp); no trailing backslash:"
$rootFolder = Read-Host 

#Collection Folder variables:
$date = Get-Date -UFormat %b-%d-%Y
$outputDir = $rootFolder + '\HybridConfigs' + '_' + $date
$cloudDir = $outputDir + '\Office365'
$errorLog = $outputDir + '\ErrorLog.log'

#Cloud output paths:
$clIntraOrgtxtpath = $cloudDir + '\Get-IntraOrgConnector_EXO.txt'
$clIOConnjson = $cloudDir + '\Get-IntraOrganizationConnector.json'
$clOauthpath = $cloudDir + '\Oauth-Configs_EXO.txt'
$clSharingPath = $cloudDir + '\Sharing-Configs_EXO.txt'
$exosharingpoljson = $cloudDir + '\Get-SharingPolicy.json'
$cloudOrgpath = $cloudDir + '\Get-OrganizationConfig_EXO.txt'
$hcwLogsCloudpath = $cloudDir + '\HCW-LogCmds_M365.txt'
$inbConnpath = $cloudDir + '\Inbound-Connector_EXO.txt'
$inbConnjson = $cloudDir + '\Get-InboundConnector.json'
$outbConnpath = $cloudDir + '\Outbound-Connector_EXO.txt'
$outbConnjson = $cloudDir + '\Get-OutboundConnector.json'
$migEndpath = $cloudDir + '\Migration-Endpoints_EXO.txt'
$accDompath = $cloudDir + '\Accepted-Domain_EXO.txt'
$exoacceptedDomjson = $cloudDir + '\Get-AcceptedDomain.json'
$remDomEXOPath = $cloudDir + '\RemoteDomains_EXO.txt'
$exoRemdomjson = $cloudDir + '\Get-RemoteDomain.json'
$onPremOrgpath = $cloudDir + '\Get-OnPremisesOrganization.txt'
$clfederatConfpath = $cloudDir + '\Federation-Configs_EXO.txt'
#$exofederatConfxml = $cloudDir + '\Federation-Trust_EXO.xml'
$exoFedTrustjson = $cloudDir + '\Get-FederationTrust.json'
$clfedorgIdjson = $cloudDir + '\Get-FederatedOrganizationIdentifier.json'
$EntraSpnpath = $cloudDir + '\Entra-SvcPrincipals-OAuth.txt'
$exoSvcPrincjson = $cloudDir + '\Get-MgServicePrincipal-EXO.json'
$exoSvcPrincxml = $cloudDir + '\Get-MgServicePrincipal-EXO.xml'
$spoSvcPrincjson = $cloudDir + '\Get-MgServicePrincipal-SPO.json'
$hybridAppjson = $cloudDir + '\Get-MgServicePrincipal-HybridApp.json'
$hybridAppxml = $cloudDir + '\Get-MgServicePrincipal-HybridApp.xml'
$exoAddpoltxt = $cloudDir + '\EmailAddressPolicies_EXO.txt'
$exoAddpoljson = $cloudDir + '\Get-EmailAddressPolicy.json'
$MigServerTestJson = $cloudDir + '\Test-MigrationServerAvailability_AutoD.json'
$exoOrgReljson =  $cloudDir + '\Get-OrganizationRelationship.json'

#Create Collection Folder:
if (!(Test-Path $outputDir)){
    New-Item -itemtype Directory -Path $outputDir
    }

#Collecting HCW "XHCW" file info:
$hcwPath = Get-ChildItem "$env:APPDATA\Microsoft\Exchange Hybrid Configuration\*.xhcw" -ErrorAction SilentlyContinue | ForEach-Object {Get-Content $_.fullname} -ErrorAction SilentlyContinue
[XML]$hcwLog = "<root>$($hcwPath)</root>"

#Find all 'Set, New, Remove', cmdlets executed by HCW against On-Prem:
function HCWLogs-OnPrem {
if (!(!$hcwPath)){
$title1 = "===='Set-' Commands Executed On-Premises===="
$title1 | Out-File $hcwLogsOPPath
$hcwLog.SelectNodes('//invoke') | Where-Object {$_.cmdlet -like "*Set*" -and $_.type -like "*OnPremises*"} | ForEach-Object {
    New-Object -Type PSObject -Property @{
        'Date'=$_.start;
        'Duration'=$_.duration;
        'Session'=$_.type;
        'Cmdlet'=$_.cmdlet;
        'Comment'=$_.'#comment'
        }
    } | Out-File -Append $hcwLogsOPPath

$title2 = "===='New-' Commands Executed On-Premises===="
$title2 | Out-File -Append $hcwLogsOPPath
$hcwLog.SelectNodes('//invoke') | Where-Object {$_.cmdlet -like "*New*" -and $_.type -like "*OnPremises*"} | ForEach-Object {
    New-Object -Type PSObject -Property @{
        'Date'=$_.start;
        'Duration'=$_.duration;
        'Session'=$_.type;
        'Cmdlet'=$_.cmdlet;
        'Comment'=$_.'#comment'
        }
    } | Out-File -Append $hcwLogsOPPath

$title3 = "===='Remove-' Commands Executed On-Premises===="
$title3 | Out-File -Append $hcwLogsOPPath
$hcwLog.SelectNodes('//invoke') | Where-Object {$_.cmdlet -like "*Remove*" -and $_.type -like "*OnPremises*"} | ForEach-Object {
    New-Object -Type PSObject -Property @{
        'Date'=$_.start;
        'Duration'=$_.duration;
        'Session'=$_.type;
        'Cmdlet'=$_.cmdlet;
        'Comment'=$_.'#comment'
        }
    } | Out-File -Append $hcwLogsOPPath
} else {
    Write-Host -ForegroundColor White "No HCW logs found on this machine..."
}
}

#Find all 'Set, New, Remove', cmdlets executed by HCW against EXO:
function HCWLogs-Cloud {
if (!(!$hcwPath)) {
$title = "===='Set-' Commands Executed in M365===="
$title | Out-File $hcwLogsCloudpath
$hcwLog.SelectNodes('//invoke') | Where-Object {$_.cmdlet -like "*Set*" -and $_.type -like "*Tenant*"} | ForEach-Object {
    New-Object -Type PSObject -Property @{
        'Date'=$_.start;
        'Duration'=$_.duration;
        'Session'=$_.type;
        'Cmdlet'=$_.cmdlet;
        'Comment'=$_.'#comment'
        }
    } | Out-File -Append $hcwLogsCloudpath

$title = "===='New-' Commands Executed in M365===="
$title | Out-File -Append $hcwLogsCloudpath
$hcwLog.SelectNodes('//invoke') | Where-Object {$_.cmdlet -like "*New*" -and $_.type -like "*Tenant*"} | ForEach-Object {
    New-Object -Type PSObject -Property @{
        'Date'=$_.start;
        'Duration'=$_.duration;
        'Session'=$_.type;
        'Cmdlet'=$_.cmdlet;
        'Comment'=$_.'#comment'
        }
    } | Out-File -Append $hcwLogsCloudpath

$title = "===='Remove-' Commands Executed in M365===="
$title | Out-File -Append $hcwLogsCloudpath
$hcwLog.SelectNodes('//invoke') | Where-Object {$_.cmdlet -like "*Remove*" -and $_.type -like "*Tenant*"} | ForEach-Object {
    New-Object -Type PSObject -Property @{
        'Date'=$_.start;
        'Duration'=$_.duration;
        'Session'=$_.type;
        'Cmdlet'=$_.cmdlet;
        'Comment'=$_.'#comment'
        }
    } | Out-File -Append $hcwLogsCloudpath
    }else {
        Write-Host -ForegroundColor Magenta "No HCW logs found on this machine..."
    }
}
#Cloud folder creation:
function CloudDir-Create {
    Write-Host -ForegroundColor Cyan "Creating collection folder..."
    if (!(Test-Path $cloudDir)) {
    New-Item -itemtype Directory -Path $cloudDir
    }
}

#Collect EXO Yes-No Function:
function EXO-RemoteQ {
    Write-Host -ForegroundColor Yellow "Do you wish to collect M365 data? Y/N:"
    $ans = Read-Host
    if ((!$ans) -or ($ans -eq 'y') -or ($ans -eq 'yes')){
        $ans = 'yes'
        #Create collection folder:
        CloudDir-Create
        if ($null -eq $domain) {
            Write-Host -ForegroundColor Yellow "Enter your vanity domain name:"
            $domain = Read-Host
        }
        Write-Host -ForegroundColor Cyan "Checking HCW logs..."
        HCWLogs-Cloud
        #Connect to EXO & Collect Data:
        Remote-EXOPS
        #Collect Entra Data:
        AAD-Collection
    } else {
            $ans = 'no'
            Write-Host -ForegroundColor Cyan "Skipping M365 data collection..."
        }
}
#Remote EXO PS Function:
function Remote-EXOPS {
    Write-Host -ForegroundColor Cyan "Connecting to Exchange Online..."
    Write-Host -ForegroundColor Yellow "Enter your M365 admin username (Ex: admin@yourdomain.onmicrosoft.com):"
    $exoUPN = Read-Host
    try {
        Import-Module ExchangeOnlineManagement
        Connect-ExchangeOnline -UserPrincipalName $exoUPN -ShowBanner:$false
    } catch {
        $exoV3Fail = "EXO Remote PS Connection Failed"
        }
    if ($null -ne $exoV3Fail) {
        Write-Host -ForegroundColor Cyan "Remote EXO connection failed." 
        Write-Host -ForegroundColor Cyan "Ensure that the EXO V3 module is installed (https://aka.ms/exops-docs) and that basic auth is disabled in your tenant." 
    }
    else {
        EXO-Collection
    }
}

#Service Principal Collection for OAuth:
function AAD-Collection {
Write-Host -ForegroundColor Cyan "Connecting to Entra ID..."
    try {
        #Install-Module Microsoft.Graph -Scope CurrentUser (if not installed, this will install the MS Graph module)
        Connect-MgGraph -Scopes "Application.Read.All" -NoWelcome
    }
    catch {
        $GraphFail = "Connection to Entra ID Failed. Ensure MgGraph Powershell Module is installed and run the cmdlet below manually to collect OAuth service principal info:"
        $AADSvcPrinCmdlet = "Get-MgServicePrincipal -Filter 'AppId eq '00000002-0000-0ff1-ce00-000000000000''"
    }
    if ($null -ne $GraphFail) {
        Write-Host -ForegroundColor Yellow $GraphFail
        Write-Host -ForegroundColor White $AADSvcPrinCmdlet
        Write-Host -ForegroundColor Cyan "Refer to: https://learn.microsoft.com/en-us/powershell/microsoftgraph/installation?view=graph-powershell-1.0 for additional details on MS Graph module."
    } else {
        $svcPrinText = "=== Entra ID Service Principals for OAuth ==="
        $svcPrinText | Out-File -Append $EntraSpnpath
        $exoSvcId = '00000002-0000-0ff1-ce00-000000000000'
        $skypeSvcId = '00000004-0000-0ff1-ce00-000000000000'
        $exosvcPrinc = Get-MgServicePrincipal -Filter "AppId eq '$exoSvcid'"
        $skypsvcPrinc = Get-MgServicePrincipal -Filter "AppId eq '$skypeSvcid'"
        $hybridapp = Get-MgServicePrincipal -Filter "startsWith(DisplayName, 'ExchangeServerApp')"
        $exosvcPrinc | ConvertTo-Json | Out-File $exoSvcPrincjson
        $exosvcPrinc | Export-Clixml $exoSvcPrincxml
        $skypsvcPrinc | ConvertTo-Json | Out-File $spoSvcPrincjson
        $exosvcPrinc | FL AppDisplayName,ObjectType,AccountEnabled,AppId | Out-File -Append $EntraSpnpath

        Add-Content $EntraSpnpath -Value "Registered 'ServicePrincipalNames':"
        Add-Content $EntraSpnpath -Value ""
        $exosvcPrinc | Select -ExpandProperty ServicePrincipalNames | Out-File -Append $EntraSpnpath
        Add-Content $EntraSpnpath -Value ""

        Add-Content $EntraSpnpath -Value "===Hybrid Application Info===:"
        $hybridapp | FL DisplayName,Appid,Description,Notes | Out-File -Append $EntraSpnpath
    }
}

#EXO Data Collection:
function EXO-Collection {
Write-Host -ForegroundColor Cyan "Collecting data from Exchange Online..."
$shpol = "===Sharing Policy Details===:"
$shpol | Out-File $clSharingPath
$sharePol = Get-SharingPolicy 
$sharePol |FL | Out-File -Append $clSharingPath
$sharePol | ConvertTo-Json | Out-File $exosharingpoljson
Add-Content $clSharingPath -Value "===Org Relationship Details===:"
$orgRel = Get-OrganizationRelationship 
$orgRel |FL | Out-File -Append $clSharingPath
$orgRel | ConvertTo-Json | Out-File $exoOrgReljson

#Fed Org Info:
$fedoiText = "===Federated Organization Information==="
$fedoiText | Out-File $clfederatConfpath
$fedOI = Get-FederatedOrganizationIdentifier
$fedOI | FL | Out-File -Append $clfederatConfpath
$fedOI | ConvertTo-Json | Out-File $clfedorgIdjson
Start-Sleep -Seconds 2
$fedtrusttext = "===Federation Trust Info===:"
$fedtrusttext | Out-File -Append $clfederatConfpath
$fedtrustexo = Get-FederationTrust
#$fedtrustexo | Export-Clixml $exofederatConfxml
$fedtrustexo | ConvertTo-Json | Out-File $exoFedTrustjson
$fedtrustexo |FL Name,TokenIssuer*,WebRequestorRedirectEpr | Out-File -Append $clfederatConfpath
$fedtrusttext = "For full Federation Trust details, see 'Get-FederationTrust.json' file."
$fedtrusttext | Out-File -Append $clfederatConfpath

$cloudOrgtext = "===Organization Config Details===:"
$cloudOrgtext | Out-File $cloudOrgpath
$orgConfig = Get-OrganizationConfig 
$orgConfig |FL | Out-File -Append $cloudOrgpath

$opOrgtext = "===On-Premises Organization==="
$opOrgtext | Out-File $onPremOrgpath
$opOrg = Get-OnPremisesOrganization
$opOrg | FL | Out-File -Append $onPremOrgpath

$migEndtext = "===Migration Endpoints===:"
$migEndtext | Out-File $migEndpath
$migEndp = Get-MigrationEndpoint
$migEndp | FL | Out-File -Append $migEndpath

#OAuth Configs:
$iOrgText = "===IntraOrg Connector===:"
$iOrgText | Out-File $clOauthpath
$CliOrgConn = Get-IntraOrganizationConnector
if ($null -ne $CliOrgConn) {
    $CliOrgConn | FL | Out-File -Append $clOauthpath
    $CliOrgConn | ConvertTo-Json | Out-File $clIOConnjson
    #$CliOrgConn |Export-Clixml $clIntraOrgxmlpath

} else {
    $iOrgText2 = "***No OAuth Configs found***"
    $iOrgText2 | Out-File -Append $clOauthpath
}
Start-Sleep -Seconds 2

#Mail Flow Configs:
$outbCtext = "===O365 Outbound Connector Details===:"
$outbCtext | Out-File $outbConnpath
$o365OutConn = Get-OutboundConnector -IncludeTestModeConnectors $true
$o365OutConn | FL | Out-File -Append $outbConnpath
$o365OutConn | ConvertTo-Json | Out-File $outbConnjson 
$inbCtext = "===O365 Inbound Connector Details===:"
$inbCtext | Out-File $inbConnpath
$o365InConn = Get-InboundConnector 
$o365InConn |FL | Out-File -Append $inbConnpath
$o365InConn | ConvertTo-Json | Out-File $inbConnjson

$accDtext = "===Accepted Domain===:"
$accDtext | Out-File $accDompath
$accDomain = Get-AcceptedDomain $domain
$accDomain | FL | Out-File -Append $accDompath
$accDomain | ConvertTo-Json | Out-File $exoacceptedDomjson

$remDexotext = '===EXO Remote Domains==='
$remDexotext | Out-File $remDomEXOPath
$remDomEXO = Get-RemoteDomain
$remDomEXO | FL | Out-File -Append $remDomEXOPath
$remDomEXO | ConvertTo-Json | Out-File $exoRemdomjson

$addpoltext = "===EXO Email Address Policies===:"
$addpoltext | Out-File $ExOaddpoltxt
$addpolexo = Get-EmailAddressPolicy
$addpolexo | FL | Out-File -Append $ExOaddpoltxt

#Close Connection to EXO:
Write-Host -ForegroundColor White "Closing connection to Exchange Online..."
Get-PSSession | Remove-PSSession
}

#Prompt for domain name:
Write-Host -ForegroundColor Yellow "Enter your primary domain name:"
$domain = Read-Host

#EXO Collection Prompt:
$ans = EXO-RemoteQ

#Goodbye:
Write-Host -ForegroundColor White "Collection complete. Review or submit any files located in '$outputDir' to Microsoft."

#Revert execution policy if needed:
if ($execPol -ne 'Unrestricted'){
    Write-Host -ForegroundColor Cyan "Changing execution policy back to '$execPol'..."
    Set-ExecutionPolicy $execPol -Force
}
