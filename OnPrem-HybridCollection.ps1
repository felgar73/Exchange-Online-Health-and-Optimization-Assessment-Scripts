#Hybrid Collection Script for On-Premises Exchange Configs:

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
$outputDir = $rootFolder +'\HybridConfigs' + '_' + $date
$onPremDir = $outputDir + '\OnPremises'
#$cloudDir = $outputDir + '\Office365'
$errorLog = $outputDir + '\ErrorLog.log'

#OnPrem output paths:
$hybConfigjson = $onPremDir + '\Get-HybridConfig.json'
$hybtxtPath = $onPremDir + '\Get-HybridConfig.txt'
$sendConnJson = $onPremDir + '\SendConnectors.json'
$sendConnTxtPath = $onPremDir + '\Sendconnector.txt'
$recConnjson = $onPremDir + '\ReceiveConnectors.json'
$recConnTxtPath = $onPremDir + '\ReceiveConnectors.txt'
$exchCertjson = $onPremDir + '\Get-ExchangeCertificate.json'
$exchCertpath = $onPremDir + '\Exchange-Certificates.txt'
$opsharePolicyjson = $onPremDir + '\Get-SharingPolicy.json'
$opsharPath = $onPremDir + '\Sharing-Configs_OnPrem.txt'
$opfedorgIdjson = $onPremDir + '\Get-FederatedOrganizationIdentifier.json'
$opfederatConfpath = $onPremDir + '\Federation-Configs_OnPrem.txt'
$opfedTrustjson = $onPremDir + '\Get-FederationTrust.json'
$ewsVdirjson = $onPremDir + '\Get-WebServicesVirtualDirectory.json'
$ewstxtPath = $onPremDir + '\EWS-TxtOutput.txt'
$opOauthConfigjson = $onPremDir + '\OAuthConfig.json'
$opOauthPath = $onPremDir + '\OAuth-Configs_OnPrem.txt'
$authSvrjson = $onPremDir + '\Get-AuthServer.json'
$hcwLogsOPPath = $onPremDir + '\HCW-LogCmds_OnPrem.txt'
$opRemDomtxt = $onPremDir + '\RemoteDomains_OnPrem.txt'
$opRemdomjson = $onPremDir + '\Get-RemoteDomain.json'
$authConfigpath = $onPremDir + '\Get-AuthConfig_OnPrem.txt'
$OPaddpoltxt = $onPremDir + '\EmailAddressPolicy_OnPrem.txt'
$opAddpoljson = $onPremDir + '\Get-EmailAddressPolicy.json'
$opOrgConfigtxt = $onPremDir + '\OrganizationConfig-OnPrem.txt'
$partnerAppjson = $onPremDir + '\Get-PartnerApplication.json'
$skypeIntTxt = $onPremDir + '\SkypeIntegration-Configs.txt'
$opIOConnjson = $onPremDir + '\Get-IntraOrganizationConnector.json'
$opOrgReljson =  $onPremDir + '\Get-OrganizationRelationship.json'
$OpacceptedDomJson = $onPremDir + '\Get-AcceptedDomain.json'
$allserversJson = $onPremDir + '\Get-AllServers.json'

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
#OnPrem folder creation/validation:
function OnPremDir-Create {
    Write-Host -ForegroundColor Cyan "Creating collection folder..."
    if (!(Test-Path $onPremDir)) {
     New-Item -itemtype Directory -Path $onPremDir 
    } 
}

#Collect On-Prem Yes-No Function:
function OnPrem-CollectQ{
Write-Host -ForegroundColor Yellow "Ready to collect on-premises data? Y/N:"
  $ans = Read-Host
    if ((!$ans) -or ($ans -eq 'y') -or ($ans -eq 'yes')){
        $ans = 'yes'
    #Remote PS Connection Prompt:
    OnPrem-RemoteQ
    #Enter Hybrid server names:
    Write-Host -ForegroundColor Yellow "Enter your Hybrid server names separated by commas (Ex: server1,server2):"
    $script:hybsvrs = (Read-Host).split(",") | foreach {$_.trim()}
    #$hybsvrs | ConvertTo-Json | Out-File $allserversJson

    #Check/Create output folder:
    OnPremDir-Create
    Write-Host -ForegroundColor Cyan "Checking for HCW logs..."
    HCWLogs-OnPrem
    #Collect data:
    OnPrem-Collection
    } else {
        $ans = 'no'
        Write-Host -ForegroundColor Cyan "Skipping on-premises data collection..."
    }
}
#Remote On-Prem Exchange Functions:
function OnPrem-RemoteQ {
    Write-Host -ForegroundColor Yellow "Do you need to create a remote connection to Exchange On-Premises? Y/N:"
  $ans = Read-Host
    if ((!$ans) -or ($ans -eq 'y') -or ($ans -eq 'yes')){
        $ans = 'yes'
        Remote-ExchOnPrem
    } else {Write-Host -ForegroundColor Cyan "Skipping remote powershell connection to on-premises..."}
}
function Remote-ExchOnPrem {
    Write-Host -ForegroundColor Yellow "Enter your On-Premises Exchange server FQDN:"
    $fqdn = Read-Host 
    $opCreds = Get-Credential -Message "Enter your Exchange admin credentials:" -UserName $env:USERDOMAIN\$env:USERNAME
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$fqdn/powershell/ -Credential $opCreds -Authentication Kerberos
       try {
        Import-PSSession $Session -DisableNameChecking -AllowClobber
       }
       catch {
           Write-Host -ForegroundColor Red "Failed to create remote session, please try again..."
           exit
       }
}
#Exch Certs Function:
function ExchCert-Collection {
    $script:exchCerts = $hybsvrs | foreach {Get-ExchangeCertificate -Server $_}
    $exchCerts | FL | Out-File $exchCertpath
    $exchCerts | ConvertTo-Json | Out-File $exchCertjson
}
#EWS VDir Collection:
function EWS-VdirCollect {
    $script:ewsVdir = $hybsvrs | foreach {Get-WebServicesVirtualDirectory -Server $_ -ADPropertiesOnly}
    $ewsVdir | FL | Out-File $ewstxtPath
    $ewsVdir | ConvertTo-Json | Out-File $ewsVdirjson
}

#On-Prem Data collection:
function OnPrem-Collection {
    Write-Host -ForegroundColor Cyan "Collecting Hybrid configuration details, please wait..."
    #Expand PS output:
    $fenlimit = $FormatEnumerationLimit
    if ($fenlimit -ne '-1'){
        $FormatEnumerationLimit=-1
    }
    #On-Prem Configs Collection:
    $hybtext = "===Hybrid Servers Entered===:"
    $hybtext | Out-File $hybtxtPath
    $hybsvrs | Out-File -Append $hybtxtPath
    $hybsvrs | ConvertTo-Json | Out-File $allserversJson
    Add-Content $hybtxtPath -Value "===Hybrid Configuration===:"
    $hybConf = Get-HybridConfiguration
    $hybConf | FL | Out-File -Append $hybtxtPath
    $hybConf | ConvertTo-Json | Out-File $hybConfigjson
    Start-Sleep -Seconds 2
    
    $OrgConfigtitle = "===On-Premises Organization Config Details===:"
    $OrgConfigtitle | Out-File $opOrgConfigtxt
    $orgConfig = Get-OrganizationConfig
    $orgConfig |FL | Out-File -Append $opOrgConfigtxt
    Start-Sleep -Seconds 2

    $shpol =  "===Sharing Policy Details===:" 
    $shpol | Out-File $opsharPath
    $sharePol = Get-SharingPolicy 
    $sharePol | FL | Out-File -Append $opsharPath
    $sharePol | ConvertTo-Json | Out-File $opsharePolicyjson
    Start-Sleep -Seconds 2
    Add-Content $opsharPath -Value "===Org Relationship Details===:"
    $orgRel = Get-OrganizationRelationship 
    $orgRel |FL | Out-File -Append $opsharPath
    $orgRel | ConvertTo-Json | Out-File $opOrgReljson
    Start-Sleep -Seconds 2

    #Federation Config Info:
    $fedinfotext = "===Federated Organization Identifier===:"
    $fedinfotext | Out-File $opfederatConfpath
    $fedIdent = Get-FederatedOrganizationIdentifier -IncludeExtendedDomainInfo: $false
    $fedIdent | FL | Out-File -Append $opfederatConfpath
    $fedIdent | ConvertTo-Json | Out-File $opfedorgIdjson
    Add-Content $opfederatConfpath -Value "===Federation Information===:"
    $fedInfo = Get-FederationInformation -DomainName $domain -ErrorAction SilentlyContinue
    $fedInfo |FL | Out-File -Append $opfederatConfpath
    Add-Content $opfederatConfpath -Value "===Federation Trust Info===:"
    $fedtrust = Get-FederationTrust
    $fedtrust | ConvertTo-Json | Out-File $opfedTrustjson
    $fedtrust | FL Name,Org*certificate,TokenIssuerUri,TokenIssuerEpr,WebRequestorRedirectEpr | Out-File -Append $opfederatConfpath
    $fedtrusttxt = "For additional Federation Trust details, see 'Get-FederationTrust.json' file."
    Add-Content $opfederatConfpath -Value $fedtrusttxt
    Start-Sleep -Seconds 2
    
    #Exch Certs:
    ExchCert-Collection

    #Mail Flow:
    $sendTitle = "===Send Connector Details===:"
    $sendtitle | Out-File $sendConnTxtPath
    $sendConn = Get-SendConnector |? {$_.AddressSpaces -like '*onmicrosoft.com*'}
    $sendConn | ConvertTo-Json | Out-File $sendConnjson
    #$sendConn | Export-Clixml $sendConnxmlPath
    $sendConn | FL | Out-File -Append $sendConnTxtPath
    Start-Sleep -Seconds 2
    $recTitle = "===Receive Connector Details===:"
    $recTitle  | Out-File $recConnTxtPath
    $recvConn = Get-ReceiveConnector |?{$_.TlsDomainCapabilities -like '*outlook*'}
    #$recvConn | Export-Clixml $recConnxmlPath
    $recvConn | ConvertTo-Json | Out-File $recConnjson
    $recvConn |FL | Out-File -Append $recConnTxtPath
    $accDomain = Get-AcceptedDomain
    $accDomain | ConvertTo-Json | Out-File $OpacceptedDomJson
    #Remote Domains:
    $remText = "===Remote Domains===:"
    $remText | Out-File $opRemDomtxt
    $remDom = Get-RemoteDomain
    $remDom | FL | Out-File -Append $opRemDomtxt
    $remDom | ConvertTo-Json | Out-File $opremDomjson
    #Email Address Policies:
    $addpoltext = "===On-Premises Email Address Policies===:"
    $addpoltext | Out-File $OPaddpoltxt
    $addpolOP = Get-EmailAddressPolicy
    $addpolOP | FL | Out-File -Append $OPaddpoltxt
    $addpolOP | convertto-json | Out-File $opAddpoljson
    Start-Sleep -Seconds 2
    
    #EWS VDir Collect function:
    EWS-VdirCollect
    
    #OAuth Config Details:
    $iOrgConn = Get-IntraOrganizationConnector
    if (!$iOrgConn){
        $iocFailtext = "No IntraOrg Connector detected, OAuth may not be configured..."
        Write-Host -ForegroundColor Cyan $iocFailtext
        }
    $iorgtext = "===IntraOrg Connector===:"
    $iorgtext | Out-File $opOauthPath
    if ($iocFailtext -ne $null) {
        $iocFailtext | Out-File -Append $opOauthPath
    }
    $iOrgConn | FL | Out-File -Append $opOauthPath
    Add-Content $OPoauthPath -Value "===IntraOrganization Configs===:"
    $iOrgConf = Get-IntraOrganizationConfiguration -WarningAction:SilentlyContinue
    $iOrgConf | FL | Out-File -Append $OPoauthPath
    $iOrgConf | ConvertTo-Json | Out-File $opIOConnjson
    Add-Content $OPoauthPath -Value "===Partner Application Details===:"
    $ptnrapp = Get-PartnerApplication
    $ptnrapp | ConvertTo-Json | Out-File $partnerAppjson 
    $ptnrapp |FL Name,Enabled,Applicationidentifier,UseAuthServer,LinkedAccount| Out-File -Append $OPoauthPath
    Add-Content $opOauthPath -Value "===Auth Server Settings===:"
    $authsvr = Get-AuthServer
    $authsvr | ConvertTo-Json | Out-File $authSvrjson
    $authsvr | FL Name,type,realm,enabled,Domainname,ApplicationIdentifier,TokenIssuingEndpoint,AuthorizationEndpoint,IsDefaultAuthorizationEndpoint | Out-File -Append $opOauthPath
    Add-Content $OPoauthPath -Value "**Additional Auth Server details found in Json file."
    $authConftext = "===On-Premises Auth Config===:"
    $authConftext | Out-File $authConfigpath
    $authConf = Get-AuthConfig
    $authConf | Out-File -Append $authConfigpath
     
    #Test OAuth Config: 
    function OAuth-Test-OP {
    Write-Host -ForegroundColor Yellow "Would you like to test OAuth on-prem? Y/N:"
        $ans = Read-Host
        if ((!$ans) -or ($ans -eq 'y') -or ($ans -eq 'yes')){
            $ans = 'yes'
            Test-OAuthConnectivity -Service EWS -TargetUri https://outlook.office365.com/ews/exchange.asmx -Mailbox $testMbx -Verbose 
        } else {$ans = 'no'
        Write-Host -ForegroundColor White "Skipping OAuth Test..."
            }
    }    

    #Skype Integration Details:
    $skypText = "===Skype On-prem Integration Details===:"
    $skypText | Out-File $skypeIntTxt
    $sfbUser = Get-MailUser Sfb* -ErrorAction SilentlyContinue
    $sfbUser | FL | Out-File -Append $skypeIntTxt
    $userAppRole = Get-ManagementRoleAssignment -Role UserApplication -GetEffectiveUsers |? {$_.EffectiveUserName -like 'Exchange*'}
    $archiveAppRole = Get-ManagementRoleAssignment -Role ArchiveApplication -GetEffectiveUsers |? {$_.EffectiveUserName -like 'Exchange*'}
    $userAppRole |FL Role, *User*, WhenCreated | Out-File -Append $skypeIntTxt
    $archiveAppRole |FL Role, *User*, WhenCreated | Out-File -Append $skypeIntTxt

    #Close remote connection:
    Write-Host -ForegroundColor White "Closing connection to Exchange..."
    Get-PSSession | Remove-PSSession
    Start-Sleep -Seconds 2
    }

#Prompt for domain name:
Write-Host -ForegroundColor Yellow "Enter your primary domain name:"
$domain = Read-Host

#On-prem Collection Prompt:
$ans = OnPrem-CollectQ

#Goodbye:
Write-Host -ForegroundColor White "Collection complete. Review or submit any files located in '$outputDir' to Microsoft."

#Revert execution policy if needed:
if ($execPol -ne 'Unrestricted'){
    Write-Host -ForegroundColor Cyan "Changing execution policy back to '$execPol'..."
    Set-ExecutionPolicy $execPol -Force
}
