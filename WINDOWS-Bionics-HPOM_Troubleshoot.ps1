<#
.SYNOPSIS
    Check various HPOM server related configurations and if a fix is required then perform a fix operation.

.DESCRIPTION
    This script will perform various tests on HPOM Server to check if all the configurations are correct and if any error is found then make a note of it in the log
    The User can then verify the log and make changes to the server by opting the Fix option accordingly

.PARAMETER -Check
    Checks a set of predefined set of configurations
.PARAMETER -Fix
    Performs fix operation if any error is found
.PARAMETER -HostOMIPs
    During the Fix , If an error is found this option will configure the new sets of HOST OM IPs. 
.PARAMETER -HostOMNames
    During the Fix , If an error is found this option will configure the new sets of HOST OM Names - FQDN. 
.PARAMETER -PrimaryOMHost
    If an error is found this option will configure the new primary OM host.
.PARAMETER -MgrID
    This option is used to set the manager ID
.PARAMETER -UpdateHostFile
    By Default this option is False
    If the user wishes to update the hostfile then the user has to set this as true

.INPUTS    
    -Check
    -Fix
    -Fix -HostOMIPs <IP Address> -HostOMNames <Server FQDN> -PrimaryOMHost <Server FQDN> -MgrID <Manager ID> -UpdateHostFile True
    -Fix -HostOMIPs <IP Address> -HostOMNames <Server FQDN> -PrimaryOMHost <Server FQDN> -MgrID <Manager ID> -UpdateHostFile False
    -Fix -HostOMIPs <IP Address> -HostOMNames <Server FQDN> -PrimaryOMHost <Server FQDN> -MgrID <Manager ID>

.OUTPUTS
    <Output of the script>
    Please check the log file for more details - C:\temp\WINDOWS-Bionics_<Servername>_HPOM_Troubleshoot_yyyymmdd.log

.NOTES
  Script         : WINDOWS-Bionics-HPOM_Troubleshoot.ps1
  Author         : Arjun Jagdish (arjun.jagdish@dxc.com)
  Requirements   : Powershell v4.0
  Creation Date  : 23-October-2018
  History        : version 1.0 23-October-2018 - Initial script release

.EXAMPLE
    .\WINDOWS-Bionics-HPOM_Troubleshoot.ps1 -Check
    .\WINDOWS-Bionics-HPOM_Troubleshoot.ps1 -Fix
    .\WINDOWS-Bionics-HPOM_Troubleshoot.ps1 -Fix -HostOMIPs "123.456.789.0" -HostOMNames "abcdef.ghi" -PrimaryOMHost "abcdef.ghi" -MgrID "123-ABC-456-DEF" -UpdateHostFile True
    .\WINDOWS-Bionics-HPOM_Troubleshoot.ps1 -Fix -HostOMIPs "123.456.789.0" -HostOMNames "abcdef.ghi" -PrimaryOMHost "abcdef.ghi" -MgrID "123-ABC-456-DEF" -UpdateHostFile False
    .\WINDOWS-Bionics-HPOM_Troubleshoot.ps1 -Fix -HostOMIPs "123.456.789.0" -HostOMNames "abcdef.ghi" -PrimaryOMHost "abcdef.ghi" -MgrID "123-ABC-456-DEF"
#>

Param ( 
 [Parameter( 
     ParameterSetName = 'Check', 
     Mandatory = $False, 
     HelpMessage = 'The operation for the script to perform')] 
 [switch]$Check = $True, 
  
 [Parameter( 
     ParameterSetName = 'HPOMFix', 
     Mandatory = $False, 
     HelpMessage = 'The operation for the script to perform' 
 )] 
 [switch]$Fix = $False, 
  
 [Parameter( 
     ParameterSetName = 'HPOMFix', 
     Mandatory = $False, 
     HelpMessage = 'Mention the HPOM Host IPs')] 
 [string] $HostOMIPs, 
  
 [Parameter( 
     ParameterSetName = 'HPOMFix', 
     Mandatory = $False, 
     HelpMessage = 'Mention the HPOM Host Names')] 
  
 [string]$HostOMNames, 
  
 [Parameter( 
     ParameterSetName = 'HPOMFix', 
     Mandatory = $False, 
     HelpMessage = 'Mention the primary HPOM Host server')] 
 [string]$PrimaryOMHost, 
  
 [Parameter( 
     ParameterSetName = 'HPOMFix', 
     Mandatory = $False, 
     HelpMessage = 'Mention the Manager ID')] 
 [string]$MgrID, 
  
 [Parameter( 
     ParameterSetName = 'HPOMFix', 
     Mandatory = $False, 
     HelpMessage = 'Update HostFile?')] 
     [ValidateSet('True','False')] 
 [string[]]$UpdateHostFile = $False 
  
 ) 
  
 $moduleList = @() 
  
 #----------------------------------------------------------[Declarations]---------------------------------------------------------- 
  
 $sScriptVersion                    = "1.0" 
 $MinPSVer                          = 4 
 $Reportdate                        = Get-Date -format yyyyMMdd 
 $ScriptDesc                        = "HPOM_Troubleshoot" 
 $logdir                            = "C:\temp" 
 $logname                           = "WINDOWS-Bionics_"+$env:COMPUTERNAME+"_${ScriptDesc}_${Reportdate}.log" 
 $logfile                           = Join-Path $logdir $logname 
 $AllowedWinVersions                = @("Windows Server 2008 R2","Windows Server 2012 R2","Windows Server 2016") 
  
 $HPOM_MANAGER_DATA                 = @() 
 $Ovconfget_mngrarg_seccoreauth     = "sec.core.auth" 
 $Global:Errors                     = @() 
 $ovHome                            = "C:\Program Files\HP OpenView\bin\win64" 
 $global:ovconfget_manager          = "$ovHome\ovconfget.exe" 
 $global:ovconfget_manager_arg      = "sec.core.auth MANAGER" 
 $global:ovconfget_managerID_arg    = "sec.core.auth MANAGER_ID" 
 $global:opcagt                     = "$ovHome\opcagt.bat" 
 $global:certcheck                  = "$ovHome\ovcert.exe" 
 $global:certcheck_arg              = "-status" 
 $Global:OVMgrIDarg                 = "-ns sec.core.auth -set MANAGER_ID $Global:MgrID" 
 $Global:OMService                  = "OVCTRL" 
 $Global:Python                     = "C:\Program Files\Opsware\agent\lcpython15\python.exe" 
 $Global:CustAttrib                 = "C:\Program Files\Opsware\agent_tools\get_cust_attr.bat" 
 $Global:HPOM_MANAGER               = "HPOM_MANAGER" 
 $Global:CScript                    = "c:\windows\system32\cscript.exe" 
 $Global:AuthMgrExec                = '"C:\Program Files\HP OpenView\bin\win64\OpC\install\opcactivate.vbs"' + " -srv $Global:PrimaryOMHost -cert_srv $Global:PrimaryOMHost -force_config_mode" 
 $Global:CertReqArg                 = "-certreq" 
 $global:OABuffering                = $Null 
 $Global:PolicyCmd                  = "$ovHome\ovpolicy" 
 $Global:PolicyList                 = "-list" 
 $Global:SubCount                   = 0 
 $Global:OABuffering                = $null 
 $BuffCount                         = 0 
 $BufferLoop                        = 0 
 $sleeptime                         = 45 
  
 $global:Err_Flag                   = $false 
  
 $Global:Error0                     = $False 
 $Global:Error1                     = $Null 
 $Global:Error2                     = $Null 
 $Global:Error3                     = $Null 
 $Global:Error4                     = $Null 
 $Global:Error5                     = $Null 
  
 #-------------------------------------------------------------------------------------------------------------------- 
  
  
 Function Write-Log              { 
     Param 
     ( 
         $Text, 
         [Int]$Flag 
     ) 
  
     $Date = Get-Date -Format yyyy-MM-dd-HH:mm:ss 
  
     If ($OutputOnly -eq $false) { 
         Switch($Flag) 
         { 
             0 { Write-Output "$Date INFO: $Text" | Tee-Object -FilePath $LogFile -Append } 
             1 { 
                 $DefaultColor = $host.UI.RawUI.ForegroundColor 
                 $host.UI.RawUI.ForegroundColor = "Yellow" 
                 Write-Output "$Date WARN: $Text" | Tee-Object -FilePath $LogFile -Append 
                 $host.UI.RawUI.ForegroundColor = $DefaultColor 
             } 
             2 { 
                 $DefaultColor = $host.UI.RawUI.ForegroundColor 
                 $host.UI.RawUI.ForegroundColor = "Red" 
                 Write-Output "$Date ERRO: $Text" | Tee-Object -FilePath $LogFile -Append 
                 $host.UI.RawUI.ForegroundColor = $DefaultColor 
              } 
             3 { Write-Output "$Date DEBU: $Text" | Tee-Object -FilePath $LogFile -Append } 
         } 
     } 
     else { 
         Switch($Flag) 
         { 
             0 { Write-Output "$Date INFO: $Text" | Out-File $LogFile -append } 
             1 { Write-Output "$Date WARN: $Text" | Out-File $LogFile -append } 
             2 { Write-Output "$Date ERRO: $Text" | Out-File $LogFile -append } 
             3 { Write-Output "$Date DEBU: $Text" | Out-File $LogFile -append } 
         } 
     } 
 } 
 Function LoadModules         () { 
     if($moduleList -gt 0){ 
         Write-Log "Searching for module components..." 0 
         $loaded = Get-Module -Name $moduleList -ErrorAction SilentlyContinue | ForEach-Object {$_.Name} 
         $registered = Get-Module -Name $moduleList -ListAvailable -ErrorAction SilentlyContinue | ForEach-Object {$_.Name} 
  
         foreach ($module in $registered) { 
             if ($loaded -notcontains $module) { 
                 Write-Log "Loading module $module" 0 
                 Try { 
                     Import-Module $module 
                 } 
                 Catch { 
                     Write-Log "Error in importing module $module - $_.Exception.Message" 2 
                     Exit 1 
                 } 
             } 
         } 
    } 
    else { 
         Write-Log "No modules to load. Proceeding to main script execution" 0 
     } 
 } 
 Function UnLoadModules       () { 
     if($moduleList -gt 0){ 
         Write-Log "Searching for imported module components..." 0 
  
         $loaded = Get-Module -Name $moduleList -ErrorAction SilentlyContinue | ForEach-Object {$_.Name} 
         $registered = Get-Module -Name $moduleList -ListAvailable -ErrorAction SilentlyContinue | ForEach-Object {$_.Name} 
  
         foreach ($module in $registered) { 
             if ($loaded -contains $module) { 
                 Write-Log "Removing module $module" 0 
                 Try { 
                     Remove-Module $module 
                 } 
                 Catch { 
                     Write-Log "Error in Removing module $module - $_.Exception.Message" 2 
                 } 
             } 
         } 
     } 
     else { 
         Write-Log "No modules to load. Proceeding to main script execution" 0 
     } 
 } 
 Function ValidatePrereq      () {  
 Try {  
         Write-Log "Validating PowerShell version on the server" 0  
         $PSVer = $psversiontable.PSVersion.Major  
         if ($PSVer -lt $MinPSVer)  
         {  
             Write-Log "Unsupported PowerShell Version found. Please install PowerShell $MinPSVer or higher and re-run the script." 2  
             $global:Err_Flag = $true  
             Return  
         }  
          
         Write-Log "Validating Windows Server Version" 0  
         $NewNodeInfo = Get-WmiObject Win32_OperatingSystem -ComputerName $env:COMPUTERNAME | Select-Object Caption -ErrorAction Stop  
         $NewNodeOS = $NewNodeInfo.Caption  
         $AllowedWinVersions | %{if($NewNodeOS.ToUpper() -match $_.ToString().ToUpper()){$global:WinVerFlag = $true}} 
         if(!($global:WinVerFlag)) 
         {  
             Write-Log "Server must be Windows Server 2008R2 or 2012R2 or newer" 2  
             $global:Err_Flag = $true  
             Return  
         } 
  
         Write-Log "Validating if tools being installed " 0 
         if(!(Test-path $Global:CustAttrib)) 
         { 
             Write-Log "Opsware / agent tools needs to be installed" 2  
             $global:Err_Flag = $true  
             Return  
         } 
  
                  
     }  
  
 Catch {  
     Write-Log "Error in Validating Prerequistes and Inputs values - $_.Exception.Message" 2  
     $global:Err_Flag = $true  
     }  
 } 
  
 ########## Check Functions ########## 
 Function CheckOAInstall      () { 
     Try {$Result_ServiceInstall = Get-Service -Name $Global:OMService -ErrorAction SilentlyContinue} 
     Catch {Write-Log "$Global:OMService is Not Installed" 2 
         exit 1 
         } 
     If ($Result_ServiceInstall) { 
         Write-Log "$Global:OMService is Installed" 0 
         } 
     Else { 
         Write-Log "$Global:OMService is Not Installed" 2 
         exit 1 
     } 
 } 
 Function CheckErrors         () { 
 If (!$Global:Errors) { 
     Write-Log "No Errors Found" 0 
      
 } 
 Else { 
     Write-Log "[ERROR(s) FOUND:] $Global:Errors" 2 
  
     } 
 } 
 Function ConfigurationChecks () { 
 Write-Log "Starting Configuration Checks" 0 
     CheckOMHost 
     CheckHostsFile 
     CheckOMService 
     CheckOVMgr 
     CheckOVMgrID 
     CheckCertStatus 
     CheckPolicyStatus 
 } 
 Function CheckOMHost         () { 
     Foreach ($IP in $Global:HostOMIPs) { 
         If (Test-NetConnection $IP -Port 383 -InformationLevel Quiet) { 
                 $Global:OMIP = $IP 
                 $Global:OMIP = $Global:OMIP.trim() 
                 Write-Log "Found OM IP @ $Global:OMIP" 0 
                 Break 
         } 
          
         else { 
             Write-Log "No IP Cound be Contacted @ $IP" 1 
         } 
     } 
     If (!$Global:OMIP) { 
         $Global:Error0 = "True" 
         $Global:Errors = $Global:Errors += "Error Cannot find IP port 383 `n" 
         } 
 } 
 Function CheckHostsFile      () { 
     $ResultHostsFile = Select-String -Path "C:\Windows\System32\drivers\etc\hosts" -pattern $Global:PrimaryOMHost 
         If ($ResultHostsFile -like "*$Global:OMIP*") { 
             $Global:Error1 = $False 
             Write-Log "Hosts File is Correct $Global:PrimaryOMHost" 0 
         } 
         Else { 
             $Global:Error0 = "True" 
             $Global:Error1 = "True" 
             $Global:Errors = $Global:Errors += "Error in Hosts File `n" 
             Write-Log "Hosts File is Incorrect" 2 
         } 
 } 
 Function CheckBuffering      () { 
     StartProcess $global:opcagt 
     $Result_Buffer = $Global:Result_StartProcess 
     If ($Result_Buffer -like "*Message Agent is not buffering.*") { 
         $global:OABuffering = $False 
         Write-Log "OA is NOT Buffering" 0 
     } 
     Else { 
         $Global:Error0 = $True 
         $global:OABuffering = $True 
         $Global:Errors = $Global:Errors += "Agent is Buffering `n" 
         Write-Log "OA is Buffering" 2 
     } 
 } 
 Function CheckOMService      () { 
     $Result_Service = Get-Service -Name $Global:OMService | Select -ExpandProperty Status 
     If ($Result_Service -eq "Running") { 
         Write-Log "$Global:OMService is $Result_Service" 0 
         } 
     Else { 
         $Global:Error0 = $True 
         $Global:Error3 = $True 
         $Global:Errors = $Global:Errors += "$Global:OMService is $Result_Service `n" 
         Write-Log "$Global:OMService is $Result_Service" 2 
     } 
 } 
 Function CheckOVMgr          () { 
     StartProcess $global:ovconfget_manager $global:ovconfget_manager_arg 
     $Result_Mgr = $Global:Result_StartProcess 
     #echo "---->$Result_Mgr<----" 
     #echo "---->$Global:PrimaryOMHost<----" 
     If (($Result_Mgr -eq $Global:PrimaryOMHost) -or ($Result_Mgr -eq $Global:OMIP)) { 
        Write-Log "sec.core.auth MANAGER is Correct" 0 
     } 
     Else { 
         $Global:Error0 = $True 
         $Global:Error2 = $True 
         $Global:Errors = $Global:Errors += "sec.core.auth MANAGER is Incorrect `n" 
         Write-Log "sec.core.auth MANAGER is Incorrect" 2 
     } 
 } 
 Function CheckOVMgrID        () { 
     StartProcess $global:ovconfget_manager $global:ovconfget_managerID_arg 
     $Result_MgrID = $Global:Result_StartProcess 
     If ($Result_MgrID -eq $Global:MgrID) { 
        Write-Log "sec.core.auth MANAGER_ID Correct" 0 
     } 
     Else { 
         $Global:Error0 = $True 
         $Global:Error6 = $True 
         $Global:Errors = $Global:Errors += "sec.core.auth MANAGER_ID Incorrect `n" 
         Write-Log "sec.core.auth MANAGER_ID Incorrect" 2 
     } 
 } 
 Function CheckCertStatus     () { 
     StartProcess $global:certcheck $global:certcheck_arg 
     $Result_CertStatus = $Global:Result_StartProcess 
     If ($Result_CertStatus -like "Status: Certificate is installed.") { 
         Write-Log "$Result_CertStatus" 0 
     } 
     Else { 
         $Global:Error0 = $True 
         $Global:Error4 = $True 
         $Global:Errors = $Global:Errors += "$Result_CertStatus `n" 
         Write-Log $Result_CertStatus 1 
     } 
 } 
 Function CheckPolicyStatus   () { 
     StartProcess $Global:PolicyCmd $Global:PolicyList 
     $Result_Policy = $Global:Result_StartProcess 
     If ($Result_Policy -like "*INFO:    No policies are installed on host 'localhost'.*") { 
         $Global:Error0 = $True 
         $Global:Error5 = $True 
         $Global:Errors = $Global:Errors += "Endpoint is Missing Policies `n" 
         Write-Log "OA Missing Policies" 2 
     } 
     Else { 
         Write-Log "OA Policies Applied" 0 
     } 
 } 
 ##################################### 
  
 Function StartProcess        () { 
     $Cmd = $args[0] 
     $args = $args[1] 
     $Global:SubCount = $Global:SubCount + 1 
     #echo "args ---->$args<----" 
     Write-Log "***** Launching Sub Process #$Global:SubCount" 0 
     $ProcessInfo = New-Object System.Diagnostics.ProcessStartInfo 
     $ProcessInfo.Filename = $Cmd 
     $ProcessInfo.RedirectStandardError = $true 
     $ProcessInfo.RedirectStandardOutput = $true 
     $ProcessInfo.UseShellExecute = $false 
     $ProcessInfo.Arguments = $args 
     $Process = New-Object System.Diagnostics.Process 
     $Process.StartInfo = $ProcessInfo 
     $Process.Start() | Out-Null 
     $Process.WaitForExit() 
     $ResultOut = $Process.StandardOutput.ReadToEnd() 
     $ResultError = $Process.StandardError.ReadToEnd() 
     $ResultOut = $ResultOut.trim() 
     #echo "---->$ResultOut<----" 
     $Global:Result_StartProcess = $ResultOut 
 } 
  
 ########### Fix Functions ########### 
 Function FixHostFile         () { 
  
     Write-Log "Resolving Hosts File" 0 
     Copy-Item "C:\Windows\system32\drivers\etc\hosts" "C:\Windows\system32\drivers\etc\hosts.bak1" 
     #Get HPOM MANAGER Attributes 
     StartProcess $Global:CustAttrib $Global:HPOM_MANAGER 
     $Result_Host_Entries = $Global:Result_StartProcess 
     Add-Content -Path "C:\Windows\system32\drivers\etc\hosts" -Value $Result_Host_Entries 
     Write-Log "Hosts File Update Complete" 0 
     CheckHostsFile 
     If ($Global:Error1 -eq $True) { 
         Write-Log "Confirm HPSA Custom Attributes are Correct" 2 
         break 
     } 
 } 
 Function FixAuthMgr          () { 
     Write-Log "Resolving Auth MGR" 0 
     StartProcess $Global:CScript $Global:AuthMgrExec 
     CheckOVMgr 
 } 
 Function FixAuthMgrID        () { 
     Write-Log "Resolving Auth MGR ID" 0 
       StartProcess $global:ovconfget_manager $Global:OVMgrIDarg 
     CheckOVMgrID 
 } 
 Function FixOVService        () { 
     Write-Log "Resolving OV Service" 0 
     Stop-Service -Name $Global:OMService 
     Start-Service -Name $Global:OMService 
     CheckOMService 
 } 
 Function FixOVCert           () { 
     Write-Log "Resolving OV Cert" 0 
     StartProcess $global:certcheck $Global:CertReqArg 
     CheckCertStatus 
 } 
 ##################################### 
  
 Function Get_DefaultData        { 
  
     & $CustAttrib "HPOM_MANAGER" | %{

     <#
         $IP      = $_.tostring().replace("  ","*").replace(" ","").split("*")[0]
         $SN_FQDN = $_.tostring().replace("  ","*").replace(" ","").split("*")[1] 
         $SN      = $_.tostring().replace("  ","*").replace(" ","").split("*")[2] 
    #>
         $IP      = $_.tostring().replace(" ","*").split("*")[0]
         $SN_FQDN = $_.tostring().replace(" ","*").split("*")[1] 
         $SN      = $_.tostring().replace(" ","*").split("*")[2] 

         if($SN -eq $Null)
         {
            Write-Log "Couldnt read HPOM Server Name from Custom Attributes because of wrong format - Please check and update" 2 
            $global:Err_Flag = $true  
            Return 
         }
         else
         {



  
         $Result = new-object psobject -Property @{ 
             ServerIP   = $IP 
             ServerFQDN = $SN_FQDN 
             ServerName = $SN 
             } 
         $HPOM_MANAGER_DATA += $Result 
     } 
      
     $HPOM_MANAGER_DATA_HostIPs   = $HPOM_MANAGER_DATA.ServerIP 
     if($HPOM_MANAGER_DATA.ServerFQDN) 
     { 
         $Global:HPOM_MANAGER_DATA_HostNames = $HPOM_MANAGER_DATA.ServerFQDN 
     } 
     else 
     { 
         $Global:HPOM_MANAGER_DATA_HostNames = $HPOM_MANAGER_DATA.ServerName 
     } 
     &$ovconfget_manager "sec.core.auth" | %{ 
      
         IF($_.TOSTRING().STARTSWITH("MANAGER_ID=")) 
         { 
             $Global:MANAGER_ID = $_.TOSTRING().SPLIT("=")[1].toupper() 
         } 
         elseif($_.TOSTRING().STARTSWITH("MANAGER=")) 
         { 
             $Global:MANAGERNAME = $_.TOSTRING().SPLIT("=")[1].toupper() 
         } 
     } 
  
     if($Global:HostOMIPs) 
     { 
         $Global:HostOMIPs | %{ 
         $_.split(",") | %{  
                 if(!($HPOM_MANAGER_DATA_HostIPs.contains($_))) 
                 { 
                     "Invalid HOST OM IP - Provide valid IP" 
                     $global:Err_Flag = $True 
                 } 
             } 
         }   
     } 
     else 
     { 
         $Global:HostOMIPs = $HPOM_MANAGER_DATA_HostIPs 
     } 
  
  
     if($Global:HostOMNames) 
     { 
         $Global:HostOMNames | %{ 
             $_.split(",") | %{  
                 if(!($Global:HPOM_MANAGER_DATA_HostNames.contains($_))) 
                 { 
                     "Invalid HOST SERVER NAMES - Provide valid FQDN" 
                     $global:Err_Flag = $True 
                 } 
             } 
         }   
     } 
     ELSE 
     { 
         $Global:HostOMNames = $Global:HPOM_MANAGER_DATA_HostNames 
     } 
  
     if($Global:PrimaryOMHost) 
     { 
         if($Global:PrimaryOMHost.ToUpper() -ne $Global:MANAGERNAME.ToUpper()) 
         { 
             "Invalid Primary Host Server" 
              $global:Err_Flag = $True 
         } 
     } 
     else 
     { 
         $Global:PrimaryOMHost = $Global:MANAGERNAME.ToUpper() 
     } 
  
     if($Global:MgrID) 
     { 
         if($Global:MgrID.ToUpper() -ne $Global:MANAGER_ID.ToUpper()) 
         { 
             "Invalid Manager ID" 
              $global:Err_Flag = $True 
         } 
     } 
     else 
     { 
          $Global:MgrID = $Global:MANAGER_ID.ToUpper() 
     } 
      
     Write-Log "HPOM_MANAGER Host: $($Global:HPOM_MANAGER_DATA_HostNames.ToUpper())" 0 
     Write-Log "MANAGER ID: $($Global:MANAGER_ID.ToUpper())" 0 
     Write-Log "MANAGER NAME: $($Global:MANAGERNAME.ToUpper())" 0 
     Write-Log "Host OM IPs: $($Global:HostOMIPs)" 0 
     Write-Log "Host OM Names: $($Global:HostOMNames.ToUpper())" 0 
     Write-Log "Primary OM Host: $($Global:PrimaryOMHost.ToUpper())" 0 
     }  
 } 
 Function CheckAll               { 
     Write-Log "Initiating Checking all settings" 0 
  
     #CheckOAInstall 
     try  { 
             Write-Log "Checking OA Install" 0  
             CheckOAInstall 
         } 
     catch{ 
             Write-Log "Failed Checking OA Install" 2 
             $global:Err_Flag = $true  
             Return 
         } 
      
     #ConfigurationChecks 
     try  { 
             Write-Log "Checking All Configurations" 0  
             ConfigurationChecks 
         } 
     catch{ 
             Write-Log "Failed Checking Configurations" 2 
             $global:Err_Flag = $true  
             Return 
         } 
      
     #CheckErrors 
     try  { 
             Write-Log "Checking if Errors Exist" 0  
             CheckErrors 
         } 
     catch{ 
             Write-Log "Failed Checking Errors" 2 
             $global:Err_Flag = $true  
             Return 
         }  
      
 } 
 Function FixAll                 { 
      If ($Global:Error0 = $True) { 
          
         Write-Log "Attempting to resolve errors." 0 
         If ($Global:Error1 -eq $True) { 
             if($Global:UpdateHostFile) 
             { 
                 try  { 
                     Write-Log "Attempting to Fix Host File." 0 
                     FixHostFile 
                 } 
                 Catch{ 
                     Write-Log "Failed to Fix Host File" 2 
                     $global:Err_Flag = $true  
                     Return                 
                 } 
             } 
             else 
             { 
                 Write-Log "HostFile Needs to be updated but is disabled as per default settings" 1 
             }             
         } 
         If ($Global:Error2 -eq $True) { 
             try  { 
                     Write-Log "Attempting to Fix Auth Manager." 0 
                     FixAuthMgr 
                 } 
             Catch{ 
                     Write-Log "Failed to Fix Auth Manager" 2 
                     $global:Err_Flag = $true  
                     Return                 
                 } 
         } 
         If ($Global:Error3 -eq $True) { 
             try  { 
                     Write-Log "Attempting to Fix OV Service." 0 
                     FixOVService 
              } 
             Catch{ 
                     Write-Log "Failed to Fix OV Service" 2 
                     $global:Err_Flag = $true  
                     Return                 
                 }         
         } 
         If ($Global:Error4 -eq $True) { 
             try  { 
                     Write-Log "Attempting to Fix OV Certificate." 0 
                     FixOVCert 
              } 
             Catch{ 
                     Write-Log "Failed to Fix OV Certificate" 2 
                     $global:Err_Flag = $true  
                     Return                 
                 }   
         } 
         If ($Global:Error6 -eq $True) { 
             try  { 
                     Write-Log "Attempting to Fix Auth Manager ID." 0 
                     FixAuthMgrID 
              } 
             Catch{ 
                     Write-Log "Failed to Fix Auth Manager ID" 2 
                     $global:Err_Flag = $true  
                     Return                 
                 } 
         }         
     } 
      
     # Final check to confirm not buffering 
     Do { 
                  
         Start-Sleep $sleeptime 
         Write-Log  "Sleeping $sleeptime Seconds" 0 
         $global:OABuffering = $null 
         $BuffCount = 0 
         $Global:Error0 = $false 
         $Global:Errors = $null 
         Write-Log "Testing Communications (buffering)" 0 
          
         try  { 
                     Write-Log "Checking for Buffering." 0 
                     CheckBuffering 
              } 
         Catch{ 
                 Write-Log "Failed to Check for Buffering" 2 
                 $global:Err_Flag = $true  
                 Return                 
             }  
  
         try  { 
                     Write-Log "Checking Configurations." 0 
                     ConfigurationChecks 
              } 
         Catch{ 
                 Write-Log "Failed to Check for Buffering" 2 
                 $global:Err_Flag = $true  
                 Return                 
             } 
  
         try  { 
                     Write-Log "Checking for Errors." 0 
                     CheckErrors 
              } 
         Catch{ 
                 Write-Log "Failed to Check for Buffering" 2 
                 $global:Err_Flag = $true  
                 Return                 
             }  
          
          
      
         If   ($global:OABuffering -eq "True") { 
             Write-Log "[ERROR:] Failed to Resolve Errors" 2 
              
         } 
         Else { 
             Write-Log  "Finished Resolving Errors" 0 
             Write-Log  "If this is a production server, you must access ESL and MtP this endpoint!" 1 
              
         } 
  
         $BuffCount ++ 
         $BufferLoop ++ 
     } 
  
     While (($BuffCount -lt 1) -and ($BufferLoop -lt 10))    
 } 
  
 Write-Log "Script Version - $sScriptVersion" 0 
 Write-Log "Beginning Script Execution" 0 
 Write-Log "Validating Prerequisites and Inputs" 0   
 ValidatePrereq 
 LoadModules 
  
 if($global:Err_Flag -eq $false) 
 { 
     try  {Get_DefaultData} 
     Catch{ 
         Write-Log "Failed to collect custom details - kindly configure " 2 
         $global:Err_Flag = $true  
         Return 
     } 
 } 
  
 if($global:Err_Flag -eq $false) 
 { 
     if($Check) { CheckAll } 
     if($Fix)   { FixAll } 
} 
  
 UnLoadModules 
  
 If ($global:Err_Flag -eq $false) { 
     Write-Log "Completed Script Execution - More details please check the log - $logfile" 0 
     If ($OutputOnly) { 
         $OutObj 
     } 
     Exit 0 
 } 
 Else { 
     Write-Log "Completed Script with Errors - More details please check the log - $logfile" 2 
     If ($OutputOnly) { 
         $OutObj 
         Write-Output "" 
         Write-Output "Completed Script with Errors - More details please check the log - $logfile" 
     } 
     Exit 1 
} 
