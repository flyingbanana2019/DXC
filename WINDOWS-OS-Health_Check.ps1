﻿<#

Name                          : hc.ps1
Brief Desc                    : Wintel Health Check Script
Created & Maintained by       : Abdullah, Effi-Azzari <effi-azzari.abdullah@hp.com>
Version                       : 3.0
Date Created                  : 08/09/2013
Last Updated				  : 31/03/2014

ToDo 
- 
	
Current Script status
- Totally rewrite the whole code from the original health check script
	> To have more customizable script which can add more checks if required in the future
	> Can customize the HTML table more efficiently
	
What's new?
- 08/09/2013
	> Rewrite the whole script starting with the HTML code especially for the table format using Adobe Dreamweaver
- 03/10/2013	
	> Rewrite the way to get the remote machine date & time locally
- 11/3/2014 : 9.23PM
	> Rewrite the way to get the last patch date & time using Get-Hotfix
			: 10.25PM
	> Rewrite for more efficient & the most accurate way to retrieve the local server time using wmiobject.LocalDateTime
- 12/3/2014 : 7.47AM
	> Reconstruct & customized the functions
- 13/3/2014
	> Rewrite the way it checks the Automatic services, better result & preview
	> Adding 2 more columns - UNC Path & RDP Checks
	> No more using external app tool, RDP connection check is now hard coded.
- 15/3/2014
	> Adding some details on Automatic Services - now we can services name which are stopped + those which are not successfully restart
- 17/3/2014
	> Reconstruct the CSS code & HTML table
- 20/3/2014
	> Bug fix for automatic services result which show negative calculation & also repeated variables.
	> 'Last boot' & 'uptime' will show proper result when there's no date captured.
- 26/3/2014
	> Fixed value which the vars should be reset to $null on every server.
	> Fixed results when server can be pinged but permission is denied - other checks continues
- 31/3/2014
	> Fix null value in server local time. More detailed errors calculated.
	> Local server time, volume C: & Automatic services - Include RPC Error handling
	
#>

Function WMIDateStringToDate($T) { 
    [System.Management.ManagementDateTimeconverter]::ToDateTime($T) 
} 

Function send_email { 
$FromAddress = "Health Check Status Report <No-Reply@hp.com>"
$ToAddress = "effi-azzari.abdullah@hp.com"
$CCAddress = "effi-azzari.abdullah@hp.com"
$MessageSubject = "Ericsson Health Check Status Report v$version"
#$MessageBody = $mailcontent
$SmtpServer = "SE-SMTP.ericsson.se"
#$Attachment = $file
Send-MailMessage -From $FromAddress -To $ToAddress -Cc $CCAddress -Subject $MessageSubject -BodyAsHtml $effi_report -SmtpServer $SmtpServer
} 

#Main Area
clear
$erroractionpreference = "SilentlyContinue"
$date = get-date
$version = 3.0
$Account = 'Ericsson'
$Filename = ".\HC_" + $date.Hour + $date.Minute + "_" + $Date.Day + "-" + $Date.Month + "-" + $Date.Year + ".htm"
$effi_report = @"
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />

<title>Server Health Check</title>

<style type="text/css">
table.t_data
{
    font-family: "Trebuchet MS", Arial, Helvetica, sans-serif;
	font-size: 13px;
	color: #FFF;
	background-color: #000;
	text-align: center;
    border-spacing: 1px;
    margin: 0 auto 0 auto;
}
table.t_data thead th, table.t_data thead td
{
    background-color: #9f9;
    font-family: "Trebuchet MS", Arial, Helvetica, sans-serif;
	font-size: 10px;
	color: #FFF;
	text-align: center;
    padding: 5px;
    margin: 1px;
}
table.t_data tbody th, table.t_data tbody td
{
    background-color: #fff;
    padding: 2px;
}

.HCHead2 {
	font-family: "Trebuchet MS", Arial, Helvetica, sans-serif;
	font-size: 11px;
	color: #FFF;
	background-color: #09F;
	text-align: center;
}

.HCInfo {
	font-family: "Trebuchet MS", Arial, Helvetica, sans-serif;
	font-size: 11px;
	color: #000000;
	text-align: center;
	background-color: #FFF;
}

.HCInfoRed {
	font-family: "Trebuchet MS", Arial, Helvetica, sans-serif;
	font-size: 11px;
	color: #000000;
	text-align: center;
	background-color: #F00;
}

.HCInfotextred {
	font-family: "Trebuchet MS", Arial, Helvetica, sans-serif;
	font-size: 11px;
	color: #FF0000;
	text-align: center;
	background-color: #FFF;
}
.HCInfotextgreen {
	font-family: "Trebuchet MS", Arial, Helvetica, sans-serif;
	font-size: 11px;
	color: #33CC66;
	text-align: center;
	background-color: #FFF;
}

.bgyellow { background-color: #FF0 }

#apDiv1 {
	position: absolute;
	left: 512px;
	top: 212px;
	width: 3px;
	height: 0px;
	z-index: 1;
}
#apDiv2 {
	position: absolute;
	left: 97px;
	top: 235px;
	width: 386px;
	height: 134px;
	z-index: 1;
}
</style>
</head>

<body>
<table class="t_data">
  <tr class="HCHead" style="padding:1px">
    <th colspan="11" scope="col">$Account Wintel Health Check
    </th>
  </tr>
  
  <tr class="HCHead2" style="padding:1px">
    <th scope="col">No</th>
    <th scope="col">Server</th>
    <th scope="col">ServerLocalTime</th>
    <th scope="col">Lastboot</th>
    <th scope="col">Uptime</th>
    <th scope="col">LastPatch</th>
    <th scope="col">Patch Date &amp; Time</th>
    <th scope="col">C: Threshold</th>
    <th scope="col">Automatic Services</th>
	<th scope="col">Test Path</th>
	<th scope="col">RDP Test</th>
  </tr>
"@
$servers = Get-Content servers.txt
$servercount = ($servers | Measure-Object).count
Foreach ($server in $servers) { 
	$Services = ''
	$count += 1
	Write-Progress -Activity "Initiating Server Health Check" -status "Processing ($count/$servercount): $server" -CurrentOperation "Checking server connections.." -PercentComplete ($count/$servercount * 100)
	#Populate Ping Results
	If (Test-Connection -ComputerName $server -Count 1 -TimeToLive 254 -Quiet){ $check = $true } else { $check = $false }
	
				#Results cleanup
				$localtime = ''
				$Lastboot = ''
				$Uptime = ''
				$LastPatch = ''
				$PatchDT = ''
				$VolC = ''
				$AutoServ = ''
				$TestPath = ''
				$RDPQ = ''
				$logicaldisk = ''
				$lh = ''
				
				#Important Variables cleanup
				$ltime = ''
				$computers = ''
				$lctime = ''
				$tPath = ''
				$hf = ''
				$volErr = ''
				$servErr = 0
				$servResult = 0
				$percentagefree = ''
				
				#Cleanup End
	
		switch ($check){
		$False {
			$check = ''
			
			$localtime = "N/A"
			$Lastboot = "N/A"
			$Uptime = "N/A"
			$LastPatch = "N/A"
			$PatchDT = "N/A"
			$VolC = "N/A"
			$AutoServ = "N/A"
			$TestPath = "N/A"
			$RDPQ = "N/A"
			
		    $effi_report += " 	 <tr class=HCInfoRed> "
			$effi_report += "    <td scope=col>$count</td> "
			$effi_report += "    <td scope=col>$server</td> "
			$effi_report += "    <td scope=col>$localtime</td> "
			$effi_report += "    <td scope=col>$Lastboot</td> "
			$effi_report += "    <td scope=col>$Uptime</td> "
			$effi_report += "    <td scope=col>$LastPatch</td> "
			$effi_report += "    <td scope=col>$PatchDT</td> "
			$effi_report += "    <td scope=col>$VolC</td> "
			$effi_report += "    <td scope=col>$AutoServ</td> "
			$effi_report += "    <td scope=col>$TestPath</td> "
			$effi_report += "    <td scope=col>$RDPQ</td> "
			$effi_report += "    </tr> "
				
			;break
			}
		$True {
		    	
			   $check = ''
			   $Error.Clear()
			   $ltime = Get-WmiObject Win32_LocalTime -ComputerName $server
			   $computers = Get-WMIObject -computername $server -class Win32_OperatingSystem 
			   Foreach ($system in $computers) { 
							
					$Bootup = $system.LastBootUpTime
					$LastBootUpTime = WMIDateStringToDate($Bootup)
					
					$LocalServerTime = $system.LocalDateTime
					$CurrentLST = WMIDateStringToDate($LocalServerTime)
			    	
						 $now = $CurrentLST 
			    		 $Uptime = $now - $lastBootUpTime 
						 
			    		 $d = $Uptime.Days 
			   			 $h = $Uptime.Hours 
			    		 $m = $uptime.Minutes 
			    		 $ms= $uptime.Milliseconds 
				   }   
				
				#Check Local server time
				If ($Error -match 'Access is Denied') {$localtime = 'Access Denied'} elseIf ($Error -match 'RPC server is unavailable') { $localtime = 'RPC Error'}
				else {
						$lctime = $ltime.Month.ToString() + "/" + $ltime.Day.ToString() + "/" +$ltime.Year.ToString() + " " +$ltime.Hour.ToString() + ":" +$ltime.Minute.ToString()
						$localtime = $lctime
					 }
								
				#Last boot & Uptime
				$LastBoot = $LastBootUpTime
				$Uptime = $d
				
				#Check Last Patch KB , Date & Time
				Write-Progress -Activity "Initiating Server Health Check" -status "Processing ($count/$servercount): $server" -CurrentOperation "Checking last KB patch.." -PercentComplete ($count/$servercount * 100)
				$Error.Clear()
				$lh = ((Get-HotFix -ComputerName $server | Select-Object description,hotfixid,installedby,@{l="InstalledOn";e={[DateTime]::Parse($_.psbase.properties["installedon"].value,$([System.Globalization.CultureInfo]::GetCultureInfo("en-US")))}} | sort installedon)[-1])
				
				If(($Error -match 'null array')-or($Error -match 'RPC server is unavailable')){
					$hf = $false
					$AutoServ = 'No Date Captured'
					$LastPatch = $null
					$PatchDT = $null
					} 
				else {
					$hf = $true
					$LastPatch = $lh.HotFixID
					$PatchDT = $lh.InstalledOn
					}
												
				#Check Vol C:
				Write-Progress -Activity "Initiating Server Health Check" -status "Processing ($count/$servercount): $server" -CurrentOperation "Checking system volume.." -PercentComplete ($count/$servercount * 100)
				$Error.Clear()
				$logicaldisk = get-wmiobject win32_logicaldisk -Computername $server | Where-Object {$_.DeviceID -eq "C:"}
				
				If (($Error -match 'divide by zero')-or($Error -match 'Access is Denied')-or($Error -match 'RPC server is unavailable')) { 
						$volErr = $true 
						$VolC = 'Error Occured'
						} 
					else {
						$volErr = $false
						Foreach ($drive in $logicaldisk)
						{
							$totalsize = $drive.size
							$freespace = $drive.freespace		
							[float]$TotalFreeSpace = [Math]::Round($freespace / 1073740824, 2)
							[float]$percentagefree = [Math]::Round(($freespace / $totalsize ) * 100, 0)
						}
						$VolC = $percentagefree
						}
				
				#Automatic service check
				Write-Progress -Activity "Initiating Server Health Check" -status "Processing ($count/$servercount): $server" -CurrentOperation "Checking automatic services.." -PercentComplete ($count/$servercount * 100)
				$Ignore=@( 
						'Microsoft .NET Framework NGEN v4.0.30319_X64', 
						'Microsoft .NET Framework NGEN v4.0.30319_X86', 
						'Multimedia Class Scheduler', 
						'Performance Logs and Alerts', 
						'SBSD Security Center Service', 
						'Shell Hardware Detection', 
						'Software Protection', 
						'TPM Base Services',
						'Windows Licensing Monitoring Service'
						'Google Update Service (gupdate)'; 
						 )
				
				$Error.Clear()
				$Services = Get-WmiObject Win32_Service -ComputerName $server | Where {$_.StartMode -eq 'Auto' -and $Ignore -notcontains $_.DisplayName -and $_.State -ne 'Running'}
				$ServStartErr = ''
				$AutoServ = ''
				$servErr = 0
								
				If ($Error -match 'Access is Denied') {
					$servc = 0
					$serv = 0
					$servResult = 0
					$servErr = 1
					$ServStartErr = "</br><font color=#FF0>Access Denied</font>"
				
				} elseif ($Error -match 'RPC server is unavailable') {
					$servc = 0
					$serv = 0
					$servResult = 0
					$servErr = 1
					$ServStartErr = "</br><font color=#FF0>RPC Error</font>"
						}
					else {								
						If ($Services -eq $null) {
							$servc = 0
							$service = ''
							$servResult = 0
							$serv = 0	
							$servErr = 0
						} 
						else {
								$servc = ($services | Measure-Object).count
								$service = ''
								$servErr = ''
								$serv = 0 
								$servResult = 0
								$s = ''
																						
								Foreach ($service in $Services){
								$s = $service.displayname
								Write-Progress -Activity "Initiating Server Health Check" -status "Processing ($count/$servercount): $server" -CurrentOperation "Checking stopped automatic services: $s" -PercentComplete ($count/$servercount * 100)
								$result = $null
									switch ($Service.state) {
										'Stopped' {
													$rstart = $service.StartService()
													$rvstart = $rstart.returnvalue
													IF (($rvstart -eq "0") -or ($rvstart -eq "10")){$result = 1} else {$result = 0}
													$serv += $result
													break;
													}
										'Start Pending' {
														$result = 1
														$serv += $result
														break;
														} 
										'Stop Pending' {break;}
										'Running' {
													$result = 1
													$serv += $result
													break;
													}
										'Continue Pending' {
															$result = 1
															$serv += $result
															break;
															}
										'Pause Pending' {break;}
										'Paused' {break;}
										'Unknown' {
													$result = 0
													$serv += $result
													break;
													}
										
										default { 
											$result = 0
											$serv += $result
											break;
										} 
									}
										If ($result -eq $null) {$ServStartErr += "</br><font color=#009933>" + $service.displayname + "</font>"}
										elseIf ($result -eq 0) {$ServStartErr += "</br><font color=#FF0000>" + $service.displayname + "</font>"}
										elseIf ($result -eq 1) {$ServStartErr += "</br><font color=#009933>" + $service.displayname + "</font>"}
										else {$ServStartErr += "</br><font color=#FF0000>" + $service.displayname + "</font>"} #unknown result
								} 
						}				
				}		
											
				#How many automatic service which are failed to start?
				$servResult = $servc - $serv
				
				#Automatic services full result
				$AutoServ = "($servResult/$servc)" + $ServStartErr
				
				#Test Path / Permission
				Write-Progress -Activity "Initiating Server Health Check" -status "Processing ($count/$servercount): $server" -CurrentOperation "Testing permission.." -PercentComplete ($count/$servercount * 100)
				$tPath = "\\" + $server + "\C$"
				If ((Test-Path -Path $tPath) -eq $True) {$TestPath = $true}	else {$TestPath = $false}	
					If ($TestPath -eq $true) {$TestPath = 'Passed'} elseif ($TestPath -eq $false) {$TestPath = 'Denied'}
					else {$TestPath = 'Error'}
				
				#RDP Connection Test
				Write-Progress -Activity "Initiating Server Health Check" -status "Processing ($count/$servercount): $server" -CurrentOperation "Testing RDP Connection.." -PercentComplete ($count/$servercount * 100)
					$Port = 3389
					$sleepTime = 2 
						Try {
							$socket = New-Object Net.Sockets.TcpClient($server, $Port)
				       		if ($socket.Connected) {$RDPQ = 'OK'}
							$socket.Close()
						}
						catch [System.Management.Automation.MethodInvocationException] {
				                if ($_.Exception.InnerException.GetType() -eq [System.Net.Sockets.SocketException]) {
				                        $RDPQ = 'Failed'
				                        Start-Sleep $sleepTime
				                }
				                else {
				                        throw $_.Exception.InnerException
				                }
						}
				
				<#Continue with checks
				Mem
				ConnTest
				and more ... as a function / future plans
				#>
				
				$effi_report += " <tr class=HCInfo> "
				$effi_report += "    <td scope=col>$count</td> "
				$effi_report += "    <td scope=col>$server</td> "
				IF ($localtime -match 'Denied') {$effi_report += "    <td class=HCInfoRed scope=col>Access Denied</td> "} elseif ($localtime -match 'RPC'){$effi_report += "    <td class=HCInfoRed scope=col>RPC Error</td> "}  else {$effi_report += "    <td scope=col>$localtime</td> "}
				IF ($Lastboot -eq $null){$effi_report += "    <td class=HCInfoRed scope=col>No Date Captured</td> "} else {$effi_report += "    <td scope=col>$Lastboot</td> "}
				IF ($Uptime -eq $null){$effi_report += "    <td class=HCInfoRed scope=col>N/A</td> "} else {$effi_report += "    <td scope=col>$Uptime Days</td> "}
				IF ($LastPatch -eq $null) {$effi_report += "    <td class=HCInfoRed scope=col>N/A</td> "} else {$effi_report += "    <td scope=col>$LastPatch</td> "}
				IF ($PatchDT -eq $null) {$effi_report += "    <td class=HCInfoRed scope=col>No Date Captured</td> "} else {$effi_report += "    <td scope=col>$PatchDT</td> "}
				IF ($volErr -eq $true){$effi_report += "    <td class=HCInfoRed scope=col>$VolC</td> "} elseif ($VolC -le 10){$effi_report += "    <td class=HCInfoRed scope=col>$VolC %</td> " } elseif ($VolC -eq $null){$effi_report += "    <td class=HCInfoRed scope=col>N/A</td> " } else { $effi_report += "    <td scope=col>$VolC %</td> " }
				IF ($servErr -eq 0) {$effi_report += "    <td scope=col>$AutoServ</td> "} elseif($servErr -eq 1){$effi_report += "    <td class=HCInfoRed scope=col>$AutoServ</td> "} else {$effi_report += "    <td class=bgyellow scope=col>$AutoServ</td> "}
				IF ($TestPath -match 'Passed'){$effi_report += "    <td scope=col>$TestPath</td> "} else {$effi_report += "    <td class=HCInfoRed scope=col>$TestPath</td> "}
				IF ($RDPQ -match 'OK'){$effi_report += "    <td scope=col>$RDPQ</td> "} else {$effi_report += "    <td class=HCInfoRed scope=col>$RDPQ</td> "}
				$effi_report += "    </tr> "
						    
			;break
			}
		default {Write-host "Unknown Error has occurred"; break}
		} 
}

$effi_report += @"
	</table>
	</body>
	</html>
"@


#$effi_report | out-file -encoding ASCII -filepath $Filename
send_email

