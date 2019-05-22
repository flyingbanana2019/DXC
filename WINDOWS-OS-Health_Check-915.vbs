'******************************************************************************************
'NAME:        WINDOWS-OS-Health Check-915.vbs
'VERSION:     2.10.0
'INTERNAL VERSION: 010
'
'DESCRIPTION: This script is used to create a server status report after the server has 
'             a unexpected reboot or to verify it when you want.
'             Any problem with this script you can send a eMail to sergio.lambrisca@hp.com
'             and I will help you.
'
'INPUT:       /Verbose
'             /EventMessage: "Message to add in System Event"
'
'OUTPUT:      A status of server in Memory, Virtual memory, System Events (last three boot, HW errors, Patch Instaling) Logical Disk, Shared, 
'             Printer, Services, Process, Network and Active Directory (if server is Domain Controller)
'
'MODIFICATION LOG: 
'05/07/2011  Sergio Lambrisca	  Created.
'10/02/2011  Sergio Lambrisca	  Add parameters.
'30/02/2011  Sergio Lambrisca	  Add System Events.
'24/02/2011  Sergio Lambrisca	  Add Print.
'11/04/2012  Sergio Lambrisca   Change Print how Print Services Role.
'23/04/2012  Sergio Lambrisca   Add MEMORY.DMP information
'29/04/2012  Sergio Lambrisca	  Correct DebugInfoType problem in W2000 servers
'24/06/2012  Sergio Lambrisca   Add Overal Status and Date of creation result file
'29/06/2012  Sergio Lambrisca   Add Up Time value, Get HW events for HP and DELL servers
'04/07/2012  Sergio Lambrisca   Correct Server Role Detection in W2000 and W2003
'20/07/2012  Sergio Lambrisca   Change the parameter /Verbose:true to /Verbose only
'27/07/2012  Sergio Lambrisca   Correct get HW Dell events
'16/08/2012  Sergio Lambrisca   Correct get HW HP events, add information code to detect power lost
'29/08/2012  Sergio Lambrisca   Add NIC status
'09/01/2013  Sergio Lambrisca   Solve problems in events display.  In some case the server have a boot in the last week and the script don´t show any event.
'11/01/2013  Sergio Lambrisca   Add OS Architecture in the OS detail
'******************************************************************************************
Option Explicit
On Error Resume Next 

'=================================================================================
' VARIABLE DEFINITIONS
'=================================================================================
Dim objWMIService, objWMISet, objWMIServiceNested, objWMISetNested, intObjRow, intObjRowNested, objShell
Dim objWMIADService, objWMIADSet
Dim objFSO, objFile
Dim objWMIDLastBoot
Dim objWMIDEventTimeWritten, objDTMFromDate
Dim objParameters

Public strOSVersion				:strOSVersion= ""
Public strADComputerRole	:strADComputerRole= ""
Public strOSArchitecture  :strOSArchitecture= ""
Public strCommand

Dim intUpTime              :intUpTime= 0
Dim intCounter					   :intCounter= 0

Dim strComputerDestination :strComputerDestination = "."
Dim intExitCode					   :intExitCode = 0 'Used to signal if the script finish with error or not
Dim strOveralStatus			   :strOveralStatus= "OK"

Dim bolWriteEventMessage	 :bolWriteEventMessage= vbFalse
Dim bolAlertCheck				   :bolAlertCheck= vbFalse
Dim bolEventRead           :bolEventsRead= vbFalse
Dim bolShutDown            :bolShutDown= vbFalse
Dim bolAbnormalShutDown    :bolAbnormalShutDown= vbFalse
Dim bolOVOMonitored				 :bolOVOMonitored= vbFalse

Dim arrADComputerRole, arrDriveType, arrMemoryType, arrSKU, arrServiceSkiped, arrTrustDirection
Dim arrEventType, arrExtendedPrinterStatus, arrPrinterError, arrDebugInfoType
Dim arrNetConnectionStatus, arrConfigManagerErrorCode
Dim arrRoleServices(8,4)

Dim strStatusSO, strStatusLogicalHD, strStatusShared, strStatusNIC, strStatusService, strStatusProcess, strStatusEventMessage  'Used to store the status of each server test; these variavels are part or Overall Status
Dim strStatusPrintServices
Dim strStatusAD, strStatusADRepPartners, strStatusADReplication, strStatusADTrust, strStatusADServices 'Used to store the status of each server test; these variavels are part or Overall Status
Dim strStatusADSYSVOL
Dim strTitle, strTitle2, strAlert
Dim strOVOVersion
Dim strParameterEventMessage, bolIsVerbose, strEventMessage
Dim strOSLanguage, strShareType
Dim strWindowsDirectory, strDebugFilePath


'Initialization of Constant
Const constGigaByte = 1073741824 '1Gb
Const constMegaByte = 1048576 '1Mb
Const constKiloByte = 1024 '1Kb
Const constMemoryWarningUmbral= 0.20
Const constVtMemoryWarningUmbral= 0.5
Const constRebootCount= 3
Const constHDSpaceWarningUmbral= 0.10
Const constProcHandleWarningUmbral= 5000
Const constProcMemoryWarningUmbral= 500  'Mb
Const constProcPercWarningUmbral= 45
Const constDayEvents= -7

Const CONVERT_TO_LOCAL_TIME = True
    
'Initialization of Varaibles
arrADComputerRole= Array("Standalone Workstation", _
					"Member Workstation", _ 
					"Standalone Server", _ 
					"Member Server", _ 
					"Backup Domain Controller", _ 
					"Primary Domain Controller")

arrDebugInfoType= Array("None", _
					"Complete Memory Dump", _
					"Kernel Memory Dump", _
					"Small Memory Dump")

arrEventType= Array("Information", _
				"Error", _
				"Warning", _
				"Information", _
				"Security Audit Success", _
				"Security Audit Failure")
    
arrDriveType= Array("Unknow", _
				"No Root Directory", _
				"Removable Disk", _
				"Local Disk", _
				"Network Drive", _
				"Compac Disc", _
				"RAM Disk")

arrMemoryType= Array("Unknown", _
				"Other", _
				"DRAM", _
				"Synchronous DRAM", _
				"Cache DRAM", _
				"EDO", _
				"EDRAM", _
				"VRAM", _
				"SRAM", _
				"RAM", _
				"ROM", _
				"Flash", _
				"EEPROM", _
				"FEPROM", _
				"EPROM", _
				"CDRAM", _
				"3DRAM", _
				"SDRAM", _
				"SGRAM", _
				"RDRAM", _
				"DDR", _
				"DDR-2")

arrSKU= Array("Undefined", _
				"Ultimate Edition", _
				"Home Basic Edition", _
				"Home Premium Edition", _
				"Enterprise Edition", _
				"Home Basic N Edition", _
				"Business Edition", _
				"Standard Server Edition", _
				"Datacenter Server Edition", _
				"Small Business Server Edition", _
				"Enterprise Server Edition", _
				"Starter Edition", _
				"Datacenter Server Core Edition", _
				"Standard Server Core Edition", _
				"Enterprise Server Core Edition", _
				"Enterprise Server Edition for Itanium-Based Systems", _
				"Business N Edition", _
				"Web Server Edition", _
				"Cluster Server Edition", _
				"Home Server Edition", _
				"Storage Express Server Edition", _
				"Storage Standard Server Edition", _
				"Storage Workgroup Server Edition", _
				"Storage Enterprise Server Edition", _
				"Server For Small Business Edition", _
				"Small Business Server Premium Edition")

arrNetConnectionStatus= Array("Disconnected", _
				"Connecting", _
				"Connected", _
				"Disconnecting", _
				"Hardware not present", _
				"Hardware disabled", _
				"Hardware malfunction", _
				"Media disconnected", _
				"Authenticating", _
				"Authentication succeeded", _
				"Authentication failed", _
				"Invalid address", _
				"Credentials required")


arrConfigManagerErrorCode= Array("Device is working properly", _
									"Device is not configured correctly", _
									"Windows cannot load the driver for this device", _
									"Driver for this device might be corrupted, or the system may be low on memory or other resources", _
									"Device is not working properly. One of its drivers or the registry might be corrupted", _
									"Driver for the device requires a resource that Windows cannot manage", _
									"Boot configuration for the device conflicts with other devices", _
									"Cannot filter", _
									"Driver loader for the device is missing", _
									"Device is not working properly. The controlling firmware is incorrectly reporting the resources for the device", _
									"Device cannot start", _
									"Device failed", _
									"Device cannot find enough free resources to use", _
									"Windows cannot verify the device's resources", _
									"Device cannot work properly until the computer is restarted", _
									"Device is not working properly due to a possible re-enumeration problem", _
									"Windows cannot identify all of the resources that the device uses", _
									"Device is requesting an unknown resource type", _
									"Device drivers must be reinstalled", _
									"Failure using the VxD loader", _
									"Registry might be corrupted", _
									"System failure. If changing the device driver is ineffective, see the hardware documentation. Windows is removing the device", _
									"Device is disabled", _
									"System failure. If changing the device driver is ineffective, see the hardware documentation", _
									"Device is not present, not working properly, or does not have all of its drivers installed", _
									"Windows is still setting up the device", _
									"Windows is still setting up the device", _
									"Device does not have valid log configuration", _
									"Device drivers are not installed", _
									"Device is disabled. The device firmware did not provide the required resources", _
									"Device is using an IRQ resource that another device is using", _
									"Device is not working properly. Windows cannot load the required device drivers")


arrServiceSkiped= Array("ccmsetup", _
				"Diagnostic Policy Service", _
				"Distributed Transaction Coordinator", _
				"eTrust Policy Compliance", _
				"KtmRm for Distributed Transaction Coordinator", _
				"Microsoft .NET Framework NGEN v4.0.30319_X64", _
				"Microsoft .NET Framework NGEN v4.0.30319_X86", _
				"Microsoft Software Shadow Copy Provider", _
				"Performance Logs and Alerts", _
				"Shell Hardware Detection", _
				"Software Protection", _
				"TPM Base Services", _
				"Volume Shadow Copy", _
				"Windows Font Cache Service", _
				"Windows Licensing Monitoring Service", _
				"Windows Remote Management (WS-Management)", _
				"Windows Service Pack Installer update service Properties", _
				"Windows Update")

arrExtendedPrinterStatus= Array("Other", _
				"Ready", _
				"Idle", _
				"Printing", _
				"Warming Up", _
				"Stopped Printing", _
				"Offline", _
				"Paused", _
				"Error", _
				"Busy", _
				"Not Available", _
				"Waiting", _
				"Processing", _
				"Initialization", _
				"Power Save", _
				"Pending Deletion", _
				"I/O Active", _
				"Manual Feed")

arrPrinterError= Array("", _
				"Other", _
				"No Error", _
				"Low Paper", _
				"No Paper", _
				"Low Toner", _
				"No Toner", _
				"Door Open", _
				"Jammed", _
				"Offline", _
				"Service Requested", _
				"Output Bin Full")

arrTrustDirection= Array("Inbound", _
				"Outbound", _
				"Bidirectional")

'arrRoleServices(5,0)= "Active Directory Domain Services" (Name of Rol)
'arrRoleServices(5,1)= "W2003 W2008" (Version of OS where the service is supported)
'arrRoleServices(5,2)= "Windows Time" (Services Name)
'arrRoleServices(5,3)= False (True= Service Enabled | False= Service Disabled)
'arrRoleServices(5,4)= False (True= Service Running | False= Service Not Running)
arrRoleServices(0,0)= "Active Directory Domain Services"
arrRoleServices(0,1)= "W2003 W2008"
arrRoleServices(0,2)= "File Replication Service"
arrRoleServices(0,3)= False
arrRoleServices(0,4)= False
arrRoleServices(1,0)= "Active Directory Domain Services"
arrRoleServices(1,1)= "W2003 W2008"
arrRoleServices(1,2)= "Intersite Messaging"
arrRoleServices(1,3)= False
arrRoleServices(1,4)= False
arrRoleServices(2,0)= "Active Directory Domain Services"
arrRoleServices(2,1)= "W2003 W2008"
arrRoleServices(2,2)= "Kerberos Key Distribution Center"
arrRoleServices(2,3)= False
arrRoleServices(2,4)= False
arrRoleServices(3,0)= "Active Directory Domain Services"
arrRoleServices(3,1)= "W2008"
arrRoleServices(3,2)= "Netlogon"
arrRoleServices(3,3)= False
arrRoleServices(3,4)= False
arrRoleServices(4,0)= "Active Directory Domain Services"
arrRoleServices(4,1)= "W2003"
arrRoleServices(4,2)= "Net Logon"
arrRoleServices(4,3)= False
arrRoleServices(4,4)= False
arrRoleServices(5,0)= "Active Directory Domain Services"
arrRoleServices(5,1)= "W2003 W2008"
arrRoleServices(5,2)= "Windows Time"
arrRoleServices(5,3)= False
arrRoleServices(5,4)= False
arrRoleServices(6,0)= "Print Services"
arrRoleServices(6,1)= "W2003 W2008"
arrRoleServices(6,2)= "Print Spooler"
arrRoleServices(6,3)= False
arrRoleServices(6,4)= False
arrRoleServices(7,0)= "CITRIX"
arrRoleServices(7,1)= "W2003 W2008"
arrRoleServices(7,2)= "Independent Management Architecture"
arrRoleServices(7,3)= False
arrRoleServices(7,4)= False
arrRoleServices(8,0)= "CITRIX"
arrRoleServices(8,1)= "W2003 W2008"
arrRoleServices(8,2)= "Citrix MFCOM Service"
arrRoleServices(8,3)= False
arrRoleServices(8,4)= False

'=================================================================================
' TAKE GENERAL DATA
'=================================================================================
Set objParameters = WScript.Arguments.Named
'First parameter
If objParameters.Exists("Verbose") Then
	bolIsVerbose= vbTrue
Else
	bolIsVerbose= vbFalse
End If 
'Second parameter
strParameterEventMessage= objParameters.Item("EventMessage") 


If Not IsEmpty(strParameterEventMessage) Then
	Set objShell = CreateObject("WScript.Shell")
	strCommand= "eventcreate /T Warning /ID 999 /L System /D " & Chr(34) & strParameterEventMessage & Chr(34)
	objShell.Run strCommand
End If
Set objParameters= Nothing

strOVOVersion= F_GetOpenViewVersion
If strOVOVersion <> "Was impossible to detect OVO version" Then
	bolOVOMonitored= vbTrue
End If

Set objWMIDLastBoot= CreateObject("WbemScripting.SWbemDateTime")
Set objWMIService = GetObject("winmgmts:{impersonationLevel=Impersonate}!\\" & strComputerDestination & "\root\cimv2")
Set objWMIServiceNested = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputerDestination & "\root\cimv2")

Set objWMISet= objWMIService.ExecQuery("Select AddressWidth from Win32_Processor")
For Each intObjRow In objWMISet
  strOSArchitecture= intObjRow.AddressWidth
Next


Set objWMISet= objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
For Each intObjRow In objWMISet
	strOSVersion= FWMI_OSVersion(intObjRow.Caption)
	objWMIDLastBoot.Value= intObjRow.LastBootUpTime
	intUpTime= DateDiff("h",objWMIDLastBoot.GetVarDate,Now())
	
	Select Case intObjRow.OSLanguage 
		Case 9
			strOSLanguage = "English"
		Case 1033
			strOSLanguage = "English - United States"
		Case 1034
			strOSLanguage = "Spanish - Traditional Sort"
		Case 1046
			strOSLanguage = "Portuguese - Brazil"
		Case 2057
			strOSLanguage = "English - United Kingdom"
		Case 3082 
			strOSLanguage = "Spanish - International Sort"
		Case 11274
			strOSLanguage = "Spanish - Argentina"
		Case else
			strOSLanguage = "Other diferent to Portuguese, Spanish or English"
	End Select
	Wscript.echo F_AssemblingTitle("Health Check v2.10.0",90)
	Wscript.echo F_AssemblingTitle("",90)
	Wscript.echo F_AssemblingTitle("Status of server " & intObjRow.CSName,90)
	WScript.echo F_AssemblingTitle("Date: " & Now(),90)
	WScript.echo F_AssemblingTitle("OVO version: " & strOVOVersion,90)
	Wscript.echo F_AssemblingTitle("",90)
	Wscript.echo ""
	Wscript.echo "       Operating System: " & intObjRow.Caption
	Wscript.echo "                Version: " & intObjRow.Version
	Wscript.echo "           Architecture: " & strOSArchitecture & " bits"
	Wscript.echo "       Type Description: " & intObjRow.OtherTypeDescription
	Wscript.echo "           Service Pack: " & intObjRow.CSDVersion
	Wscript.echo "               Language: " & strOSLanguage
	Wscript.echo "      Windows directory: " & intObjRow.WindowsDirectory
	Wscript.echo "       System directory: " & intObjRow.SystemDirectory & VbCrLf
	WScript.echo "              Time Zone: " & intObjRow.CurrentTimeZone
	WScript.echo "              Last Boot: " & objWMIDLastBoot.GetVarDate
	WScript.echo "                Up Time: " &  intUpTime & " Hs." & vbCrLf
	WScript.echo "                 Memory: " & intObjRow.TotalVisibleMemorySize & " Kb (" & int(intObjRow.TotalVisibleMemorySize/constMegaByte) & " Gb.)"
	WScript.echo "            Free Memory: " & intObjRow.FreePhysicalMemory & " Kb (" & int(intObjRow.FreePhysicalMemory/constMegaByte) & " Gb.)"
	Wscript.echo "         Virtual Memory: " & intObjRow.TotalVirtualMemorySize & " Kb (" & int(intObjRow.TotalVirtualMemorySize/constMegaByte) & " Gb.)"
	Wscript.echo "    Free Virtual Memory: " & intObjRow.FreeVirtualMemory & " Kb (" & int(intObjRow.FreeVirtualMemory/constMegaByte) & " Gb.)"
	WScript.echo "Free Space in Pag. File: " & intObjRow.FreeSpaceInPagingFiles & " Kb (" & int(intObjRow.FreeSpaceInPagingFiles/constMegaByte) & " Gb.)"
	WScript.Echo vbCrLf	
	        
	If intObjRow.FreePhysicalMemory < (intObjRow.TotalVisibleMemorySize * constMemoryWarningUmbral) Then
		strStatusSO = "The server has little available memory.  Please check." & VbCrLf 
		strOveralStatus = "With Warnings"
	Else
		strStatusSO = "The memory of the server is OK." & VbCrLf
	End If 

	If intObjRow.FreeVirtualMemory < (intObjRow.TotalVirtualMemorySize * constMemoryWarningUmbral) Then
		strStatusSO = strStatusSO & "The server has little virtual memory available.  Please check." & VbCrLf 
		strOveralStatus = "With Warnings"
	Else
		strStatusSO = strStatusSO & "The virtual memory available of the server is OK." & VbCrLf
	End If 
	strWindowsDirectory= intObjRow.WindowsDirectory
Next

Set objWMISet = objWMIService.ExecQuery("Select * from Win32_OSRecoveryConfiguration")
For Each intObjRow In objWMISet
	Wscript.echo "        Debug File Path: " & intObjRow.DebugFilePath
	If (strOSVersion = "W2000") Then
		WScript.echo "    Debugging File Type: "
	Else
		WScript.echo "    Debugging File Type: " & arrDebugInfoType(intObjRow.DebugInfoType)
	End If
	WScript.echo "Overwrite Existing File: " & intObjRow.OverwriteExistingDebugFile
	
	If (InStr(intObjRow.DebugFilePath, "%SystemRoot%") > 0) Then
		strDebugFilePath= Replace(intObjRow.DebugFilePath,"%SystemRoot%",strWindowsDirectory)
	Else
		strDebugFilePath= intObjRow.DebugFilePath
	End If
Next

Set objFSO= CreateObject("scripting.FileSystemObject")
If (objFSO.FileExists(strDebugFilePath)) Then
	Set objFile= objFSO.GetFile(strDebugFilePath)
	WScript.echo "        Dump file Exist: " & "True"
	WScript.echo "      Dump file Created: " & objFile.DateCreated
	WScript.echo "Dump file Last Modified: " & objFile.DateLastModified
Else
	WScript.echo "        Dump file Exist: " & "False"
End If
WScript.Echo vbCrLf

Set objWMISet = objWMIService.ExecQuery("Select Domain,DomainRole,Model,Manufacturer from Win32_ComputerSystem")
For Each intObjRow In objWMISet
	Wscript.echo "                 Domain: " & intObjRow.Domain
	Wscript.echo "          Computer Role: " & arrADComputerRole(intObjRow.DomainRole)
  WScript.Echo vbCrLf
	Wscript.echo "           Manufacturer: " & intObjRow.Manufacturer
	Wscript.echo "                  Model: " & intObjRow.Model
	strADComputerRole= arrADComputerRole(intObjRow.DomainRole)
Next
Wscript.echo VbCrLf & VbCrLf

'=================================================================================
' LAST BOOT INFORMATION
'=================================================================================
bolAlertCheck= vbFalse
strAlert= ""
strTitle= " Action                   |Time Generated        |Event ID |Event Type             |Source               |Computer             |User                 |Description                                                                                                           "
WScript.echo F_AssemblingTitle("System Event Log's (Critical, Error and Warning)", len(strTitle))
Wscript.echo strTitle
Wscript.echo F_AssemblingTitle("",len(strTitle))

Set objDTMFromDate = CreateObject("WbemScripting.SWbemDateTime")
objDTMFromDate.SetVarDate DateAdd("d", constDayEvents, Now()), FALSE

Set objWMIDEventTimeWritten = CreateObject("WbemScripting.SWbemDateTime")

Set objWMISet = objWMIService.ExecQuery("Select CategoryString, ComputerName, EventCode, Message, TimeGenerated, SourceName, EventType, User from Win32_NTLogEvent WHERE " _
										& " (Logfile = 'System') AND "_
										& " (TimeGenerated >= '" & objDTMFromDate & "')")

For Each intObjRow In objWMISet
  bolEventRead= vbTrue
	Select Case Trim(intObjRow.SourceName)
		Case "EventLog"
			Select Case intObjRow.EventCode  'Reboot Check
				Case 6006
					bolWriteEventMessage= vbTrue
					bolShutDown= vbTrue
					strAlert= "Shutdown"
				Case 6008
					bolWriteEventMessage= vbtrue
					bolAbnormalShutDown= vbTrue
					bolAlertCheck= vbTrue
					strAlert= "Abnormal Shutdown"
				Case 6009
					bolWriteEventMessage= vbtrue
					strAlert= "Boot"
			End Select

		Case "Save Dump"   'Patch Installation
			Select Case intObjRow.EventCode
				Case 1001
					bolWriteEventMessage= vbtrue
					strAlert= "Save Dump"
			End Select

		Case "NtServicePack"   'Patch Installation
			Select Case intObjRow.EventCode
				Case 4353, 4377
					bolWriteEventMessage= vbTrue
					strAlert= "Patch Installation" 
			End Select

		Case "Windows Update Agent"   'Patch Installation
			Select Case intObjRow.EventCode
				Case 20
					bolWriteEventMessage= vbTrue
					strAlert= "Patch Installation Error" 
			End Select

		Case "Ntfs", "PartMgr"   'Disk errors
			Select Case intObjRow.EventCode
				Case 41, 55, 59
					bolWriteEventMessage= vbTrue
					strAlert= "Disk error" 
			End Select

		Case "USER32"  'Users Boot
			Select Case intObjRow.EventCode
				Case 1073, 1074, 1075, 1076
					bolWriteEventMessage= vbTrue
					strAlert= "User boot info." 

			End Select

		Case "Kernel-Power"  '???
			Select Case intObjRow.EventCode
				Case 41
					bolWriteEventMessage= vbtrue
					strAlert= "Kernel Power" 
			End Select

		Case "Service Control Manager"
			Select Case intObjRow.EventCode
				Case 7000, 7038, 7041
					bolWriteEventMessage= vbtrue
					strAlert= "Service start error" 
			End Select
		
		Case "Microsoft-Windows-Kerberos-Key-Distribution-Center", "Microsoft-Windows-Security-Kerberos", "KDC"
			Select Case intObjRow.EventCode
				Case 28
					bolWriteEventMessage= vbtrue
					strAlert= "KKDC AD Replication Availability" 
				Case 12, 22
					bolWriteEventMessage= vbtrue
					strAlert= "KKDC AD Trust Configuration" 
				Case 19, 20, 29
					bolWriteEventMessage= vbtrue
					strAlert= "KKDC Certificate Availability" 
				Case 26, 27
					bolWriteEventMessage= vbtrue
					strAlert= "KKDC Encryption Type Configuration" 
				Case 9, 10
					bolWriteEventMessage= vbtrue
					strAlert= "KKDC Password Configuration" 
				Case 5, 6, 15, 23
					bolWriteEventMessage= vbtrue
					strAlert= "KKDC Service Availability" 
				Case 13, 14, 16
					bolWriteEventMessage= vbtrue
					strAlert= "KKDC Key Integrity" 
				Case Else
					bolWriteEventMessage= vbtrue
					strAlert= "KKDC (Another)" 
			End Select

		Case "Srv"  'DoS Attack
			Select Case intObjRow.EventCode
				Case 2025, 2026, 2027
					bolWriteEventMessage= vbtrue
					strAlert= "Warning DoS"
					bolAlertCheck= vbTrue
				Case 333, 2019, 2020
					bolWriteEventMessage= vbtrue
					strAlert= "Warning M. Pool"
					bolAlertCheck= vbTrue
			End Select

		Case Else
			If (arrEventType(intObjRow.EventType - 1) = "Warning") Or (arrEventType(intObjRow.EventType - 1) = "Error") Then
				Select Case Trim(intObjRow.SourceName)
					Case "Foundation Agents", "Server Agents", "NIC Agents", "hpilo2", "hpilo3", "Storage Agents", "HP Smart Array" ' "Cissesrv"  'HP Hardware
						Select Case intObjRow.EventCode
							'Excluded Events
							Case 400, 277, 1182, 24581, 24578, 24582, 24598
								'These events are excluded
							'HPILO2
							Case 57
								bolWriteEventMessage= vbtrue
								strAlert= "ILO Error"
							'ASR Errors
							Case 1090
								bolWriteEventMessage= vbtrue
								strAlert= "Automatic Server Recovery"
							'POST Errors
							Case 1092, 1123
								bolWriteEventMessage= vbtrue
								strAlert= "POST error"
							'Thermal Errors
							Case 1082, 1083, 1134, 1135
								bolWriteEventMessage= vbtrue
								strAlert= "Thermal Temperature error"
							Case 1084, 1091, 1136
								bolWriteEventMessage= vbtrue
								strAlert= "Thermal Temperature OK"
							'Fan Errors
							Case 1085, 1086, 1129, 1130, 1131, 1132, 1133
								bolWriteEventMessage= vbtrue
								strAlert= "Fan error"
							Case 1087
								bolWriteEventMessage= vbtrue
								strAlert= "Fan Return OK"
							'Power Errors
							Case 1103, 1124, 1125, 1126, 1127, 1128, 1137, 1138, 1139, 1155, 1156, 1158, 1159, 1160, 1161, _
							     1162, 1163, 1164, 1165, 1166, 1167, 1168, 1172, 1173, 1174, 1175, 1176, 1177, 1178
								bolWriteEventMessage= vbtrue
								strAlert= "Power Lost"
							Case 1029, 1157
								bolWriteEventMessage= vbtrue
								strAlert= "Power Return"
							'Processor Errors
							Case 1173, 1174, 1175, 1176, 1088, 1114
								bolWriteEventMessage= vbtrue
								strAlert= "Processor Error"
							Case 1089
								bolWriteEventMessage= vbtrue
								strAlert= "Processor Return OK"
							'Storage Logical Errors
							Case 1179, 1180, 1062, 1069, 1145, 1200
								bolWriteEventMessage= vbtrue
								strAlert= "Storage Logical Error"
							'Storage Physical Errors
							Case 1063, 1064, 1067, 1068, 1070, 1075, 1076, 1077, 1104, 1146, 1147, 1151, 1152, _
								 1153, 1154, 1155, 1164, 1185, 1188, 1189, 1190, 1196, 1199, 1201, 1202, 1203, _
								 1212, 1213, 1214, 1215, 1216, 1217, 1218, 1219, 1220
								bolWriteEventMessage= vbtrue
								strAlert= "Storage Physical Error"
							Case 1078, 1165, 1179
								bolWriteEventMessage= vbtrue
								strAlert= "Storage Physical OK"
							'Memory Errors
							Case 1024, 1025, 1026, 1027, 1028, 1035, 1036, 1037, 1038, 1039
								bolWriteEventMessage= vbtrue
								strAlert= "Memory Error"
							Case 1030
								bolWriteEventMessage= vbtrue
								strAlert= "Memory OK"
							'ILO Errors
							Case 1109, 1110, 1111, 1112, 1113, 1116, 1117
								bolWriteEventMessage= vbtrue
								strAlert= "ILO Error"
							Case 1118
								bolWriteEventMessage= vbtrue
								strAlert= "ILO return OK"
							'NIC Errors
							Case 1283, 1291, 1293
								bolWriteEventMessage= vbtrue
								strAlert= "NIC Error"
							Case 1082, 1290, 1292
								bolWriteEventMessage= vbtrue
								strAlert= "NIC return OK"
							'Any Other Errors
							Case Else
								bolWriteEventMessage= vbtrue
								strAlert= "HW Evt"
						End Select
					Case "Server Administrator"  'DELL Hardware
						Select Case intObjRow.EventCode
							'Excluded Events
							Case 2334, 2242, 2243, 2180, 2181, 2188, 2189, 2199
								'These events are excluded
							'ASR Errors
							Case 1006
								bolWriteEventMessage= vbtrue
								strAlert= "Automatic Server Recovery"
							'Thermal Errors
							Case 1004, 1007
								bolWriteEventMessage= vbtrue
								strAlert= "Thermal Temperature error"
							'Power Errors
							Case 1013, 1050, 1051, 1053, 1054, 1055, 1150, 1151, 1153, 1154, 1155, 1350, 1351, 1353, 1354, 1355, _
							     1500, 1501, 1503, 1504, 1505, 1700, 1701, 1703, 1704, 1705
								bolWriteEventMessage= vbtrue
								strAlert= "Power Lost"
							Case 1052, 1152, 1352, 1502, 1702
								bolWriteEventMessage= vbtrue
								strAlert= "Power Return"
							'Fan Errors
							Case 1100, 1101, 1103, 1104, 1105, 1450, 1451, 1452, 1453, 1454, 1455
								bolWriteEventMessage= vbtrue
								strAlert= "Fan error"
							Case 1102
								bolWriteEventMessage= vbtrue
								strAlert= "Fan Return OK"
							'Redundancy Unit Messages
							Case 1300, 1301, 1302, 1303, 1305, 1306
								bolWriteEventMessage= vbtrue
								strAlert= "Redundancy error"
							Case 1304
								bolWriteEventMessage= vbtrue
								strAlert= "Redundancy return OK"
							'Memory Errors
							Case 1403, 1404
								bolWriteEventMessage= vbtrue
								strAlert= "Memory Error"
							'Processor Errors
							Case 1600, 1601, 1603, 1604, 1605
								bolWriteEventMessage= vbtrue
								strAlert= "Processor Error"
							Case 1602
								bolWriteEventMessage= vbtrue
								strAlert= "Processor Return OK"
							'Storage Errors
							Case 2174, 2048, 2049, 2050, 2051, 2054, 2055, 2056, 2057, 2058, 2059, 2060, 2067, 2070, _
							     2074, 2076, 2077, 2079, 2080, 2081, 2082, 2094, 2095, 2098, 2099, 2100, 2101, 2102, _
							     2103, 2104, 2106, 2107, 2108, 2109, 2110, 2111, 2112, 2114, 2115, 2312, 2101, 2121, _
							     2116, 2117, 2118, 2120, 2122, 2123, 2131, 2132, 2163, 2169, 2174, 2176, 2177, 2178, _
							     2179, 2183, 2184, 2185, 2187, 2195, 2196, 2197, 2198, 2200, 2201, 2202, 2203, 2204, _
							     2205, 2210, 2211, 2212, 2213, 2247, 2248, 2278, 2358
								bolWriteEventMessage= vbtrue
								strAlert= "Storage Error"
							Case 2052, 2053, 2061, 2062, 2063, 2064, 2065, 2075, 2083, 2085, 2086, 2087, 2088, 2089, _
							     2090, 2091, 2092, 2105, 2121, 2124, 2158, 2170, 2171, 2172, 2175, 2176
								bolWriteEventMessage= vbtrue
								strAlert= "Storage return OK"
							'Any Other Errors
							Case Else
								bolWriteEventMessage= vbtrue
								strAlert= "HW Evt"
						End Select
				End Select
			End If
	End Select
	
	objWMIDEventTimeWritten.Value= intObjRow.TimeGenerated
	If IsNull(intObjRow.Message) Then
		strEventMessage= "The description for Event cannot be found. Either the component that raises this event is not installed " _
							& "on your local computer or the installation is corrupted. You can install or repair the component on the " _
							& "local computer."
	Else
		strEventMessage= Replace(intObjRow.Message,VbCrLf,"  ")
	End If
		
	If bolWriteEventMessage Then
		WScript.echo " " _
			& F_strDataTrim(strAlert,25,0) & "|" _
			& F_strDataTrim(objWMIDEventTimeWritten.GetVarDate(CONVERT_TO_LOCAL_TIME),22,1) & "|" _
			& F_strDataTrim(intObjRow.EventCode,8,0) & " |" _
			& F_strDataTrim(arrEventType(intObjRow.EventType),22,0) & " |" _
			& F_strDataTrim(intObjRow.SourceName,20,0) & " |" _
			& F_strDataTrim(intObjRow.ComputerName,20,0) & " |" _
			& F_strDataTrim(intObjRow.User,20,0) & " |" _
			& F_strDataTrim(strEventMessage,200,0)
    	strAlert=""
		bolWriteEventMessage= vbFalse
	End If
Next
If Err.Number <> 0 Then
  If intUpTime < 169 Then
    strStatusEventMessage= "WARNING:  The read of events finished unexpected and the server had a boot in last week." & vbCrLf _
                         & "          It's possible you couldn't see all filtered events." & vbCrLf _
                         & "          Please, to verify further, login to the server and verify the system events looking for reboot events." & vbCrLf
  Else
    strStatusEventMessage= "WARNING:  The read of events finished unexpected." & vbCrLf _
                         & "          It's possible you couldn't see all filtered events." & vbCrLf
  End If
	strOveralStatus = "With Warnings"
Else
  strStatusEventMessage= "In the last week the server didn't have reboots or all of them were expected." & vbCrLf
  If intUpTime < 169 Then
    If bolEventRead Then
      If Not bolShutDown And Not bolAbnormalShutDown Then
        strStatusEventMessage= "WARNING:  The server had a reboot in the last week and the script read events but it couldn't detect reboot events." & vbCrLf _
                             & "          Please check further in the system events log." & vbCrLf
      Else
        If bolShutDown Then
          strStatusEventMessage= "In the last week the server had expected reboots." & vbCrLf
        End If
        If bolAbnormalShutDown Then
          strOveralStatus = "With Warnings"
          If bolShutDown Then
            strStatusEventMessage= "WARNING: In the last week the server had unexpected reboot.  Please check further in the system events log." & vbCrLf _
                                 & "         The server also had expected reboots." & vbCrLf
          Else
            strStatusEventMessage= "WARNING: In the last week the server had unexpected reboot.  Please check further in the system events log." & vbCrLf
          End If
        End If
      End If
    Else
      strStatusEventMessage= "WARNING:  The script couldn't read events and the server had a reboot in the last week." & vbCrLf _
                           & "          Please check further in the system events log." & vbCrLf
    End If
  End If  
End If

bolAlertCheck= vbFalse
If bolIsVerbose Then
	Wscript.echo VbCrLf
End If



'=================================================================================
' DISK INFORMATION
'=================================================================================
strAlert= ""
strTitle= " Action  |Drive Status |Disk |Format |Type of Disk      |Total Space |Available Space |Dirty "
bolAlertCheck= vbFalse
strStatusLogicalHD= ""
If bolIsVerbose Then
	Wscript.echo F_AssemblingTitle("Logical  Disk Information",len(strTitle))
	Wscript.echo strTitle
	Wscript.echo F_AssemblingTitle("",len(strTitle))
End If
Set objWMISet = objWMIService.ExecQuery("Select * from Win32_LogicalDisk")
For Each intObjRow In objWMISet
	If strOSVersion = "W2000" Then
		If (Int(intObjRow.FreeSpace/constGigaByte) < Int(intObjRow.Size * constHDSpaceWarningUmbral/constGigaByte)) or _
			(intObjRow.Status <> "OK") Then
			strAlert= "Warning"
			bolAlertCheck= vbTrue
			if (Int(intObjRow.FreeSpace/constGigaByte) < Int(intObjRow.Size * constHDSpaceWarningUmbral/constGigaByte)) Then
				strStatusLogicalHD = strStatusLogicalHD & "The hard disk " & intObjRow.Name & " have low space. It have less than 10% of total disk space.  Please check." & VbCrLf 
			End If
			if (intObjRow.Status <> "OK") Then
				strStatusLogicalHD = strStatusLogicalHD & "The HD " & intObjRow.Name & " have an error.  Please check." & VbCrLf 
			End If
		Else
			strAlert= ""    	
		End If
		If bolIsVerbose Then
			Wscript.echo " " _
						& F_strDataTrim(strAlert,8,0) & "|" _
						& F_strDataTrim(intObjRow.Status,13,0) & "|" _
						& F_strDataTrim(intObjRow.Name,5,0) & "|" _
						& F_strDataTrim(intObjRow.FileSystem,7,0) & "|" _
						& F_strDataTrim(arrDriveType(intObjRow.DriveType),18,0) & "|" _
						& F_strDataTrim(Int(intObjRow.Size/constGigaByte),8,1) & " Gb.|" _
						& F_strDataTrim(Int(intObjRow.FreeSpace/constGigaByte),12,1) & " Gb.|"
		End If
	Else
		If (Int(intObjRow.FreeSpace/constGigaByte) < Int(intObjRow.Size * constHDSpaceWarningUmbral/constGigaByte)) or _
			(intObjRow.Status <> "OK") or _
			(intObjRow.VolumeDirty = True) Then
			strAlert= "Warning"
			bolAlertCheck= vbTrue
			if (Int(intObjRow.FreeSpace/constGigaByte) < Int(intObjRow.Size * constHDSpaceWarningUmbral/constGigaByte)) Then
				strStatusLogicalHD = strStatusLogicalHD & "The HD " & intObjRow.Name & " have low space. It has less than 10% of total disk space.  Please check." & VbCrLf 
			End If
			if (intObjRow.Status <> "OK") Then
				strStatusLogicalHD = strStatusLogicalHD & "The HD " & intObjRow.Name & " have an error.  Please check." & VbCrLf 
			End If
			If (intObjRow.VolumeDirty = True) Then
				strStatusLogicalHD = strStatusLogicalHD & "The HD " & intObjRow.Name & " is Dirty.  Please check." & VbCrLf 
			End If
		Else
			strAlert= ""    	
		End If
		If bolIsVerbose Then
			Wscript.echo " " _
						& F_strDataTrim(strAlert,8,0) & "|" _
						& F_strDataTrim(intObjRow.Status,13,0) & "|" _
						& F_strDataTrim(intObjRow.Name,5,0) & "|" _
						& F_strDataTrim(intObjRow.FileSystem,7,0) & "|" _
						& F_strDataTrim(arrDriveType(intObjRow.DriveType),18,0) & "|" _
						& F_strDataTrim(Int(intObjRow.Size/constGigaByte),8,1) & " Gb.|" _
						& F_strDataTrim(Int(intObjRow.FreeSpace/constGigaByte),12,1) & " Gb.|" _
						& F_strDataTrim(intObjRow.VolumeDirty,5,0)
		End If
	End If
Next
If not bolAlertCheck Then
	strStatusLogicalHD = "All the Logical Disks of the server are OK." & VbCrLf
Else
	strOveralStatus = "With Warnings"
End If 
bolAlertCheck= vbFalse
If bolIsVerbose Then
	Wscript.echo VbCrLf
End If


'=================================================================================
' SHARED INFORMATION
'=================================================================================
strAlert= ""
strTitle= " Action  |Status     |Name             |Path                                     |Type             "
strStatusShared= ""
If bolIsVerbose Then
	Wscript.echo F_AssemblingTitle("Shared  Information", len(strTitle))
	Wscript.echo strTitle
	Wscript.echo F_AssemblingTitle("",len(strTitle))
End If
Set objWMISet = objWMIService.ExecQuery("Select Name, Path, Type, Status from Win32_Share")
For Each intObjRow In objWMISet
	If intObjRow.Status <> "OK" Then
		strAlert= "Warning"
		bolAlertCheck= vbTrue
		strStatusShared= strStatusShared & intObjRow.Name & " (" & strShareType & ") with error.  Please check." & VbCrLf 
	Else
		strAlert= ""
	End If

	Select Case intObjRow.Type 
		Case 0 
			strShareType = "Disk Drive"
		Case 1
			strShareType = "Print Queue"
		Case 2
			strShareType = "Device"
		Case 3
			strShareType = "IPC"
		Case -2147483648
			strShareType = "Disk Drive Admin"
		Case -2147483649
			strShareType = "Print Queue Admin"
		Case -2147483650
			strShareType = "Device Admin"
		Case -2147483651
			strShareType = "IPC Admin"
	End Select
	If bolIsVerbose Then
		If (intObjRow.Type <> 1) and (intObjRow.Type <> -2147483649) Then
			Wscript.echo " " _
					& F_strDataTrim(strAlert,8,0) & "|" _
					& F_strDataTrim(intObjRow.Status,10,0) & " |" _
					& F_strDataTrim(intObjRow.Name,16,0) & " |" _
					& F_strDataTrim(intObjRow.Path,40,0) & " |" _
					& F_strDataTrim(strShareType,18,0)
		End If
	End If
Next
If not bolAlertCheck Then
	strStatusShared= "All the Shared of the server are OK." & VbCrLf
Else
	strOveralStatus = "With Warnings"
End If 
bolAlertCheck= vbFalse
If bolIsVerbose Then
	Wscript.echo VbCrLf
End If



'=================================================================================
' NIC INFORMATION
'=================================================================================
strAlert= ""
strTitle= " Action  |Status                   |Name                                     |MAC Address       |Speed      "
strStatusNIC= ""
If bolIsVerbose Then
	Wscript.echo F_AssemblingTitle("NIC Information", len(strTitle))
	Wscript.echo strTitle
	Wscript.echo F_AssemblingTitle("",len(strTitle))
End If
Set objWMISet = objWMIService.ExecQuery("Select *  from Win32_NetworkAdapter where NetConnectionStatus <> 'NULL'")
For Each intObjRow In objWMISet
	If arrConfigManagerErrorCode(intObjRow.ConfigManagerErrorCode) <> "Device is working properly" Then
		strAlert= "Warning"
		bolAlertCheck= vbTrue
		strStatusNIC= strStatusNIC & "The NIC " & intObjRow.Caption & " have the error " & arrConfigManagerErrorCode(intObjRow.ConfigManagerErrorCode) & ". Please check." & VbCrLf 
	Else
		strAlert= ""
	End If

	If bolIsVerbose Then
		Wscript.echo " " _
					& F_strDataTrim(strAlert,8,0) & "|" _
					& F_strDataTrim(arrNetConnectionStatus(intObjRow.NetConnectionStatus),24,0) & " |" _
					& F_strDataTrim(intObjRow.NetConnectionID,40,0) & " |" _
					& F_strDataTrim(intObjRow.MACAddress,17,0) & " |" _
					& F_strDataTrim(Int(intObjRow.Speed / constMegaByte) & " Mbps.",10,0)
	
	End If
Next
If not bolAlertCheck Then
	strStatusNIC= "All the NIC's of the server are OK." & VbCrLf
Else
	strOveralStatus = "With Warnings"
End If 
bolAlertCheck= vbFalse
If bolIsVerbose Then
	Wscript.echo VbCrLf
End If




'=================================================================================
' SERVICE INFORMATION
'=================================================================================
strAlert= ""
strTitle= " Action  |Status     |Service Name                             |State            |Start Mode"
strStatusService= ""
If bolIsVerbose Then
	Wscript.echo F_AssemblingTitle("Services Information",len(strTitle))
	Wscript.echo strTitle
	Wscript.echo F_AssemblingTitle("",len(strTitle))
End If
Set objWMISet = objWMIService.ExecQuery("Select DisplayName, State, StartMode, Status from Win32_Service")
For Each intObjRow In objWMISet
' General Service Check
	If intObjRow.Status <> "OK" Then
		strAlert= "Warning"
		bolAlertCheck= vbTrue
		strStatusService = strStatusService & "The service '" & intObjRow.DisplayName & "' has a status '" & intObjRow.Status & "'.  Please Check." & VbCrLf  
	End If
	If (intObjRow.StartMode = "Auto") And (intObjRow.State <> "Running") Then
		For intCounter=0 to uBound(arrServiceSkiped)
			If Trim(intObjRow.DisplayName) = arrServiceSkiped(intCounter) Then
				strAlert= "Skiped"
				strStatusService = strStatusService & "The service '" & intObjRow.DisplayName & "' has a start mode '" & intObjRow.StartMode & "' and now it is '" & intObjRow.State & "'.  It was skipped because it is a normal status." & VbCrLf  
				intCounter= uBound(arrServiceSkiped)
			End If
		Next
		If strAlert <> "Skiped" Then
			strAlert= "Warning"
			bolAlertCheck= vbTrue
			strStatusService = strStatusService & "The service '" & intObjRow.DisplayName & "' has a start mode '" & intObjRow.StartMode & "' and now it is '" & intObjRow.State & "'.  Please Check." & VbCrLf  
		End If
	End If

'arrRoleServices(5,0)= "Active Directory Domain Services" (Name of Rol)
'arrRoleServices(5,1)= "W2003 W2008" (Version of OS where the service is supported)
'arrRoleServices(5,2)= "Windows Time" (Services Name)
'arrRoleServices(5,3)= False (True= Service Enabled | False= Service Disabled)
'arrRoleServices(5,4)= False (True= Service Running | False= Service Not Running)

	For intCounter=0 to uBound(arrRoleServices,1)
		If (Trim(intObjRow.DisplayName) = arrRoleServices(intCounter,2)) And (InStr(arrRoleServices(intCounter,1),strOSVersion)>0) Then
			If (intObjRow.StartMode = "Auto") Or (intObjRow.StartMode = "Manual") Then
				arrRoleServices(intCounter,3)= True
				If (intObjRow.State = "Running") Then
					arrRoleServices(intCounter,4)= True
				End If
				intCounter= uBound(arrRoleServices,1)
			End If
		End If
	Next

	If bolIsVerbose Then
		Wscript.echo " " _
					& F_strDataTrim(strAlert,7,0) & " |" _
					& F_strDataTrim(intObjRow.Status,10,0) & " |" _
					& F_strDataTrim(intObjRow.DisplayName,40,0) & " |" _
					& F_strDataTrim(intObjRow.State,16,0) & " |" _
					& F_strDataTrim(intObjRow.StartMode,9,0)
	End If
	If strAlert <> "" Then
		strAlert= ""
	End If

Next
If not bolAlertCheck Then
	strStatusService = "All the Services of the server are OK." & VbCrLf 
Else
	strOveralStatus = "With Warnings"
End If 
bolAlertCheck = vbFalse
If bolIsVerbose Then
	Wscript.echo VbCrLf
End If

'=================================================================================
' PROCESS INFORMATION
'=================================================================================
strAlert= ""
strTitle= " Action    |S. Id |P. Id |% CPU |Process Name   |Executable Path                                              |Handles  |Non Paged|Peak Non |Paged    |Peak     |Memory Kb.   |Page File Usage Kb."
strTitle2= "           |      |      |      |               |                                                             |         |Pool     |Pged Pool|Pool     |Pged Pool|             |"
If bolIsVerbose Then
	Wscript.echo F_AssemblingTitle("Process Information",len(strTitle))
	Wscript.echo strTitle
	Wscript.echo strTitle2
	Wscript.echo F_AssemblingTitle("",len(strTitle))
End If
Set objWMISet = objWMIService.ExecQuery("Select * from Win32_Process")
For Each intObjRow In objWMISet
	If strOSVersion = "W2000" Then
		If bolIsVerbose Then
			Wscript.echo " " _
			           & F_strDataTrim(strAlert,9,0) & " |" _
			           & F_strDataTrim(intObjRow.SessionId,5,0) & " |" _
			           & F_strDataTrim(intObjRow.ProcessId,5,0) & " |" _
			           & F_strDataTrim("",5,0) & " |" _
			           & F_strDataTrim(intObjRow.Caption,14,0) & " |" _
			           & F_strDataTrim(intObjRow.ExecutablePath,60,0) & " |" _
			           & F_strDataTrim(intObjRow.HandleCount,8,0) & " |" _
			           & F_strDataTrim(intObjRow.QuotaNonPagedPoolUsage,8,0) & " |" _
			           & F_strDataTrim(intObjRow.QuotaPeakNonPagedPoolUsage,8,0) & " |" _
			           & F_strDataTrim(intObjRow.QuotaPagedPoolUsage,8,0) & " |" _
			           & F_strDataTrim(intObjRow.QuotaPeakPagedPoolUsage,8,0) & " |" _
			           & F_strDataTrim(intObjRow.WorkingSetSize,12,0) & " |" _
			           & F_strDataTrim(intObjRow.PageFileUsage,15,0)
		End If
	Else
		Set objWMISetNested = objWMIServiceNested.ExecQuery("Select * from Win32_PerfFormattedData_PerfProc_Process where IdProcess = " & intObjRow.ProcessId)
'		If objWMISetNested.Count > 0 Then
			For Each intObjRowNested In objWMISetNested
'				WScript.Echo "Handles: " & intObjRow.HandleCount & " - Max. Handles Supported: " & constProcHandleWarningUmbral 
				If (intObjRow.ProcessId <> 0) Then
					If (intObjRow.Caption <> "wmiprvse.exe") Then
						If intObjRow.HandleCount > constProcHandleWarningUmbral Then
							strAlert= "Warning H"
							bolAlertCheck = vbTrue
							strStatusProcess = strStatusProcess & "The process '" & intObjRow.Caption & "' has '" & intObjRow.HandleCount & "' handles in use.  Please Check." & VbCrLf  
						End If
						If int(intObjRowNested.PercentProcessorTime) > constProcPercWarningUmbral Then
							If strAlert = "" Then
								strAlert= "Warning %"
							Else
								strAlert= strAlert & "%"
							End If
							bolAlertCheck = vbTrue
							strStatusProcess = strStatusProcess & "The process '" & intObjRow.Caption & "' has '" & intObjRowNested.PercentProcessorTime & "'% of processor use.  Please Check." & VbCrLf  
						End If
					Else
						strAlert= "Skipped"
					End If
				End If
				If bolIsVerbose Then
					Wscript.echo " " _
					           & F_strDataTrim(strAlert,9,0) & " |" _
					           & F_strDataTrim(intObjRow.SessionId,5,0) & " |" _
					           & F_strDataTrim(intObjRow.ProcessId,5,0) & " |" _
				    	       & F_strDataTrim(intObjRowNested.PercentProcessorTime & "%",5,0) & " |" _
				        	   & F_strDataTrim(intObjRow.Caption,14,0) & " |" _
					           & F_strDataTrim(intObjRow.ExecutablePath,60,0) & " |" _
					           & F_strDataTrim(intObjRow.HandleCount,8,0) & " |" _
					           & F_strDataTrim(intObjRow.QuotaNonPagedPoolUsage,8,0) & " |" _
					           & F_strDataTrim(intObjRow.QuotaPeakNonPagedPoolUsage,8,0) & " |" _
					           & F_strDataTrim(intObjRow.QuotaPagedPoolUsage,8,0) & " |" _
				    	       & F_strDataTrim(intObjRow.QuotaPeakPagedPoolUsage,8,0) & " |" _
				        	   & F_strDataTrim(intObjRow.WorkingSetSize,12,0) & " |" _
					           & F_strDataTrim(intObjRow.PageFileUsage,15,0) & " |"
				End If
				strAlert= ""
			Next
'		End If
	End If
	strAlert= ""
Next
If not bolAlertCheck Then
	strStatusProcess= "All the Process of the server are OK." & VbCrLf 
Else
	strStatusProcess= "Some process passed the threshold of handles or percent of use of processor defined in the script." & VbCrLf _
					& "Some time this is a normal situation, some time this may be a problem.  Please evaluate the situation." & VbCrLf & strStatusProcess
	strOveralStatus = "With Warnings"
End If 
bolAlertCheck= vbFalse
If bolIsVerbose Then
	Wscript.echo VbCrLf
End If

'=================================================================================
' WINDOWS ROLE DETECTION
'=================================================================================
'=================================================================================
' ROLE PRINT SERVICES
'=================================================================================

If F_bolWindowsRole(strOSVersion, "Print Services") Then
	strAlert= ""
	strTitle= " Action     |Name                           |Status                      |Error             |Driver Name                              |Port Name          |Print Processor |Share |Shared Name     |Location                                "
	strStatusPrintServices= ""
	If bolIsVerbose Then
		Wscript.echo F_AssemblingTitle("Print Information", len(strTitle))
		Wscript.echo strTitle
		Wscript.echo F_AssemblingTitle("",len(strTitle))
	End If

	Set objWMISet = objWMIService.ExecQuery("Select * from Win32_Printer")
	For Each intObjRow In objWMISet
		If (intObjRow.ExtendedPrinterStatus = 9) OR (LCase(intObjRow.PrintProcessor) <> "winprint") And (LCase(intObjRow.PrintProcessor) <> "Citrix Print Processor") Then
			bolAlertCheck= vbTrue
			If (intObjRow.ExtendedPrinterStatus = 9) Then
				strAlert= "Warning"
				strStatusPrintServices= strStatusPrintServices & "The printer '" & intObjRow.Name & "' has an error.  Please check." & VbCrLf 
			Else
				strAlert= "Warning PSF"
				strStatusPrintServices= strStatusPrintServices & "The printer '" & intObjRow.Name & "' has a print server distinct to WinPrint.  It is recommended to change the print server to WinPrint." & VbCrLf 
			End If
		Else
			strAlert= ""
		End If
		If bolIsVerbose Then
			Wscript.echo " " _
					& F_strDataTrim(strAlert,11,0) & "|" _
					& F_strDataTrim(intObjRow.Name,30,0) & " |" _
					& F_strDataTrim(arrExtendedPrinterStatus(intObjRow.ExtendedPrinterStatus),27,0) & " |" _
					& F_strDataTrim(arrPrinterError(intObjRow.DetectedErrorState),17,0) & " |" _
					& F_strDataTrim(intObjRow.DriverName,40,0) & " |" _
					& F_strDataTrim(intObjRow.PortName,18,0) & " |" _
					& F_strDataTrim(intObjRow.PrintProcessor,15,0) & " |" _
					& F_strDataTrim(intObjRow.Shared,5,0) & " |" _
					& F_strDataTrim(intObjRow.ShareName,15,0) & " |" _
					& F_strDataTrim(intObjRow.Location,40,0)

		End If
	Next
	If not bolAlertCheck Then
		strStatusPrintServices= "All the Printer of the server are OK." & vbCrLf
	End If 
	bolAlertCheck= vbFalse
	If bolIsVerbose Then
		Wscript.echo VbCrLf
	End If
Else
	strStatusPrintServices= "The server isn't Print Server."
End If


'=================================================================================
' ROLE ACTIVE DIRECTORY DOMAIN SERVICES
' This script check essentials things that are essential to the proper operation
' of Active Directory in a server.  The checks are based on Active Directory 
' Management Pack Scripts.
' 
' - Check Replication Partners Status
' - Check Replication Operations Pending
' - Check Status of Domain Trust
' - Check essentials services (File Replication, Intersite Messaging, Kerberos Key Distribution Center, Net Logon and Windows Time)
' - Check availability of SYSVOL share
' http://technet.microsoft.com/en-us/library/cc180916.aspx
' http://www.activexperts.com/network-monitor/windowsmanagement/adminscripts/monitoring/ad/
'=================================================================================

strStatusAD= ""
If F_bolWindowsRole(strOSVersion, "Active Directory Domain Services") Then
	If bolIsVerbose Then
		Wscript.echo F_AssemblingTitle("Active Directory Role information",100)
		Wscript.echo F_AssemblingTitle("",100) & VbCrLf
		strAlert= ""
		strTitle= " Action  |Domain                              |Naming context DN                                         |Source DSA DN                                               |Last Synch |Number of consecutive synchronization failures"
		Wscript.echo F_AssemblingTitle("Active Directory Replication Partners Status",len(strTitle))
		Wscript.echo strTitle
		Wscript.echo F_AssemblingTitle("", Len(strTitle))
	End If
'*** Check Replication Partners ***
	Set objWMIADService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputerDestination & "\root\MicrosoftActiveDirectory")
	Set objWMIADSet= objWMIADService.ExecQuery ("Select * from MSAD_ReplNeighbor")
    For each intObjRow in objWMIADSet
		If (intObjRow.LastSyncResult <> 0) or (intObjRow.NumConsecutiveSyncFailures <> 0) Then
			strAlert= "Warning"
		    bolAlertCheck= vbTrue
		Else
			strAlert= ""
		End If
		If bolIsVerbose Then
			Wscript.echo " " _
			           & F_strDataTrim(strAlert,7,0) & " |" _
			           & F_strDataTrim(intObjRow.Domain,35,0) & " |" _
			           & F_strDataTrim(intObjRow.NamingContextDN,57,0) & " |" _
			           & F_strDataTrim(intObjRow.SourceDsaDN,59,0) & " |" _
			           & F_strDataTrim(intObjRow.LastSyncResult,10,0) & " |" _
			           & F_strDataTrim(intObjRow.NumConsecutiveSyncFailures,14,0)
		End If
    Next
	If not bolAlertCheck Then
		strStatusADRepPartners= "  All the Active Directory replication with partners are OK." & VbCrLf 
	Else
		strStatusADRepPartners= "  Some Active Directory replication with partners has error." & VbCrLf _
								& "  Some time this is a normal situation, some time this may be a problem." & VbCrLf _
								& "  Please check Active Directory health with REPADMIN or DCDIAG commands." & VbCrLf
	End If 
	bolAlertCheck= vbFalse

'*** Check Replication Status ***
	If bolIsVerbose Then
		Wscript.echo VbCrLf
		strAlert= ""
		strTitle= "Replication Serial Number |Time Enqueued |DSA DN        |DSA Adress    |Naming Context DN"
		Wscript.echo F_AssemblingTitle("Active Directory Replication Status",len(strTitle) + 62)
	End If
	Set objWMIADSet= objWMIADService.ExecQuery ("Select * from MSAD_ReplPendingOp")
	If objWMIADSet.Count = 0 Then
		If bolIsVerbose Then
			Wscript.echo "There aren't replication jobs pending." & VbCrLf
		End If
	Else
		If bolIsVerbose Then
			Wscript.echo strTitle
			Wscript.echo F_AssemblingTitle("", Len(strTitle))
		    For each intObjRow in objWMIADSet
				strAlert= "Warning"
				bolAlertCheck= vbTrue
				Wscript.echo " " _
				           & F_strDataTrim(strAlert,9,0) & " |" _
				           & F_strDataTrim(intObjRow.SerialNumber,26,0) & " |" _
				           & F_strDataTrim(intObjRow.TimeEnqueued,14,0) & " |" _
				           & F_strDataTrim(intObjRow.DsaDN,14,0) & " |" _
				           & F_strDataTrim(intObjRow.DsaAddress,14,0) & " |" _
				           & F_strDataTrim(intObjRow.NamingContextDn,14,0)
		    Next
		End If
	End If
	If not bolAlertCheck Then
		strStatusADReplication= "  All the Active Directory replication jobs are OK." & VbCrLf 
	Else
		strStatusADReplication= "  Some Active Directory replication are pending." & VbCrLf _
								& "  Some time this is a normal situation, some time this may be a problem." & VbCrLf _
								& "  Please check Active Directory health with REPADMIN or DCDIAG commands." & VbCrLf
	End If 
	bolAlertCheck= vbFalse

'*** Check Trust Status ***
	If bolIsVerbose Then
		Wscript.echo VbCrLf
		strTitle= " Action  |Domain Controller Name     |Domain                      |Trust Direction  |Status |Status String"
		Wscript.echo F_AssemblingTitle("Active Directory Trust Status",len(strTitle) + 62)
		Wscript.echo strTitle
		Wscript.echo F_AssemblingTitle("",len(strTitle) + 62)
	End If
	Set objWMIADSet= objWMIADService.ExecQuery("SELECT * FROM Microsoft_DomainTrustStatus")
	For Each intObjRow in objWMIADSet
		If (intObjRow.TrustStatus = 0) Then
			strAlert= ""
			If bolIsVerbose Then
				Wscript.echo " " _
				           & F_strDataTrim(strAlert,7,0) & " |" _
				           & F_strDataTrim(intObjRow.TrustedDCName,26,0) & " |" _
				           & F_strDataTrim(intObjRow.TrustedDomain,27,0) & " |" _
				           & F_strDataTrim(arrTrustDirection(intObjRow.TrustDirection - 1),16,0) & " |" _
				           & F_strDataTrim(intObjRow.TrustIsOK,6,0) & " |" _
				           & F_strDataTrim(intObjRow.TrustStatus & " - " & intObjRow.TrustStatusString,len(intObjRow.TrustStatusString) + 10,0)
			End If
		Else
			strAlert= "Warning"
			bolAlertCheck= vbTrue
			If bolIsVerbose Then
				Wscript.echo " " _
				           & F_strDataTrim(strAlert,7,0) & " |" _
				           & F_strDataTrim(intObjRow.TrustedDCName,26,0) & " |" _
				           & F_strDataTrim(intObjRow.TrustedDomain,27,0) & " |" _
				           & F_strDataTrim(arrTrustDirection(intObjRow.TrustDirection - 1),16,0) & " |" _
				           & F_strDataTrim(intObjRow.TrustIsOK,6,0) & " |" _
				           & F_strDataTrim(intObjRow.TrustStatus & " - " & intObjRow.TrustStatusString,len(intObjRow.TrustStatusString) + 10,0)
			End If
		End If	
	Next
	If not bolAlertCheck Then
		strStatusADTrust= "  All the Active Directory trust are OK." & VbCrLf 
	Else
		strStatusADTrust= "  Some Active Directory trust has error.  Please check Active Directory trust health." & VbCrLf
	End If 
	bolAlertCheck= vbFalse
	
'*** Check SYSVOL Status ***
	If bolIsVerbose Then
		Wscript.echo VbCrLf
		Wscript.echo F_AssemblingTitle("Active Directory SYSVOL Status",len(strTitle))
	End If
	Set objWMISet = objWMIService.ExecQuery("Select Name, Path, Type, Status from Win32_Share Where Name = 'SYSVOL' and Type = '0' and Status = 'OK'")
	If objWMISet.Count = 0 Then
		If bolIsVerbose Then
			Wscript.echo "  There is problem with the SYSVOL shared.  This shared is essential to Domain Controller servers." & VbCrLf
		End If
		strStatusADSYSVOL= "  There is problem with the SYSVOL shared.  This shared is essential to Domain Controller servers." & VbCrLf
	Else
		If bolIsVerbose Then
			Wscript.echo "  The SYSVOL shared is OK." & VbCrLf
		End If
		strStatusADSYSVOL= "  The SYSVOL shared is OK." & VbCrLf
	End If
	Wscript.echo VbCrLf


'*** Check Services ***
'arrRoleServices(5,0)= "Active Directory Domain Services" (Name of Rol)
'arrRoleServices(5,1)= "W2003 W2008" (Version of OS where the service is supported)
'arrRoleServices(5,2)= "Windows Time" (Services Name)
'arrRoleServices(5,3)= False (True= Service Enabled | False= Service Disabled)
'arrRoleServices(5,4)= False (True= Service Running | False= Service Not Running)

	If bolIsVerbose Then
		Wscript.echo VbCrLf
		Wscript.echo F_AssemblingTitle("Active Directory essentials Services Status",len(strTitle))
	End If
	For intCounter=0 to uBound(arrRoleServices,1)
		If arrRoleServices(intCounter,0) = "Active Directory Domain Services" and InStr(arrRoleServices(intCounter,1),strOSVersion) > 0 Then
			If arrRoleServices(intCounter,3) Then
				If arrRoleServices(intCounter,4) Then
					If bolIsVerbose Then
						Wscript.echo "The Service " & arrRoleServices(intCounter,2) & " is running."
					End If
				Else
					If bolIsVerbose Then
						Wscript.echo "The Service " & arrRoleServices(intCounter,2) & " isn't running."
					End If
					bolAlertCheck= vbTrue
				End If
			Else
				bolAlertCheck= vbTrue
				If bolIsVerbose Then
					Wscript.echo "The Service " & arrRoleServices(intCounter,2) & " is disabled, please check."
				End If
			End If
		End If
	Next
	If not bolAlertCheck Then
		strStatusADServices= "  All Active Directory essentials Services are OK." & VbCrLf 
	Else
		strStatusADServices= "Some Active Directory essentials Services aren't running.  Please check Active Directory essentials Services." & VbCrLf
		strOveralStatus = "With Warnings"
	End If 

Else
	strStatusAD= "The server isn't Domain Controller."
End If

Wscript.echo

'=================================================================================
' FAIL OVER CLUSTERING
'=================================================================================
'If F_bolWindowsRole(strOSVersion, "Failover Clustering") Then
'	Wscript.echo "Failover Clustering"
'Else
'	Wscript.echo "Non Failover Clustering"
'End If
'Wscript.echo


'=================================================================================
' NON WINDOWS ROLE DETECTION
'=================================================================================
'=================================================================================
' CITRIX
'=================================================================================

'NAMESPACE: CITRIX
'CITRIX Services
'         |OK         |Citrix Diagnostic Facility COM Server    |Running          |Auto     
'         |OK         |Citrix Client Network                    |Running          |Auto     
'         |OK         |Citrix Encryption Service                |Running          |Auto     
'         |OK         |Citrix SMA Service                       |Running          |Auto     
'         |OK         |Citrix Virtual Memory Optimization       |Stopped          |Manual   
'         |OK         |Citrix Health Monitoring and Recovery    |Running          |Auto     
'         |OK         |Citrix XTE Server                        |Stopped          |Manual   
'         |OK         |Citrix Print Manager Service             |Running          |Auto     
'         |OK         |Citrix ActiveSync Service                |Running          |Auto     
'         |OK         |Citrix CPU Utilization Mgmt/CPU Rebal... |Stopped          |Manual   
'         |OK         |Citrix CPU Utilization Mgmt/Resource ... |Stopped          |Manual   
'         |OK         |Citrix XML Service                       |Running          |Auto     
'         |OK         |Citrix Services Manager                  |Running          |Auto     
'         |OK         |Citrix Independent Management Archite... |Running          |Auto     
'         |OK         |Citrix MFCOM Service                     |Running          |Auto     

Wscript.echo


'=================================================================================
' VMWARE
'=================================================================================

'A provider, VMwareStatsProvider_v1, has been registered in the WMI namespace, Root\CimV2

'=================================================================================
' Overall Status
'=================================================================================
Wscript.echo "OVERALL STATUS: " & strOveralStatus
Wscript.echo "[MEMORY STATUS]"
Wscript.echo strStatusSO
Wscript.echo ""
Wscript.echo "[REBOOT STATUS]"
Wscript.echo strStatusEventMessage
Wscript.echo ""
Wscript.echo "[LOGICAL DISK STATUS]"
Wscript.echo strStatusLogicalHD
Wscript.echo ""
Wscript.echo "[SHARED STATUS]"
Wscript.echo strStatusShared
Wscript.echo ""
Wscript.echo "[NIC STATUS]"
Wscript.echo strStatusNIC
Wscript.echo ""
WScript.echo "[SERVICES STATUS]"
Wscript.echo strStatusService 
Wscript.echo ""
Wscript.echo "[PROCESS STATUS]"
Wscript.echo strStatusProcess 
Wscript.echo ""
WScript.echo "[PRINT SERVER STATUS]"
Wscript.echo strStatusPrintServices
Wscript.echo ""
WScript.echo "[ACTIVE DIRECTORY STATUS]"
if strStatusAD = "" Then
	Wscript.echo strStatusADRepPartners
	Wscript.echo strStatusADReplication
	Wscript.echo strStatusADTrust
	Wscript.echo strStatusADSYSVOL
	Wscript.echo strStatusADServices
Else
	Wscript.echo strStatusAD
End If


Wscript.Quit(intExitCode)


' *****************************************************************************************************
' *    Name: FWMI_OSVersion
' *    Type: Function
' *   Input: Caption from Win32_OperatingSystem
' *  Output: String with W2003 or W2008
' * Purpose: 
' *    Note: N/A
' *****************************************************************************************************
Function FWMI_OSVersion(lstrOSVersion)
Dim strOSVersionResult

lstrOSVersion= Trim(lstrOSVersion)
Select Case lstrOSVersion
	Case "Microsoft(R) Windows(R) Server 2003, Standard Edition"
		strOSVersionResult= "W2003"
	Case "Microsoft(R) Windows(R) Server 2003, Enterprise Edition"
		strOSVersionResult= "W2003"
	Case "Microsoft Windows 7 Enterprise"
		strOSVersionResult= "W7"
	Case "Microsoft Windows Server 2008 R2 Standard"
		strOSVersionResult= "W2008"
	Case "Microsoft Windows Server 2008 R2 Enterprise"
		strOSVersionResult= "W2008"
	Case Else
		strOSVersionResult= "W2000"
End Select
FWMI_OSVersion= strOSVersionResult
End Function


'********************************************************************
'*
'*    Name: F_AssemblingTitle
'*    Type: Function
'* Purpose: assembling the separate title for each component
'*   Input: strTitle: string
'*  Output: Title with format
'*
'********************************************************************
Private Function F_AssemblingTitle(ByVal strTitle, ByVal intLenght)
Dim strLen, strFill

	F_AssemblingTitle= "+" & string(intLenght,"-") & "+"
	If strTitle <> "" Then
		strLen= Len(strTitle)
		strFill= int((intLenght-2-strLen)/2)
		F_AssemblingTitle= "+" & string(strFill,"-") & " " & strTitle & " " & string(strFill,"-") & "+"
	End If
End Function


'********************************************************************
'*
'*    Name: F_strDataTrim
'*    Type: Function
'* Purpose: Trim data in a specific long
'*   Input: strData: string, intSize: integer, intAlign: integer (0- Left, Other: Right)
'*  Output: Title with format
'*
'********************************************************************
Private Function F_strDataTrim(ByVal strData, ByVal intSize, ByVal intAlign)
	If Len(strData) > intSize Then
		strData= Left(strData,intSize-3) & "..."
	End If
	If IsNull(strData) Then
		F_strDataTrim= String(intSize, " ")
	Else
		If intAlign = 0 Then
			F_strDataTrim= Left(strData,intSize) & String(intSize - Len(Trim(Left(strData,intSize + 1)))," ")
		Else
			F_strDataTrim= String(intSize - Len(Trim(Left(strData,intSize + 1)))," ") & Left(strData,intSize)
		End If
	End If
End Function


'********************************************************************
'*
'*    Name: F_bolWindowsRole
'*    Type: Function
'* Purpose: Detect if certain role is installed on server
'*   Input: strWindowsVersion: string, strWindowsRole: string
'*  Output: True - False
'*
'********************************************************************
'arrRoleServices(5,0)= "Active Directory Domain Services" (Name of Rol)
'arrRoleServices(5,1)= "W2003 W2008" (Version of OS where the service is supported)
'arrRoleServices(5,2)= "Windows Time" (Services Name)
'arrRoleServices(5,3)= False (True= Service Enabled | False= Service Disabled)
'arrRoleServices(5,4)= False (True= Service Running | False= Service Not Running)

Private Function F_bolWindowsRole(ByVal strWindowsVersion, ByVal strWindowsRole)
Dim objRolWMISet, objRolWMIService, intRolObjRow

	F_bolWindowsRole= False
	Select Case strWindowsVersion
		Case "W2003", "W2000"
			Select Case strWindowsRole
				Case "Active Directory Domain Services"
					F_bolWindowsRole= False
					If (strADComputerRole = "Primary Domain Controller") or (strADComputerRole = "Backup Domain Controller") Then
						F_bolWindowsRole= True
					End If

				Case "Print Services"
					For intCounter=0 to uBound(arrRoleServices,1)
						If (arrRoleServices(intCounter,0) = strWindowsRole) And (InStr(arrRoleServices(intCounter,1),strWindowsVersion)>0) Then
							If arrRoleServices(intCounter,3) Then
								F_bolWindowsRole= True
							End If
						End If
					Next

				Case "Windows Server Backup"
					If strWindowsRole = intRolObjRow.Name Then
						F_bolWindowsRole= True
					End If
				
				Case "Windows Server Backup"

				Case "Failover Clustering"
					
				Case Else 
					F_bolWindowsRole= False
			End Select

		Case "W2008"
			Set objRolWMIService = GetObject("winmgmts:{impersonationLevel=Impersonate}!\\" & "." & "\root\cimv2")
			Set objRolWMISet = objRolWMIService.ExecQuery("SELECT Name FROM Win32_ServerFeature where Name = '" & strWindowsRole & "'")
			For Each intRolObjRow In objRolWMISet
				Select Case strWindowsRole
					Case "Active Directory Domain Services"
						If strWindowsRole = intRolObjRow.Name Then
							F_bolWindowsRole= True
						End If

					Case "Windows Server Backup"
						If strWindowsRole = intRolObjRow.Name Then
							F_bolWindowsRole= True
						End If

					Case "Failover Clustering"
						If strWindowsRole = intRolObjRow.Name Then
							F_bolWindowsRole= True
						End If
					
					Case Else 
						F_bolWindowsRole= False
				End Select
			Next
		Case Else 
			F_bolWindowsRole= False
	End Select
End Function


'OVO
'OVO Configuration
'HKLM\SOFTWARE\Hewlett-Packard\HP OpenView
'DataDir (Reg_SZ)
'InstallDir (Reg_SZ)
'
'OVO ITO Version
'HKLM\SOFTWARE\Hewlett-Packard\HP OpenView\{95E8BC5C-C38D-42E0-9983-FE70FCE81FA2}
'ProductVersion (Reg_SZ)
'HKLM\SOFTWARE\Hewlett-Packard\OpenView\ITO
'Agent Version (Reg_SZ)
'HKLM\SOFTWARE\Wow6432Node\Hewlett-Packard\OpenView\ITO
'Agent Version (Reg_SZ)
'
'Archivo de Configuracion
'C:\osit\etc\ 
'dhcp_mon.cfg
'lp_mon.cfg
'perf_mon.cfg
'tsk_mon
'cert_mon
'srv_mon
'df_mon
'ps_mon
'act_mon
'
'


'********************************************************************
'*
'*    Name: F_GetOpenViewVersion
'*    Type: Function
'* Purpose: Retrieve OVO Version
'*   Input: N/A
'*  Output: OVO Version
'*          
'********************************************************************
Private Function F_GetOpenViewVersion
Dim strResult

F_ReadRegistryKeyValue "HKLM\SOFTWARE\Hewlett-Packard\HP OpenView\{95E8BC5C-C38D-42E0-9983-FE70FCE81FA2}\ProductVersion", "REG_SZ", strResult
If (strResult = "") Then
	F_ReadRegistryKeyValue "HKLM\SOFTWARE\Hewlett-Packard\OpenView\ITO\Agent Version", "REG_SZ", strResult
	If (strResult = "") Then
		F_ReadRegistryKeyValue "HKLM\SOFTWARE\Wow6432Node\Hewlett-Packard\OpenView\ITO\Agent Version", "REG_SZ", strResult
	End If
End If
If (strResult = "") Then
	F_GetOpenViewVersion= "Was impossible to detect OVO version"
Else
	F_GetOpenViewVersion= strResult
End If
End Function


'********************************************************************
'*
'*    Name: F_ReadRegistryKeyValue
'*    Type: Function
'* Purpose: Retrieve registry key values
'*   Input: strKey: string
'*          byRef strRegistryValue: string
'*  Output: strRegistryValue: Registry key value
'*          F_ReadRegistryKeyValue: vbTrue | vbFalse
'*
'********************************************************************
Private Function F_ReadRegistryKeyValue(strKey, strType, byRef strRegistryValue)
Const constRegistryError= &h80070002
Const constRegistrySuccess= 0
Dim ErrDescription
Dim strREG_SZKeyValue		:strREG_SZKeyValue= ""

On Error Resume Next
Select Case strType
	Case "REG_SZ"
		strREG_SZKeyValue = CreateObject("WScript.Shell").RegRead(strKey)
		Select Case Err
			Case constRegistrySuccess
				F_ReadRegistryKeyValue= vbTrue
				strRegistryValue= strREG_SZKeyValue
			Case constRegistryError
				ErrDescription= Replace(Err.description, strKey, "")
				Err.clear
				CreateObject("WScript.Shell").RegRead("HKEY_ERROR\")
				F_ReadRegistryKeyValue= (ErrDescription <> Replace(Err.description, "HKEY_ERROR\", ""))
			Case Else
				F_ReadRegistryKeyValue= vbFalse
		End Select
		Err.Clear
	Case Else
End Select
End Function
