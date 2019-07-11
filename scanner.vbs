' ********************************************************************
' Script	:	scanner.vbs
' Purpose	:	Read in an XML file that contains verification tasks and check these tasks
' 				against a remote machine. 
' Author	:	Michael van Doren - Application Integration
' Modified	:	Ben Paltridge - Technical Consultant
' Date		:	1.0 Relased 21 july 2005
'			2.0 Released 12 july 2019
' ********************************************************************

On Error Resume Next

Dim Config, oFSO, oShell, LogFile, oOptions, Output_Directory
Dim Events, Summary, Root, TimeScriptStarted
Dim SuccessPoints, FailurePoints, OutputFileName
Dim oSWbemServices, oWbemLocator, oRegistry
Dim NetworkStatus, IPAddress, TimeStarted, ComputerName, Category
Dim oSWbemServicesDefault
Dim oNetworkStatus, oIPAddress,oStatus,oWave,oType,oSuccessPoints,oFailurePoints,oFailureSummary
Dim ast, bts, cfx, gen, uno
Dim g_Temp, ControlFileName
DIm g_WSUser
Dim g_WSPass

Dim g_LastSuccessfulUsername
Dim g_LastSuccessfulPassword
Dim g_IntegratedAuthentication

' Constants
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_DYN_DATA = &H80000006

' Find a home for this
g_IntegratedAuthentication = True

Const FOR_READING = 1
' Output_Directory = Replace(WScript.ScriptFullName,WScript.scriptname,"") &  "Output\"

Const adOpenStatic = 3
Const adLockOptimistic = 3

Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oShell = CreateObject("WScript.Shell")
Set oNetwork = CreateObject("WScript.Network")
Set oOptions = CreateObject("Scripting.Dictionary")
TimeScriptStarted = Now()

Dim g_CurrentDirectory : g_CurrentDirectory = AddBackslash(oShell.CurrentDirectory)
Dim Config_File ' : Config_File = g_CurrentDirectory & "config.xml"

g_Temp = AddBackslash(oShell.ExpandEnvironmentStrings("%TEMP%"))
UserDataXML = WScript.Arguments.Named ("userdataxml")
DataXML = WScript.Arguments.Named("dataxml")
ComputerName = uCase(WScript.Arguments.Named("comp"))
Config_File = UCase(WScript.Arguments.Named("config"))
InitialNode = WScript.Arguments.Named("node")
Field1 = ""
Field2 = ""
Field3 = ""

OutputFileName = OUTPUT_DIRECTORY & ComputerName & ".xml"
ControlFileName = OUTPUT_DIRECTORY & "\Control\" & ComputerName & ".XML"

TimeStarted = Now()
NetworkStatus = ""
IPAddress = ""
Status = ""
SuccessPoints = ""
FailurePoints = ""


' Load DataXML
Set oDataXML = CreateObject("Microsoft.XMLDOM")
oDataXML.async = False
oDataXML.validateonparse = False
DataXML = Replace(dataXML,"~","""")
If oDataXML.loadXML(dataXML) = False Then	
	Msgbox "Failed to open dataxml"
	WScript.Quit 
End If

' Load UserDataXML
Set oUserDataXML = CreateObject("Microsoft.XMLDOM")
oUserDataXML.async = False
oUserDataXML.validateonparse = False
UserDataXML = Replace(userdataXML,"~","""")
If oUserDataXML.loadXML(userdataXML) = False Then	
	MsgBox "Failed to open: " &  UserDataXML
	WScript.Quit
End If


InitializeLogFile(OutputFileName)

Set oWbemLocator = CreateObject("WbemScripting.SWbemLocator")

' Open Config File
If OpenConfigFile = False Then
	' UpdateHeaderInformation(OutputFileName)
	msgbox "Failed to open config file"
	WriteLog "openconfigfile","error connecting to file","failure","check your config file"
	Quit
End If

Field1 = ""
Field2 = ""

SuccessPoints = CInt(0)
FailurePoints = CInt(0)




' Let the processing begin
ProcessVerificationTasks(InitialNode)

Quit

' ############################################
' Functions
' ############################################

Function ProcessVerificationTasks(verificationNode)
	On Error Resume Next
	Dim VerificationTask, VerificationTasks
	ProcessVerificationTasks = False
	If Len(verificationNode) < 1 Then
		Exit Function
	End If

	Set VerificationTasks = Config.DocumentElement.SelectSingleNode(verificationNode)
	For Each VerificationTask In VerificationTasks.ChildNodes
		if len(VerificationTask.NodeName) < 1 then
			ProcessVerificationTasks = False
		else
			ProcessVerificationNode VerificationTask
			ProcessVerificationTasks = True
		End If 
	Next
End Function

Sub ProcessVerificationNode(VerificationTask)
	On Error Resume Next
	Dim Name, Version, Status, Startup, DateCreated, KeyName, KeyValue, ErrorIfKeyDoesNotExist,KeyType
	Dim FileName, Domain, UserName, Mask, DriveLetter, Force
	Select Case VerificationTask.nodeName
		Case "executenodes"
			ProcessVerificationTasks(VerificationTask.GetAttributeNode("query").Value)
		Case "fileexists"
			Set Name = VerificationTask.GetAttributeNode("name")
			FileExists Name.Value
		Case "fileversion"
			Set Name = VerificationTask.GetAttributeNode("name")
			Set Version = VerificationTask.GetAttributeNode("version")
			FileVersion Name.Value,Version.Value
		Case "filedatecreated"
			Set Name = VerificationTask.GetAttributeNode("name")
			Set DateCreated = VerificationTask.GetAttributeNode("datecreated")
			FileDateCreated Name.Value,DateCreated.Value
		Case "filedatemodified"
			Set Name = VerificationTask.GetAttributeNode("name")
			Set DateModified = VerificationTask.GetAttributeNode("datemodified")
			FileDateModified Name.Value,DateModified.Value
		Case "folderexists"
			Set Name = VerificationTask.GetAttributeNode("name")
			FolderExists Name.Value
		Case "foldersize"
			Set Name = VerificationTask.GetAttributeNode("name")
			FolderSize Name.Value
		Case "filesize"
			Set Name = VerificationTask.GetAttributeNode("name")
			Set Size = VerificationTask.GetAttributeNode("size")
			FileSize Name.Value, Size.Value
		Case "registrykeyvalue"
			Set Name = VerificationTask.GetAttributeNode("name")
			Set KeyType = VerificationTask.GetAttributeNode("keytype")
			Set KeyValue = VerificationTask.GetAttributeNode("keyvalue")
			RegistryKeyValue Name.Value,KeyType.Value,KeyValue.Value
		Case "registrykeyvalueset"
			Set Name = VerificationTask.GetAttributeNode("name")
			Set KeyType = VerificationTask.GetAttributeNode("keytype")
			Set KeyValue = VerificationTask.GetAttributeNode("keyvalue")
			RegistryKeyValueset Name.Value,KeyType.Value,KeyValue.Value
		Case "registrykeymulticontains"
			Set KeyName = VerificationTask.GetAttributeNode("name")
			Set KeyValue = VerificationTask.GetAttributeNode("keyvalue")
			RegistryKeyMultiContains KeyName.Value,KeyValue.Value
		Case "registrykeyexistserror"
			registrykeyexistserror VerificationTask
		Case "setdatabasevalue"
			SetDatabaseValue VerificationTask
		Case "registrykeyvalueerror"
			Set KeyName = VerificationTask.GetAttributeNode("name")
			Set KeyType = VerificationTask.GetAttributeNode("keytype")
			Set KeyValue = VerificationTask.GetAttributeNode("keyvalue")
			Set ErrorIfKeyDoesNotExist = VerificationTask.GetAttributeNode("errorifkeydoesnotexist")
			RegistryKeyValueError KeyName.Value,KeyType.Value,KeyValue.Value, ErrorIfKeyDoesNotExist.Value
		Case "servicestate"
			Set Name = VerificationTask.GetAttributeNode("name")
			Set State = VerificationTask.GetAttributeNode("state")
			ServiceState Name.Value,State.Value
		Case "pingcomputer"
			PingComputer VerificationTask
		Case "servicestop"
			Set Name = VerificationTask.GetAttributeNode("name")
			ServiceStop Name.Value
		Case "servicestart"
			Set Name = VerificationTask.GetAttributeNode("name")
			ServiceStart Name.Value
		Case "servicestartmode"
			Set Name = VerificationTask.GetAttributeNode("name")
			Set StartMode = VerificationTask.GetAttributeNode("startmode")
			ServiceStartMode Name.Value,StartMode.Value
		Case "servicestartmodeerror"
			Set Name = VerificationTask.GetAttributeNode("name")
			Set StartMode = VerificationTask.GetAttributeNode("startmode")
			ServiceStartModeError Name.Value,StartMode.Value
		Case "servicesetstartmode"
			Set Name = VerificationTask.GetAttributeNode("name")
			Set StartMode = VerificationTask.GetAttributeNode("startmode")
			ServiceSetStartMode Name.Value,StartMode.Value
		Case "databaseconnection"
			Set Name = VerificationTask.GetAttributeNode("name")
			DatabaseConnection Name.Value
		Case "scheduledtaskexists"
			Set Name = VerificationTask.GetAttributeNode("name")
			ScheduledTaskExists Name.Value
		Case "shareexists"
			Set Name = VerificationTask.GetAttributeNode("name")
			ShareExists Name.Value
		Case "createprocess"
			Set Name = VerificationTask.GetAttributeNode("name")
			CreateProcess Name.Value
		Case "processkill"
			Set Name = VerificationTask.GetAttributeNode("name")
			ProcessKill Name.Value
		Case "processrunning"
			Set Name = VerificationTask.GetAttributeNode("name")
			Set OnlyIfUserIsLoggedOn = Verificationtask.GetAttributeNode("onlyifuserisloggedon")
			ProcessRunning Name.Value, OnlyIfUserIsLoggedOn.Value
		Case "processrunningerror"
			Set Name = VerificationTask.GetAttributeNode("name")
			Set OnlyIfUserIsLoggedOn = Verificationtask.GetAttributeNode("onlyifuserisloggedon")
			ProcessRunningError Name.Value, OnlyIfUserIsLoggedOn.Value
		Case "filecontainsstring"
			FileContainsString VerificationTask.GetAttributeNode("filename").Value, VerificationTask.GetAttributeNode("searchstring").Value
		Case "fileshowstringaftermatch"
			fileshowstringaftermatch VerificationTask.GetAttributeNode("filename").Value, VerificationTask.GetAttributeNode("searchstring").Value
		Case "filereadline"
			filereadline VerificationTask.GetAttributeNode("filename").Value, VerificationTask.GetAttributeNode("searchstring").Value
		Case "filecontainserrorstring"
			Set Name = VerificationTask.GetAttributeNode("filename")
			Set SearchString = VerificationTask.GetAttributeNode("searchstring")
			FileContainsErrorString Name.Value, SearchString.Value
		Case "environmentvariableexists"
			Set Name = VerificationTask.GetAttributeNode("name")
			EnvironmentVariableExists Name.Value
		Case "smspolicyretrieval"
			Set Name = VerificationTask.GetAttributeNode("name")
			SMSPolicyRetrieval Name.Value
		Case "environmentvariablevalue"
			Set Name = VerificationTask.GetAttributeNode("name")
			Set EnvValue = VerificationTask.GetAttributeNode("envvalue")
			EnvironmentVariableValue Name.Value, EnvValue.Value
		Case "environmentvariablevaluecontains"
			Set Name = VerificationTask.GetAttributeNode("name")
			Set SubString = VerificationTask.GetAttributeNode("substring")
			EnvironmentVariableValueContains Name.Value, SubString.Value
		Case "readstdout"
			Set Name = VerificationTask.GetAttributeNode("name")
			Set SearchString = VerificationTask.GetAttributeNode("searchstring")
			Set MsgSuccess = VerificationTask.GetAttributeNode("msgsuccess")
			Set MsgFailure = VerificationTask.GetAttributeNode("msgfailure")
			ReadStdOut Name.Value,SearchString.Value,MsgSuccess.Value,MsgFailure.Value
		Case "timezonecorrect"
			TimeZoneCorrect
		Case "assistinifile"
			AssistIniFile
		Case "computerismemberofglobalgroup"
			Set Name = VerificationTask.GetAttributeNode("name")
			ComputerIsMemberOfGlobalGroup Name.Value
		Case "logicalfilesecurity"
			Set FileName = VerificationTask.GetAttributeNode("filename")
			Set Domain = VerificationTask.GetAttributeNode("domain")
			Set UserName = VerificationTask.GetAttributeNode("username")
			Set Mask = VerificationTask.GetAttributeNode("mask")
			LogicalFileSecurity FileName.Value, Domain.Value, UserName.Value, Mask.Value
		Case "logicalsharesecurity"
			Set ShareName = VerificationTask.GetAttributeNode("sharename")
			Set Domain = VerificationTask.GetAttributeNode("domain")
			Set UserName = VerificationTask.GetAttributeNode("username")
			Set Mask = VerificationTask.GetAttributeNode("mask")
			LogicalShareSecurity ShareName.Value, Domain.Value, UserName.Value, Mask.Value
		Case "ipsubnet"
			IPSubnet VerificationTask
		Case "connecttowmi"
			connecttowmi 
		Case "domainrole"
			DomainRole
		Case "lastbootuptime"
			LastBootupTime
		Case "freediskspace"
			FreeDiskSpace VerificationTask
		Case "filecontainsstringinlastlines"
			filecontainsstringinlastlines VerificationTask
		Case "filecontainsstringinlastlineserror"
			filecontainsstringinlastlineserror VerificationTask
		Case "drivesize"
			Set DriveLetter = VerificationTask.GetAttributeNode("driveletter")
			DriveSize DriveLetter.Value
		Case "reboot"
			Reboot VerificationTask
		Case "shutdown"
			Shutdown VerificationTask
		Case "copyfile"
			Set Name = VerificationTask.GetAttributeNode("name")
			Set Destination = VerificationTask.GetAttributeNode("destination")
			CopyFile Name.Value, Destination.Value
		Case "copyfolder"
			CopyFolder VerificationTask.GetAttributeNode("source").Value, VerificationTask.GetAttributeNode("destination").Value
		Case "createfolder"
			createfolder VerificationTask.GetAttributeNode("foldername").Value
		Case "deletefolder"
			deletefolder VerificationTask.GetAttributenode("foldername").Value
		Case "deletefile"
			deletefile VerificationTask.GetAttributenode("filename").Value
		Case "psexec"
			Set Name = VerificationTask.GetAttributeNode("name")
			PsExec Name.Value
		Case "sms_policyretrieved"
			sms_policyretrieved
		Case "verifydnstocomputername"
			verifydnstocomputername 
		Case "setdnsserversearchorder"
			SetDNSServerSearchOrder VerificationTask.GetAttributeNode("servicename").Value, VerificationTask.GetAttributeNode("primarydnsserver").Value
		Case "setdnssuffixsearchorder"
			setdnssuffixsearchorder VerificationTask.GetAttributeNode("dnsdomainsuffixsearchorder").Value
		Case "appenddnssuffixsearchorder"
			AppendDNSSuffixSearchOrder VerificationTask.GetAttributeNode("servicename").Value, VerificationTask.GetAttributeNode("dnsdomainsuffixsearchorder").Value
		Case "smsgenerateccr"
			SMSGenerateCCR VerificationTask.GetAttributeNode("outputdirectory").Value
		Case "appendtexttolocalfile"
			appendtexttolocalfile VerificationTask.GetAttributeNode("filename").Value, VerificationTask.GetAttributeNode("text").Value
		Case "issms2003clientinstalled"
			issms2003clientinstalled VerificationTask.GetAttributeNode("yes").Value, VerificationTask.GetAttributeNode("no").Value
		Case "issms2003clientassigned"
			issms2003clientassigned VerificationTask.GetAttributeNode("yes").Value, VerificationTask.GetAttributeNode("no").Value
		Case "smswmiquery" 
			smswmiquery VerificationTask
		Case "ingroup"
			ingroup VerificationTask.GetAttributeNode("name").Value
		Case "smsgetassignedsite"
			smsgetassignedsite
		Case "smssetassignedsite"
			smssetassignedsite VerificationTask.GetAttributeNode("sitecode").Value
		Case "smsrequestmachinepolicy"
			SMSRequestMachinePolicy
		Case "smsresetpolicy"
			SMSResetPolicy
		Case "sms_repairclient"
			sms_repairclient
		Case "sleep"
			Sleep VerificationTask
		Case "nbtstat_verifyhostname"
			nbtstat_verifyhostname VerificationTask
		Case "xml_querynode"
			xml_querynode VerificationTask
		Case "quit"
			Quit
		Case "smsguid"
			smsguid
		Case "radia_enumeratepackages"
			radia_enumeratepackages
		Case "sms_addtocollection"
			sms_addtocollection verificationTask
		Case "sms_removefromcollection"
			sms_removefromcollection verificationTask
		Case "sms_rerunadvert"
			sms_rerunadvert verificationTask
		Case "ad_distinguishedname"
			ad_distinguishedname
		Case "ad_addusertoglobalgroup"
			ad_addusertoglobalgroup verificationTask
		Case "ad_addworkstationtoglobalgroup"
			ad_addworkstationtoglobalgroup verificationTask
		Case "ad_removeworkstationfromglobalgroup"
			ad_removeworkstationfromglobalgroup verificationTask
		Case "ad_findduplicatecomputerinforest"
			ad_findduplicatecomputerinforest
		case "ad_computerpasswordage"
			ad_computerpasswordage
		case "getdnsserversearchorder" 
			getdnsserversearchorder verificationTask
		case "dhcpsettings"
			dhcpsettings verificationTask
		case "macaddress"
			macaddress verificationTask
		case "lookupgroupmembership"
			lookupgroupmembership verificationTask
		case "smsadvertisementstatusmessages"
			smsadvertisementstatusmessages verificationTask
		case "registrykeydelete"
			registrykeydelete verificationTask
		case "registrykeydeletevalue"
			registrykeydeletevalue verificationTask
		case "run"
			run verificationTask
		case "loggedonuser"
			LoggedOnUser verificationTask
		case "resourcedomain"
			ResourceDomain verificationTask
		Case "sms_mpcert"
			sms_mpcert verificationTask
		Case "sms_mplist"
			sms_mplist verificationTask
		Case "remotefolderexists"
			remotefolderexists verificationTask
		Case "remotefileexists"
			remotefileexists verificationTask
		Case "desktopiconexists"
			DesktopIconExists verificationTask
		Case "ad_addglobalgrouptolocalgroupwithauth"
			ad_addglobalgrouptolocalgroupwithauth verificationTask
		Case "adsite"
			adsite verificationTask
		Case "ad_location"
			ad_location verificationTask
		Case "sms_ccm_ctm_jobstateex_error"
			sms_ccm_ctm_jobstateex_error
		Case "peerdp_error"	
			peerdp_error
		Case "peerdp_status"
			peerdp_status VerificationTask
		Case "sms_ccm_executionrequest_error"
			sms_ccm_executionrequest_error
		Case "sms_ccm_softwaredistribution_exists"
			sms_ccm_softwaredistribution_exists VerificationTask
		Case "sms_mpproxyinformation"
			sms_mpproxyinformation 
		case "sms_policy"
			sms_policy VerificationTask
		Case "sms_lastadvertstatusmessage"
			sms_lastadvertstatusmessage VerificationTask
		Case "sms_cacheinfo"
			sms_cacheinfo VerificationTask
		Case "wmiobjectsdatahealth"
			wmiobjectsdatahealth
		Case "win32_ntlogevent"
			win32_ntlogevent VerificationTask
		Case "sms_dpfreespace"
			sms_dpfreespace VerificationTask
		Case "makeandmodel"
			makeandmodel VerificationTask
		Case "bluetooth"
			bluetooth VerificationTask
		Case "vpro"
			vpro VerificationTask
		Case "serialasset"
			serialasset VerificationTask
		Case "getcomputername"
			getcomputername
		Case "sms_cacheinfo1"
			sms_cacheinfo1 VerificationTask
		case sccm_policyreceived
			sccm_policyreceived VerificationTask
		Case "eventlogerrorcount"
			EventLogErrorCount VerificationTask
		Case "failedsystemapplications"
			failedsystemapplications
		Case "logicaldeviceerror"
			LogicalDeviceError
		Case "osd_schedulebuild"
			osd_schedulebuild VerificationTask
		Case "osd_cancelbuild"
			osd_cancelbuild VerificationTask
		Case "oldfiles"
			oldfiles VerificationTask
		Case "deleteoldfiles"
			deleteoldfiles VerificationTask
		Case "prestagebdp"
			prestagebdp VerificationTask.GetAttributeNode("source").Value
		Case Else
			MsgBox "Invalid Node: " & VerificationTask.nodeName
	End Select
End Sub

Sub WriteLog(verification, name, result, resultcomment)

	On Error Resume Next
	Dim NewNode
	if isNull(resultcomment) then
		resultcomment = ""
	end if 

	if isnull(name) = true then
		name = ""
	end if

	if isnull(result) = true then
		result = ""
	end if

	Set NewNode = LogFile.createNode("element","event","")
	NewNode.SetAttribute "result",Result
	NewNode.SetAttribute "verification",verification
	NewNode.SetAttribute "name", name
	NewNode.SetAttribute "resultcomment",resultcomment
	Root.AppendChild(NewNode)
	Set NewNode = Nothing
	If Result = "success" Then
		SuccessPoints = SuccessPoints + 1
	Elseif result = "failure" Then
		FailurePoints = FailurePoints + 1
	End If

End Sub

Sub UpdateHeaderInformation(OutputFileName)
	On Error Resume Next
	
	Exit Sub
	
	if isnull(Field1) Then
		Field1 = ""
	End If

	if isnull(Field2) then
		Field2 = ""
	End If

	if isnull(Field2) then
		Field2 = ""
	End If

	Root.setattribute "networkstatus",NetworkStatus
	Root.Setattribute "ipaddress",IPAddress
	Root.Setattribute "status",Status
	Root.Setattribute "field1",Field1
	Root.Setattribute "field2",Field2
	Root.Setattribute "field3",Field3
	Root.Setattribute "successpoints",SuccessPoints
	Root.Setattribute "failurepoints",FailurePoints
End Sub

Sub WriteControlFile(ControlFileName)
	
	Set oOutFile = oFSO.CreateTextFile(ControlFileName,True)
	oOutFile.Close()
	Set oOutFile = Nothing

End Sub

Sub InitializeLogFile(OutputFileName)
	On Error Resume Next

	' Create Base Nodes
	Set LogFile = CreateObject("Microsoft.XMLDOM")
	Set Root = LogFile.createNode("element","object","")
'	Root.SetAttribute "name",ComputerName
'	Root.SetAttribute "timestarted",TimeStarted
'	Root.SetAttribute "networkstatus",NetworkStatus
'	Root.SetAttribute "ipaddress",ipaddress
'	Root.SetAttribute "status",status
'	Root.SetAttribute "field1",Field1
'	Root.SetAttribute "field2",Fiedl2
'	Root.SetAttribute "field3",Fiedl3
'	Root.SetAttribute "successpoints",SuccessPoints
'	Root.SetAttribute "failurepoints",FailurePoints
	LogFile.appendChild(Root)

'	Set Summary = LogFile.createNode("element","summary","")
'	Root.AppendChild(Summary)
	
'	Set Events = LogFile.createNode("element","events","")
'	Root.AppendChild(Events)

	' Create summary nodes
	' ComputerName
' 	Set oComputerName = LogFile.createNode("element","computername","")
' 	oComputerName.Text = lcase(ComputerName)
' 	Summary.appendChild(oComputerName)
' 
' 	' TimeStarted
' 	Set oTimeStarted = LogFile.createNode("element","timestarted","")
' 	oTimeStarted.Text = TimeStarted
' 	Summary.appendchild(oTimeStarted)
' 
' 	' NetworkStatus
' 	Set oNetworkStatus = LogFile.createNode("element","networkstatus","")
' 	oNetworkStatus.Text = NetworkStatus
' 	Summary.appendchild(oNetworkStatus)
' 
' 	' IPAddress
' 	Set oIPAddress = LogFile.createNode("element","ipaddress","")
' 	oIPAddress.Text = IPAddress
' 	Summary.appendchild(oIPAddress)
' 
' 	' Status
' 	Set oStatus = LogFile.createNode("element","status","")
' 	oStatus.Text = Status
' 	Summary.appendchild(oStatus)
' 
' 	' Wave
' 	Set oWave = LogFile.createNode("element","wave","")
' 	oWave.Text = Field1
' 	Summary.appendchild(oWave)
' 
' 	' Type
' 	Set oType = LogFile.createNode("element","type","")
' 	oType.Text = Field2
' 	Summary.appendchild(oType)
' 	
' 	' Success Points
' 	Set oSuccessPoints = LogFile.createNode("element","successpoints","")
' 	oSuccessPoints.Text = SuccessPoints
' 	Summary.appendChild(oSuccessPoints)
' 
' 	' Failure points
' 	Set oFailurePoints = LogFile.createNode("element","failurepoints","")
' 	oFailurePoints.Text = FailurePoints
' 	Summary.appendChild(oFailurePoints)
' 
' 	' Failure Summary
' 	Set oFailureSummary = LogFile.createNode("element","failuresummary","")
' 	oFailureSummary.Text = FailureSummary
' 	Summary.appendChild(oFailureSummary)
' 
'	LogFile.save(OutputFileName)
End Sub

Function PreProcessValue(tmpvalue)
	On Error Resume Next
	If InStr(tmpValue,"%") < 1 Then
		PreProcessValue = tmpValue
		Exit Function
	End If
	PreProcessValue = Replace(tmpvalue,"%COMPUTERNAME%",ComputerName)
End Function

Sub connecttowmi()
	On Error Resume Next
	err.clear
	
	Set oSWbemServices = GetWMINamespace(ComputerName,"root\cimv2")
	If Err.Number <> 0 Then
		Status = Err.Description
		field1 = Status
		WriteLog "connecttowmi",Err.description,"failure",Err.Number
		Exit Sub
	End If

	err.clear
'	Set oRegistry = GetObject("Winmgmts:{impersonationLevel=Impersonate}!\\" & ComputerName & "\root\default:StdRegprov")
	
	Set oSWbemServicesDefault = GetWMINamespace(ComputerName,"root\Default")
	oSWbemServicesDefault.Security_.ImpersonationLevel = 3
	Set oRegistry = oSWbemServicesDefault.Get("StdRegProv")

'	Set oRegistry = GetWMINamespace("root\default:StdRegprov")
	If Err.Number <> 0 Then
		status = Err.Description 
		field1 = Status
		WriteLog "connecttowmi",Err.description,"failure",Err.Number
		Exit Sub
	End If

	If g_IntegratedAuthentication = True Then
		txt = "Integrated authentication"
	Else
		txt = g_LastSuccessfulUsername
	End If

	WriteLog "connecttowmi","Authentication","success",txt
End Sub

Sub PingComputer(node)
	On Error Resume Next
	Dim oPingResults, oPingResult
	Set oPingHost = oWbemLocator.ConnectServer (oNetwork.ComputerName, "root\cimv2")
	Set oPingResults = oPingHost.ExecQuery("SELECT ProtocolAddress, StatusCode FROM Win32_PingStatus WHERE Address = '" + ComputerName + "' and Timeout=" & node.getAttributeNode("timeout").Value)
	If Err.Number <> 0 Then
		IPAddress = Err.Number
		WriteLog "pingcomputer",Err.description,"failure",Err.Number
		NetworkStatus = "Error"
		Set oPingHost = Nothing
		Set oPingResults = Nothing
		' UpdateHeaderInformation(OutputFileName)
		Exit Sub
	End If

	For Each oPingResult In oPingResults
		Select Case oPingResult.StatusCode
				Case 0
					StatusCode = "On-line"
				Case 11001
					StatusCode = "Buffer Too Small"
				Case 11002
					StatusCode = "Destination Net Unreachable"
				Case 11003
					StatusCode = "Destination Host Unreachable"
				Case 11004
					StatusCode = "Destination Protocol Unreachable"
				Case 11005 
					StatusCode = "Destination Port Unreachable"
				Case 11006
					StatusCode = "No Resources"
				Case 11007
					StatusCode = "Bad Option"
				Case 11008
					StatusCode = "Hardware Error"
				Case 11009
					StatusCode = "Packet Too Big"
				Case 11010
					StatusCode = "Request Timed Out"
				Case 11011
					StatusCode = "Bad Request"
				Case 11012
					StatusCode = "Bad Route"
				Case 11013
					StatusCode = "TimeToLive Expired Transit"
				Case 11014
					StatusCode = "TimeToLive Expired Reassembly"
				Case 11015
					StatusCode = "Parameter Problem"
				Case 11016
					StatusCode = "Source Quench"
				Case 11017
					StatusCode = "Option Too Big"
				Case 11018
					StatusCode = "Bad Destination"
				Case 11032
					StatusCode = "Negotiating IPSEC"
				Case 11050
					StatusCode = "General Failure"
				Case Else
					StatusCode = "Unreachable"
		End Select
		NetworkStatus = StatusCode
		IPAddress = oPingResult.ProtocolAddress
	Next
	Set oPingHost = Nothing
	Set oPingResults = Nothing
	If NetworkStatus <> "On-line" Then
		WriteLog "pingcomputer",NetworkStatus,"failure",IPAddress
	Else
		WriteLog "pingcomputer",NetworkStatus,"success",IPAddress
	End If
	' UpdateHeaderInformation(OutputFileName)
End Sub

Sub SplitString(sString)
	On Error Resume Next
	
	Dim String1, FirstDash
	String1 = Mid(sString,15)
	FirstDash = InStr(1,String1,"-")
	Field2 = lCase(Left(String1,FirstDash-1))
	Field1 = lCase(Mid(String1,FirstDash+1))
End Sub

Function OpenConfigFile
	On Error Resume Next
	OpenConfigFile = False
	Set Config = CreateObject("Microsoft.XMLDOM")
	Config.async = False
	Config.validateonparse = False
	If Config.Load(Config_File) = False Then
		Exit Function
	End If

	OpenConfigFile = True
End Function

Sub Cleanup()
	On Error Resume Next
	Set oRegistry = Nothing
	Set oSWbemServices = Nothing
	Set oWbemLocator = Nothing
	Set Config = Nothing
	Set oFSO = Nothing
	Set oShell = Nothing
	Set oNetwork = Nothing
	WScript.Quit (0)
End Sub

Function WMIDateStringToDate(dtmInstallDate)
    On Error Resume Next
    WMIDateStringToDate = CDate(Mid(dtmInstallDate, 5, 2) & "/" & Mid(dtmInstallDate, 7, 2) & "/" & Left(dtmInstallDate, 4) & " " & Mid (dtmInstallDate, 9, 2) & ":" & Mid(dtmInstallDate, 11, 2) & ":" & Mid(dtmInstallDate, 13, 2))
End Function

Function AddBackslash(strPath)
	On Error Resume Next
	If Len(strPath) < 1 Then
		AddBackslash = strPath
		Exit Function
	End If
	
	If Right(strPath,1) = "\" Then
		AddBackslash = strPath
	Else
		AddBackslash = strPath & "\"
	End If
End Function

' ##################################################
' Verification tasks
' ##################################################

Sub ProcessKill(Name)
	On Error Resume Next
	Dim Item, Exists
	Exists = False
	

	Set oServices = oSWbemServices.ExecQuery("Select Name, creationdate from win32_process where name = '" & Name & "'")
	Exists = False
	For Each Item In oServices
		Item.Terminate()
		CreationDate = WMIDateStringToDate(Item.creationdate)
		Exists = True
	Next

	If Exists = True Then
		WriteLog "processkill",Name,"success","Running since " & CreationDate & " - Killed"
	Else
		WriteLog "processkill",Name,"failure","Process not running"
	End If
	Set oServices = Nothing
End Sub

Sub ProcessRunning(Name, OnlyIfUserIsLoggedOn)
	On Error Resume Next
	Dim Item, Exists
	Exists = False
	
	' Does it matter if the user is logged on
	If ucase(OnlyIfUserIsLoggedOn) = "TRUE" Then
		Set oComputerSystem = oSWbemServices.ExecQuery("Select UserName from Win32_ComputerSystem")
		For Each Item In oComputerSystem
			If Len(Item.UserName) > 0 Then
				UserLoggedOn = True
			Else
				userLoggedOn = False
			End If
		Next
		Set oComputerSystem = Nothing
	End If
	
	If UserLoggedOn = False And uCase(OnlyIfUserIsLoggedOn) = "TRUE" Then
		' Nobody is logged on so the processes are not expected to be running
		' Return a success and quit
		WriteLog "processrunning",Name,"success","No user is currently logged on"
		Exit Sub
	End If
	
	Set oServices = oSWbemServices.ExecQuery("Select Name, creationdate from win32_process where name = '" & Name & "'")
	Exists = False
	For Each Item In oServices
		Exists = True
		CreationDate = WMIDateStringToDate(Item.creationdate)
	Next

	If Exists = True Then
		WriteLog "processrunning",Name,"success",CreationDate
	Else
		WriteLog "processrunning",Name,"failure","Process not running"
	End If
	Set oServices = Nothing
End Sub

Sub ProcessRunningError(Name, OnlyIfUserIsLoggedOn)
	On Error Resume Next
	Dim Item, Exists
	Exists = False
	
	' Does it matter if the user is logged on
	If ucase(OnlyIfUserIsLoggedOn) = "TRUE" Then
		Set oComputerSystem = oSWbemServices.ExecQuery("Select UserName from Win32_ComputerSystem")
		For Each Item In oComputerSystem
			If Len(Item.UserName) > 0 Then
				UserLoggedOn = True
			Else
				userLoggedOn = False
			End If
		Next
		Set oComputerSystem = Nothing
	End If
	
	If UserLoggedOn = False And uCase(OnlyIfUserIsLoggedOn) = "TRUE" Then
		' Nobody is logged on so the processes are not expected to be running
		' Return a success and quit
		WriteLog "processrunningerror",Name,"success","No user is currently logged on"
		Exit Sub
	End If
	
	Set oServices = oSWbemServices.ExecQuery("Select Name, creationdate from win32_process where name = '" & Name & "'")
	Exists = False
	For Each Item In oServices
		Exists = True
		CreationDate = WMIDateStringToDate(Item.creationdate)
	Next

	If Exists = True Then
		WriteLog "processrunningerror",Name,"failure",CreationDate
	Else
		WriteLog "processrunningerror",Name,"success","Process not running"
	End If
	Set oServices = Nothing
End Sub

Sub ShareExists(Name)
	On Error Resume Next
	Dim Item, Exists
	Set oServices = oSWbemServices.ExecQuery("Select Name from win32_share where name = '" & Name & "'")
	Exists = False
	For Each Item In oServices
		Exists = True
	Next

	If Exists = True Then
		WriteLog "shareexists",Name,"success","Share exists"
	Else
		WriteLog "shareexists",Name,"failure","The share does not exist"
	End If
	Set oServices = Nothing
End Sub

Sub ServiceStartMode(Name,StartMode)
	On Error Resume Next
	Dim oService, blnFound
	blnFound = False
	Set oServices = oSWbemServices.ExecQuery("Select Name, StartMode from win32_service where name='" & Name & "'")
	For Each Item In oServices
		blnFound = True
		If Ucase(Item.StartMode) = UCase(StartMode) Then
			WriteLog "servicestartmode",Name,"success",StartMode
		Else
			WriteLog "servicestartmode",Name,"failure","Expecting: " & StartMode & ", Actual: " & Item.StartMode
		End If
	Next
	If blnFound = False Then
		WriteLog "servicestartmode",Name,"failure","Service does not exist"
	End If
	Set oServices = Nothing
End Sub

Sub ServiceStartModeError(Name,StartMode)
	On Error Resume Next
	Dim oService, blnFound
	blnFound = False
	Set oServices = oSWbemServices.ExecQuery("Select Name, StartMode from win32_service where name='" & Name & "'")
	For Each Item In oServices
		blnFound = True
		If Ucase(Item.StartMode) = UCase(StartMode) Then
			WriteLog "servicestartmodeerror",Name,"failure",Item.StartMode
		Else
			WriteLog "servicestartmodeerror",Name,"success",Item.StartMode
		End If
	Next
	If blnFound = False Then
		WriteLog "servicestartmodeerror",Name,"failure","Service does not exist"
	End If
	Set oServices = Nothing
End Sub


Sub ServiceState(Name,State)
	On Error Resume Next
	Dim oService, blnFound
	blnFound = False
	Set oServices = oSWbemServices.ExecQuery("Select Name, State from win32_service where name='" & Name & "'")
	For Each Item In oServices
		blnFound = True
		If Ucase(Item.State) = UCase(State) Then
			WriteLog "servicestate",Name,"success",State
		Else
			WriteLog "servicestate",Name,"failure","Expecting:" & State & ", Actual:" & Item.State
		End If
	Next
	If blnFound = False Then
		WriteLog "servicestate",Name,"failure","Service does not exist"
	End If
	Set oServices = Nothing
End Sub

Sub RegistryKeyMultiContains(Name,KeyValue)
	On Error Resume Next
	Dim RetVal, KeySplit, KeyBase, KeyName
	
	KeySplit = Split(Name,"\")
	KeyBase = KeySplit(0)
	KeyName = KeySplit(UBound(KeySplit))
	SubKeyName = Mid(Name,Len(KeyBase)+2,len(Name) - (Len(KeyBase) + Len(KeyName) + 2))
	
	' Constants
	Const HKEY_CLASSES_ROOT = &H80000000
	Const HKEY_CURRENT_USER = &H80000001 
	Const HKEY_LOCAL_MACHINE = &H80000002
	Const HKEY_USERS = &H80000003
	Const HKEY_CURRENT_CONFIG = &H80000005
	Const HKEY_DYN_DATA = &H80000006

	Select Case KeyBase
		Case "HKEY_CLASSES_ROOT"
			sKeyBase = HKEY_CLASSES_ROOT
		Case "HKEY_CURRENT_USER"
			sKeyBase = HKEY_CURRENT_USER
		Case "HKEY_LOCAL_MACHINE"
			sKeyBase = HKEY_LOCAL_MACHINE
		Case "HKEY_USERS"
			sKeyBase = HKEY_USERS
		Case "HKEY_CURRENT_CONFIG"
			sKeyBase = HKEY_CURRENT_CONFIG
		Case "HKEY_DYN_DATA"
			sKeyBase = HKEY_DYN_DATA
	End Select

	retVal = oRegistry.GetMultiStringValue (sKeyBase, SubKeyName,KeyName, RegVal)
	If RetVal <> 0 Then
		WriteLog "registrykeymulticontains",KeyBase & "\" & SubKeyName & "\" & KeyName,"failure","Registry key does not exist."
		Exit Sub
	End If

	For x = LBound(RegVal) To UBound(RegVal)
		If UCase(regval(x)) = uCase(KeyValue) Then
			Exists = True
			Exit For
		End If
	Next

	If Exists = True Then
		WriteLog "registrykeymulticontains",KeyBase & "\" & SubKeyName & "\" & KeyName,"success",KeyValue
	Else
		WriteLog "registrykeymulticontains",KeyBase & "\" & SubKeyName & "\" & KeyName,"failure",KeyValue & " does not exist."
	End If
	
End Sub

Sub RegistryKeyValue(Name,KeyType,KeyValue)
	On Error Resume Next
	Dim RetVal, KeySplit, KeyBase, KeyName
	Dim RegVal
	
	KeySplit = Split(Name,"\")
	KeyBase = KeySplit(0)
	KeyName = KeySplit(UBound(KeySplit))
	SubKeyName = Mid(Name,Len(KeyBase)+2,len(Name) - (Len(KeyBase) + Len(KeyName) + 2))
	
	Select Case KeyBase
		Case "HKEY_CLASSES_ROOT"
			sKeyBase = HKEY_CLASSES_ROOT
		Case "HKEY_CURRENT_USER"
			sKeyBase = HKEY_CURRENT_USER
		Case "HKEY_LOCAL_MACHINE"
			sKeyBase = HKEY_LOCAL_MACHINE
		Case "HKEY_USERS"
			sKeyBase = HKEY_USERS
		Case "HKEY_CURRENT_CONFIG"
			sKeyBase = HKEY_CURRENT_CONFIG
		Case "HKEY_DYN_DATA"
			sKeyBase = HKEY_DYN_DATA
	End Select

	Select Case KeyType
		Case "REG_SZ"
'			msgbox sKeyBase & vbcrlf & subkeyname & vbcrlf & keyname & vbcrlf & regval
			retVal = oRegistry.GetStringValue (sKeyBase, SubKeyName,KeyName, RegVal)
		Case "REG_DWORD"
			retVal = oRegistry.GetDWordValue(sKeyBase,SubKeyName,KeyName, RegVal)
		Case "REG_BINARY"
			retVal = oRegistry.GetBinaryValue(sKeyBase,SubKeyName,KeyName, RegVal)
	End Select

	if KeyType = "REG_BINARY" then
		str = ""
		for x = lbound(RegVal) to uBound(RegVal)
			str = str & " " & "0" & RegVal(x)
		Next
		RegVal = trim(str)
	End If

	If RetVal <> 0 Then
		WriteLog "registrykeyvalue",KeyBase & "\" & SubKeyName & "\" & KeyName,"failure","Expecting:" & KeyValue & ", Actual: Registry key does not exist."
		Exit Sub
	End If

	If Ucase(RegVal) = UCase(KeyValue) Then
		WriteLog "registrykeyvalue",KeyBase & "\" & SubKeyName & "\" & KeyName,"success",RegVal
		Field2 = RegVal
	Else
		WriteLog "registrykeyvalue",KeyBase & "\" & SubKeyName & "\" & KeyName,"failure","Expecting:" & KeyValue & ", Actual:" & RegVal
		Field2 = RegVal
	End If
	
End Sub

Sub RegistryKeyValueSet(Name,KeyType,KeyValue)
	On Error Resume Next
	Dim RetVal, KeySplit, KeyBase, KeyName
	
	KeySplit = Split(Name,"\")
	KeyBase = KeySplit(0)
	KeyName = KeySplit(UBound(KeySplit))
	SubKeyName = Mid(Name,Len(KeyBase)+2,len(Name) - (Len(KeyBase) + Len(KeyName) + 2))
	
	' Constants
	Const HKEY_CLASSES_ROOT = &H80000000
	Const HKEY_CURRENT_USER = &H80000001 
	Const HKEY_LOCAL_MACHINE = &H80000002
	Const HKEY_USERS = &H80000003
	Const HKEY_CURRENT_CONFIG = &H80000005
	Const HKEY_DYN_DATA = &H80000006

	Select Case KeyBase
		Case "HKEY_CLASSES_ROOT"
			sKeyBase = HKEY_CLASSES_ROOT
		Case "HKEY_CURRENT_USER"
			sKeyBase = HKEY_CURRENT_USER
		Case "HKEY_LOCAL_MACHINE"
			sKeyBase = HKEY_LOCAL_MACHINE
		Case "HKEY_USERS"
			sKeyBase = HKEY_USERS
		Case "HKEY_CURRENT_CONFIG"
			sKeyBase = HKEY_CURRENT_CONFIG
		Case "HKEY_DYN_DATA"
			sKeyBase = HKEY_DYN_DATA
	End Select

	Select Case KeyType
		Case "REG_SZ"
			retVal = oRegistry.SetStringValue (sKeyBase, SubKeyName,KeyName, KeyValue)
		Case "REG_DWORD"
			retVal = oRegistry.SetDWordValue(sKeyBase,SubKeyName,KeyName, KeyValue)
	End Select

	If RetVal <> 0 Then
		WriteLog "registrykeyvalueset",KeyBase & "\" & SubKeyName & "\" & KeyName,"failure","Return value=" & RetVal
		Exit Sub
	Else
		WriteLog "registrykeyvalueset",KeyBase & "\" & SubKeyName & "\" & KeyName,"success",KeyValue
	End If

End Sub

Sub registrykeydelete(node)
	On Error Resume Next
	Dim RetVal, KeySplit, KeyBase, KeyName
	
	Name = node.getAttributeNode("name").Value
	KeySplit = Split(Name,"\")
	KeyBase = KeySplit(0)
	KeyName = KeySplit(UBound(KeySplit))
	SubKeyName = Mid(Name,Len(KeyBase)+2,len(Name) - (Len(KeyBase) + Len(KeyName) + 2))
	key = SubKeyName & "\" & KeyName

	' Constants
	Const HKEY_CLASSES_ROOT = &H80000000
	Const HKEY_CURRENT_USER = &H80000001 
	Const HKEY_LOCAL_MACHINE = &H80000002
	Const HKEY_USERS = &H80000003
	Const HKEY_CURRENT_CONFIG = &H80000005
	Const HKEY_DYN_DATA = &H80000006

	Select Case KeyBase
		Case "HKEY_CLASSES_ROOT"
			sKeyBase = HKEY_CLASSES_ROOT
		Case "HKEY_CURRENT_USER"
			sKeyBase = HKEY_CURRENT_USER
		Case "HKEY_LOCAL_MACHINE"
			sKeyBase = HKEY_LOCAL_MACHINE
		Case "HKEY_USERS"
			sKeyBase = HKEY_USERS
		Case "HKEY_CURRENT_CONFIG"
			sKeyBase = HKEY_CURRENT_CONFIG
		Case "HKEY_DYN_DATA"
			sKeyBase = HKEY_DYN_DATA
	End Select

	If node.getAttributeNode("includesubkeys").Value = "true" Then
		oRegistry.EnumKey sKeyBase, Key, aKeys
		For x = LBound(aKeys) to uBound(aKeys)
			oRegistry.DeleteKey sKeyBase, Key & "\" & aKeys(x)
		Next
	End If

	Err.Clear
	ret = oRegistry.DeleteKey (sKeyBase, Key)
	If ret <> 0 Then
		WriteLog "registrykeydelete",KeyBase & "\" & Key,"failure","Error " & ret
	Else
		WriteLog "registrykeydelete",KeyBase & "\" & Key,"success","Key deleted"
	End If
	
End Sub

Sub registrykeydeletevalue(node)
	On Error Resume Next
	Dim RetVal, KeySplit, KeyBase, KeyName
	
	Name = node.getAttributeNode("name").Value
	KeySplit = Split(Name,"\")
	KeyBase = KeySplit(0)
	KeyName = KeySplit(UBound(KeySplit))
	SubKeyName = Mid(Name,Len(KeyBase)+2,len(Name) - (Len(KeyBase) + Len(KeyName) + 2))
	key = SubKeyName & "\" & KeyName

	' Constants
	Const HKEY_CLASSES_ROOT = &H80000000
	Const HKEY_CURRENT_USER = &H80000001 
	Const HKEY_LOCAL_MACHINE = &H80000002
	Const HKEY_USERS = &H80000003
	Const HKEY_CURRENT_CONFIG = &H80000005
	Const HKEY_DYN_DATA = &H80000006

	Select Case KeyBase
		Case "HKEY_CLASSES_ROOT"
			sKeyBase = HKEY_CLASSES_ROOT
		Case "HKEY_CURRENT_USER"
			sKeyBase = HKEY_CURRENT_USER
		Case "HKEY_LOCAL_MACHINE"
			sKeyBase = HKEY_LOCAL_MACHINE
		Case "HKEY_USERS"
			sKeyBase = HKEY_USERS
		Case "HKEY_CURRENT_CONFIG"
			sKeyBase = HKEY_CURRENT_CONFIG
		Case "HKEY_DYN_DATA"
			sKeyBase = HKEY_DYN_DATA
	End Select

	Err.Clear
	ret = oRegistry.DeleteValue (sKeyBase, SubKeyName, KeyName)
	If ret <> 0 Then
		WriteLog "registrykeydeletevalue",KeyBase & "\" & Key,"failure","Error " & ret
	Else
		WriteLog "registrykeydeletevalue",KeyBase & "\" & Key,"success","value deleted"
	End If
	
End Sub


Sub RegistryKeyValueError(Name,KeyType,KeyValue,ErrorIfKeyDoesNotExist)
	On Error Resume Next
	Dim RetVal, KeySplit, KeyBase, KeyName
	
	KeySplit = Split(Name,"\")
	KeyBase = KeySplit(0)
	KeyName = KeySplit(UBound(KeySplit))
	SubKeyName = Mid(Name,Len(KeyBase)+2,len(Name) - (Len(KeyBase) + Len(KeyName) + 2))
	
	' Constants
	Const HKEY_CLASSES_ROOT = &H80000000
	Const HKEY_CURRENT_USER = &H80000001 
	Const HKEY_LOCAL_MACHINE = &H80000002
	Const HKEY_USERS = &H80000003
	Const HKEY_CURRENT_CONFIG = &H80000005
	Const HKEY_DYN_DATA = &H80000006

	Select Case KeyBase
		Case "HKEY_CLASSES_ROOT"
			sKeyBase = HKEY_CLASSES_ROOT
		Case "HKEY_CURRENT_USER"
			sKeyBase = HKEY_CURRENT_USER
		Case "HKEY_LOCAL_MACHINE"
			sKeyBase = HKEY_LOCAL_MACHINE
		Case "HKEY_USERS"
			sKeyBase = HKEY_USERS
		Case "HKEY_CURRENT_CONFIG"
			sKeyBase = HKEY_CURRENT_CONFIG
		Case "HKEY_DYN_DATA"
			sKeyBase = HKEY_DYN_DATA
	End Select

	Select Case KeyType
		Case "REG_SZ"
			retVal = oRegistry.GetStringValue (sKeyBase, SubKeyName,KeyName, RegVal)
		Case "REG_DWORD"
			retVal = oRegistry.GetDWordValue(sKeyBase,SubKeyName,KeyName, RegVal)
	End Select
	
	If retVal <> 0 Then
		If uCase(ErrorIfKeyDoesNotExist) = "TRUE" Then
			WriteLog "registrykeyvalueerror",KeyBase & "\" & SubKeyName & "\" & KeyName,"failure","Registry key does not exist."
			Exit Sub
		Else
			WriteLog "registrykeyvalueerror",KeyBase & "\" & SubKeyName & "\" & KeyName,"success","Registry key does not exist."
			Exit Sub
		End If
	End If

	If Ucase(RegVal) = UCase(KeyValue) Then
		WriteLog "registrykeyvalueerror",KeyBase & "\" & SubKeyName & "\" & KeyName,"failure",RegVal
	Else
		WriteLog "registrykeyvalueerror",KeyBase & "\" & SubKeyName & "\" & KeyName,"success",RegVal
	End If
	
End Sub

Sub Folderexists(Name)
	On Error Resume Next
	
	Exists = False
	NewName = Replace(Name,"\","\\")
	Set colFolders = oSWbemServices.ExecQuery("Select * From Win32_Directory Where Name = '" & NewName & "'")

	For Each objFolder in colFolders
		Exists = True
	Next

    If Exists Then
    	WriteLog "folderexists",Name,"success",""
    Else
    	WriteLog "folderexists",Name,"failure","Folder does not exist"
    End If
    
    set colFolders = Nothing

End Sub


Sub FolderSize(Name)
	On Error Resume Next
	Dim oRemote
	
	Name = Replace(Name,"%COMPUTERNAME%",ComputerName)

	Err.Clear
	Set oRemote = oFSO.GetFolder(Name)
	If Err.Number <> 0 Then
	    	WriteLog "foldersize",Name,"failure","Folder does not exist"
		Exit Sub
	End If

    	WriteLog "foldersize",Name,"success",cstr(round(oRemote.Size/1000000,0))
		Field1 = cstr(round(oRemote.Size/1000000,0))
	Set oRemote = Nothing

End Sub

Sub oldfiles(node)
	On Error Resume Next
	Dim oRemote, age, fname
	
	Name = Replace(node.getattributenode("name").value,"%COMPUTERNAME%",ComputerName)

	Err.Clear
	Set oRemote = oFSO.GetFolder(Name)
	If Err.Number <> 0 Then
	    WriteLog "oldfiles",Name,"failure","Folder does not exist"
		Exit Sub
	End If

	For each file in oRemote.files
		fname = file
		Set oFile = oFSO.getfile(fname)
		age = DateDiff("d", oFile.DateLastModified, Now)
		if age > cint(node.getattributenode("maxage").value) then
			WriteLog "oldfiles",fname,"failure","DateLastModified (" & oFile.DateLastModified & ") = " & age & " days old"
		else
			WriteLog "oldfiles",fname,"success","DateLastModified (" & oFile.DateLastModified & ") = " & age & " days old"
		end if 
		Set oFile = Nothing
	next

	Set oRemote = Nothing

End Sub

Sub deleteoldfiles(node)
	On Error Resume Next
	Dim oRemote, age, fname
	
	Name = Replace(node.getattributenode("name").value,"%COMPUTERNAME%",ComputerName)

	Err.Clear
	Set oRemote = oFSO.GetFolder(Name)
	If Err.Number <> 0 Then
	    WriteLog "deleteoldfiles",Name,"failure","Folder does not exist"
		Exit Sub
	End If

	For each file in oRemote.files
		fname = file
		Set oFile = oFSO.getfile(fname)
		age = DateDiff("d", oFile.DateLastModified, Now)
		if age > cint(node.getattributenode("maxage").value) then
			err.clear
			ofile.delete(true)
			if err.number <> 0 then
				WriteLog "deleteoldfiles",fname,"failure","Error deleting file: err.number=" & err.number & ", err.description=" & err.description
			else
				WriteLog "deleteoldfiles",fname,"success","File Deleted - DateLastModified Age = " & age & " days old"
			end if 
		else
			WriteLog "deleteoldfiles",fname,"success","File NOT deleted - DateLastModified (" & oFile.DateLastModified & ") = " & age & " days old"
		end if 
		Set oFile = Nothing
	next

	Set oRemote = Nothing

End Sub


Sub FileExists(FileName)
	On Error Resume Next

	Exists = False
	NewName = Replace(FileName,"\","\\")
	Set colFiles = oSWbemServices.ExecQuery ("Select * from CIM_Datafile Where name = '" & NewName & "'")
	For Each objFile in colFiles
		Exists = True
	Next
	
	If Exists Then
		WriteLog "fileexists",FileName,"success",""
	Else
		WriteLog "fileexists",FileName,"failure","File does not exist"
	End If
	
	Set colFiles = Nothing
End Sub

Sub FileVersion(Name, Version)
	On Error Resume Next
	Dim FileVersion, Exists
	NewName = Replace(Name,"\","\\")
	Exists = False
	Set colFiles = oSWbemServices.ExecQuery ("Select * from CIM_Datafile Where name = '" & NewName & "'")
	For Each objFile in colFiles
		Exists = True
		FileVersion = objFile.Version
	Next
	
	If Exists = False Then
		WriteLog "fileversion",Name,"failure","File doesn't exist"
		Exit Sub
	End If

Field1 = FileVersion

	If uCase(FileVersion) = uCase(Version) Then
		WriteLog "fileversion",Name,"success",FileVersion
	Else
		WriteLog "fileversion",Name,"failure","Expecting:" & Version & ", Actual:" & FileVersion
	End If

	Set colFiles = Nothing
End Sub

Sub FileSize(Name,Size)
	On Error Resume Next
	Name = Replace(Name,"%COMPUTERNAME%",ComputerName)

	Err.Clear
	Set oRemote = oFSO.GetFile(Name)
	If Err.Number <> 0 Then
		WriteLog "filesize",Name,"failure","File does not exist"
		Exit Sub
	End If

    WriteLog "filesize",Name,"success",cstr(round(oRemote.Size/1000000,0))
    Field3 = cstr(round(oRemote.Size/1000000,0))
	Set oRemote = Nothing

End Sub


Sub FileDateCreated(Name, DateCreated)
	On Error Resume Next
	Dim Exists, FileCreationDate
	Exists = False
	NewName = Replace(Name,"\","\\")
	Set colFiles = oSWbemServices.ExecQuery ("Select * from CIM_Datafile Where name = '" & NewName & "'")
	For Each objFile in colFiles
		Exists = True
		FileCreationDate = WMIDateStringToDate(objFile.CreationDate)
	Next
	
	If Exists = False Then
		WriteLog "filedatecreated",Name,"failure","File doesn't exist"
		Exit sub
	End If
	
	If uCase(DateCreated) = uCase(FileCreationDate) Then
		WriteLog "filedatecreated",Name,"success",FileCreationDate
	Else
		WriteLog "filedatecreated",Name,"failure","Expecting:" & DateCreated & ", Actual:" & FileCreationDate
	End If

	Set colFiles = Nothing
End Sub

Sub FileDateModified(Name, DateModified)
	On Error Resume Next
	Dim Exists, FileDateModified
	
	Exists = False
	NewName = Replace(Name,"\","\\")
	Set colFiles = oSWbemServices.ExecQuery ("Select * from CIM_Datafile Where name = '" & NewName & "'")
	For Each objFile in colFiles
		Exists = True
		FileModifiedDate = WMIDateStringToDate(objFile.LastModified)
	Next

	If Exists = False Then
		Exit Sub
		WriteLog "filedatemodified",Name,"failure","File doesn't exist"
	End If

	If uCase(DateModified) = uCase(FileModifiedDate) Then
		WriteLog "filedatemodified",Name,"success",FileModifiedDate
	Else
		WriteLog "filedatemodified",Name,"failure","Expecting:" & DateModified & ", Actual:" & FileModifiedDate
	End If

	Set colFiles = Nothing
End Sub

Sub SMSpolicyRetrieval(Name)
	On Error Resume Next

	Err.Clear
	Set oServices = GetWMINamespace(ComputerName,"root\ccm")
	If Err.Number <> 0 Then
		WriteLog "smspolicyretrieval",Name,"failure",Err.number & "," & Err.Description
		Exit Sub
	End If
	set oInstance = oServices.Get("SMS_Client")
	set oParams = oInstance.Methods_("TriggerSchedule").inParameters.SpawnInstance_()
	oParams.sScheduleID = Name
	oServices.ExecMethod "SMS_Client", "TriggerSchedule", oParams
	if (Err.number <> 0) Then
		WriteLog "smspolicyretrieval",Name,"failure",Err.number & "," & Err.Description
	Else
		WriteLog "smspolicyretrieval",Name,"success",""
	end If

	
	Set oInstance = Nothing
	Set oParams = Nothing	
	Set oServices = Nothing
End Sub

Sub EnvironmentVariableExists(Name)
	On Error Resume Next
	Dim EnvVar

	Exists = False
	Set Win32_Environment = oSWbemServices.ExecQuery("select Name, VariableValue from Win32_Environment where Name ='" & Name & "'")
	For Each Item In Win32_Environment
		Exists = True
		EnvVar = Item.VariableValue
		Exit For
	Next
	
	If Exists = True Then
		WriteLog "environmentvariableexists",Name,"success",EnvVar
	Else
		WriteLog "environmentvariableexists",Name,"failure","Does not exist."
	End If
	
	Set Win32_Environment = Nothing
End Sub

Sub EnvironmentVariableValue(Name,EnvValue)
	On Error Resume Next
	Dim EnvVar

	EnvValue = PreProcessValue(EnvValue)

	Exists = False
	Set Win32_Environment = oSWbemServices.ExecQuery("select Name, VariableValue from Win32_Environment where Name ='" & Name & "'")
	For Each Item In Win32_Environment
		Exists = True
		EnvVar = Item.VariableValue
		Exit For
	Next

	If Exists = False Then
		WriteLog "environmentvariablevalue",Name,"failure","Does not exist."
		Exit Sub
	End If
	
	If uCase(EnvValue) = uCase(EnvVar) Then
		WriteLog "environmentvariablevalue",Name,"success",EnvVar
	Else
		WriteLog "environmentvariablevalue",Name,"failure","Expecting:" & EnvValue & ", Actual:" & EnvVar
	End If
	
	Set Win32_Environment = Nothing
End Sub

Sub EnvironmentVariableValueContains(Name,SubString)
	On Error Resume Next
	Dim EnvVar

	Exists = False
	Set Win32_Environment = oSWbemServices.ExecQuery("select Name, VariableValue from Win32_Environment where Name ='" & Name & "'")
	For Each Item In Win32_Environment
		Exists = True
		EnvVar = Item.VariableValue
		Exit For
	Next

	If Exists = False Then
		WriteLog "environmentvariablevaluecontains",Name,"failure","Does not exist."
		Exit Sub
	End If
	
	If instr(lCase(EnvVar), lCase(SubString)) > 0 Then
		WriteLog "environmentvariablevaluecontains",Name,"success",SubString
	Else
		WriteLog "environmentvariablevaluecontains",Name,"failure",SubString & " does not exist in variable."
	End If
	
	Set Win32_Environment = Nothing
End Sub

Sub ReadStdOut(Name,SearchString,MsgSuccess,MsgFailure)
	On Error Resume Next

	Name = PreProcessValue(Name)
	Set oExec = oShell.Exec (Name)
	If Err.Number <> 0 Then
		WriteLog "readstdout",Name,"failure","Error executing process: " & Err.Number & ", " & Err.Description
		Exit sub
	End If
	Do While oExec.Status = 0
		WScript.Sleep 100
	Loop
	StdOutAll = uCase(oExec.StdOut.ReadAll)
	If inStr(StdOutAll,uCase(SearchString)) > 0 Then
		WriteLog "readstdout",Name,"success",MsgSuccess
	Else
		WriteLog "readstdout",Name,"failure",MsgFailure
	End IF
	Set oExec = Nothing
End Sub

Sub DatabaseConnection(Name)
	On Error Resume Next
	Set oConnection = CreateObject("ADODB.Connection")
	Err.Clear
	oConnection.ConnectionString = "Provider='SQLOLEDB';Data Source='" & ComputerName & "';Initial Catalog='" & Name & "' ;Integrated Security='SSPI';"
	oConnection.ConnectionTimeout = 10
	oConnection.CommandTimeout = 5
	oConnection.Open
	If Err.Number <> 0 Then
		WriteLog "databaseconnection",Name,"failure","Error " & Err.Description & " - " & Err.Number 
	Else
		WriteLog "databaseconnection",Name,"success",StartMode
	End If
	oConnection.Close
	Set oConnection = Nothing
End Sub

Sub ScheduledTaskExists(Name)
	On Error Resume Next
	Dim Exists, Item
	Set oServices = oSWbemServices.ExecQuery("Select * from win32_ScheduledJob")
	Exists = False
	For Each Item In oServices
		If uCase(Item.Command) = uCase(Name) Then
			Exists = True
		End If
	Next
	If Exists = True Then
		WriteLog "scheduledtaskexists",Name,"success","Task exists"
	Else
		WriteLog "scheduledtaskexists",Name,"failure","The task does not exist"
	End If
	Set oServices = Nothing
End Sub

Sub TimeZoneCorrect()
	On Error Resume Next
	Dim Item, TimeZoneOffset, PostCodeCharacter, CorrectTimeZone, TimeZone
	Set oComputerSystem = oSWbemServices.ExecQuery("Select * from win32_ComputerSystem")
	For Each Item In oComputerSystem
		TimeZoneOffset = Item.CurrentTimeZone
	Next
	CorrectTimeZone = false
	PostCodeCharacter = cstr(Mid(ComputerName,2,1))
	TimeZone = "Unknown"
	Select Case PostCodeCharacter
		Case "2" ' NSW
			TimeZone = "NSW"
			If TimeZoneOffset = "600" Then
				CorrectTimeZone = True
			End If
		Case "3" ' VIC
			TimeZone = "VIC"
			If TimeZoneOffset = "600" Then
				CorrectTimeZone = True
			End If
		Case "4"
			TimeZone = "QLD"
			If TimeZoneOffset = "600" Then
				CorrectTimeZone = True
			End If
		Case "5"
			TimeZone = "SA"
			If TimeZoneOffset = "570" Then
				CorrectTimeZone = True
			End If
		Case "6"
			TimeZone = "WA"
			If TimeZoneOffset = "480" Then
				CorrectTimeZone = True
			End If
		Case "7"
			TimeZone = "TAS"
			If TimeZoneOffset = "600" Then
				CorrectTimeZone = True
			End If
		Case Else
			TimeZone = "Unknown"
			CorrectTimeZone = False
	End Select

	Set oComputerSystem = Nothing
	If CorrectTimeZone = True Then
		WriteLog "timezonecorrect",TimeZone,"success","TimeZoneOffset: " & TimeZoneOffset
		Exit Sub
	End If

	WriteLog "timezonecorrect",TimeZone,"failure","TimeZoneOffset: " & TimeZoneOffset
End Sub

Function distinguishedName(ComputerName)
	On Error Resume Next
	Const ADS_SCOPE_SUBTREE = 2
	
	distinguishedName = ""
	Set objRootDSE = GetObject("LDAP://rootDSE")
	strRootDomain = "LDAP://" & objRootDSE.Get("rootDomainNamingContext")
	Set objRootDomain = GetObject(strRootDomain)
	RootName = objRootDomain.Name
	Set objRootDSE = Nothing
	Set objRootDomain = Nothing

	Set objConnection = CreateObject("ADODB.Connection")
	Set objCommand =   CreateObject("ADODB.Command")
	objConnection.Provider = "ADsDSOObject"
	objConnection.Open "Active Directory Provider"
	Set objCommand.ActiveConnection = objConnection
	
	objCommand.Properties("Page Size") = 1000
	objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 

	objCommand.CommandText = "Select cn, distinguishedName from 'GC://" & RootName & ",DC=com' where objectClass='computer' and name='" & ComputerName & "'"

	Set objRecordSet = objCommand.Execute

	If objRecordSet.EOF Then
		Set objConnection = Nothing
		Set objCommand =  Nothing
		Set objRecordSet = Nothing
		Exit Function
	End If 

	dtmMostRecent = "January 1, 1980"

	objRecordSet.MoveFirst
	If objRecordset.RecordCount = 1 Then
		Do Until objRecordSet.EOF
			distinguishedName = objRecordSet.Fields("distinguishedName").Value
			objRecordSet.MoveNext
		Loop
	Else ' Recordset > 1 so we have duplicates
		Do Until objRecordSet.EOF
			dn = objRecordSet.Fields("distinguishedName").Value
			dtmPassword = GetPasswordAge(dn)
			if datediff("s",dtmPassword,dtmMostRecent) > 0 Then
			else
				dtmMostRecent = dtmPassword
				mostRecentdn = dn
			End If
			objRecordSet.MoveNext
		Loop
		distinguishedName = mostRecentdn
	end If

	Set objConnection = Nothing
	Set objCommand =  Nothing
	Set objRecordSet = Nothing
End Function

Function distinguishedNameUser(UserName)
	On Error Resume Next
	Const ADS_SCOPE_SUBTREE = 2
	
	distinguishedName = ""
	Set objRootDSE = GetObject("LDAP://rootDSE")
	strRootDomain = "LDAP://" & objRootDSE.Get("rootDomainNamingContext")
	Set objRootDomain = GetObject(strRootDomain)
	RootName = objRootDomain.Name
	Set objRootDSE = Nothing
	Set objRootDomain = Nothing

	Set objConnection = CreateObject("ADODB.Connection")
	Set objCommand =   CreateObject("ADODB.Command")
	objConnection.Provider = "ADsDSOObject"
	objConnection.Open "Active Directory Provider"
	Set objCommand.ActiveConnection = objConnection
	
	objCommand.Properties("Page Size") = 1000
	objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 

	objCommand.CommandText = "Select cn, distinguishedName from 'GC://" & RootName & ",DC=com' where objectClass='user' and name='" & UserName & "'"

	Set objRecordSet = objCommand.Execute

	If objRecordSet.EOF Then
		Set objConnection = Nothing
		Set objCommand =  Nothing
		Set objRecordSet = Nothing
		Exit Function
	End if 

	dtmMostRecent = "January 1, 1980"

	objRecordSet.MoveFirst
	If objRecordset.RecordCount = 1 Then
		Do Until objRecordSet.EOF
			distinguishedNameUser = objRecordSet.Fields("distinguishedName").Value
			objRecordSet.MoveNext
		Loop
	Else ' Recordset > 1 so we have duplicates
		Do Until objRecordSet.EOF
			dn = objRecordSet.Fields("distinguishedName").Value
			dtmPassword = GetPasswordAge(dn)
			if datediff("s",dtmPassword,dtmMostRecent) > 0 Then
			else
				dtmMostRecent = dtmPassword
				mostRecentdn = dn
			End If
			objRecordSet.MoveNext
		Loop
		distinguishedNameUser = mostRecentdn
	end If

	Set objConnection = Nothing
	Set objCommand =  Nothing
	Set objRecordSet = Nothing
End Function


Function ComputerIsMemberOfGlobalGroup(GroupName)
	On Error Resume Next
	Dim aMemberOf
	ReDim aMemberOf(-1)

	dn = distinguishedName(ComputerName)
	Set objComputer = GetObject ("LDAP://" & dn)
	objComputer.GetInfo
	aMemberOf = objComputer.GetEx("memberOf")
	If UBound(aMemberOf) < 0 Then
		Set objComputer = Nothing
		WriteLog "computerismemberofglobalgroup",GroupDistinguishedName,"failure","Computer is not a member of any groups"
		Exit Function
	End If

	For Each Group In aMemberOf
		FirstComma = Instr(Group,",")
		TestGroupName = Mid(Group,4,FirstComma-4)
		If Ucase(GroupName) = uCase(TestGroupName) Then
			WriteLog "computerismemberofglobalgroup",TestGroupName,"success",Group
			Set objComputer = Nothing
			Exit Function
		End If
	Next

	WriteLog "computerismemberofglobalgroup",GroupName,"failure","Computer is not a member"
	Set objComputer = Nothing
End Function

Function LogicalFileSecurity(FileName, Domain, UserName, Mask)

	On Error Resume Next
	NewFileName = Replace(FileName,"\","\\")
	err.clear
'	Set wmiFileSecSetting = GetObject("Winmgmts:{impersonationLevel=Impersonate}!\\" & ComputerName & "\root\cimv2:Win32_LogicalFileSecuritySetting.path='" & NewFileName & "'")
	Set wmiFileSecSetting = GetWMINamespace(ComputerName,"root\cimv2:Win32_LogicalFileSecuritySetting.path='" & NewFileName & "'")
	If Err <> 0 Then 
	    WriteLog "logicalfilesecurity",FileName,"failure","Description: " & Err.Description & ", Number: " & Err.Number
		Exit Function
	End If
	RetVal = wmiFileSecSetting.GetSecurityDescriptor(wmiSecurityDescriptor) 
	If Err <> 0 Then 
	    WriteLog "logicalfilesecurity",FileName,"failure","Description: " & Err.Description & ", Number: " & Err.Number
		Exit Function
	End If
	
	' array of Win32_ACE objects.
	DACL = wmiSecurityDescriptor.DACL 
	For each wmiAce in DACL 
		Set Trustee = wmiAce.Trustee 
		If Domain = "" And (cstr(uCase(Trustee.Name)) = UCase(UserName)) Then
			If Cstr(wmiAce.AccessMask) = Mask Then
			    WriteLog "logicalfilesecurity",FileName,"success",UserName & ", Mask: " & wmiAce.AccessMask
			Else
			    WriteLog "logicalfilesecurity",FileName,"failure",UserName & ", Expecting mask: " & Mask & ", Actual mask: " & wmiAce.AccessMask
			End If
			Set Trustee = Nothing
			Set wmiFileSecSetting = Nothing
			Exit Function
		Elseif (cstr(uCase(Trustee.Domain)) = UCase(Domain)) And (cstr(uCase(Trustee.Name)) = UCase(UserName)) Then
			If Cstr(wmiAce.AccessMask) = Mask Then
			    WriteLog "logicalfilesecurity",FileName,"success",Domain & "\" & UserName & ", Mask: " & wmiAce.AccessMask
			Else
			    WriteLog "logicalfilesecurity",FileName,"failure",Domain & "\" & UserName & ", Expecting mask: " & Mask & ", Actual mask: " & wmiAce.AccessMask
			End If
			Set Trustee = Nothing
			Set wmiFileSecSetting = Nothing
			Exit Function
		End If
		Set Trustee = Nothing
'		wscript.echo "ACE Type: "        & wmiAce.AceType 
	Next 
    WriteLog "logicalfilesecurity",FileName,"failure",Domain & "\" & UserName & " has no ACE in DACL"
	Set wmiFileSecSetting = Nothing

End Function

Function LogicalShareSecurity(ShareName, Domain, UserName, Mask)

	On Error Resume Next
	err.clear
'	Set wmiShareSecSetting = GetObject("Winmgmts:{impersonationLevel=Impersonate}!\\" & ComputerName & "\root\cimv2:Win32_LogicalShareSecuritySetting.Name='" & ShareName & "'")
	Set wmiSharesecSetting = GetWMINamespace(ComputerName,"root\cimv2:Win32_LogicalShareSecuritySetting.Name='" & ShareName & "'")
	If Err <> 0 Then 
	    WriteLog "logicalsharesecurity",ShareName,"failure","Description: " & Err.Description & ", Number: " & Err.Number
		Exit Function
	End If
	RetVal = wmiShareSecSetting.GetSecurityDescriptor(wmiSecurityDescriptor) 
	If Err <> 0 Then 
	    WriteLog "logicalsharesecurity",ShareName,"failure","Description: " & Err.Description & ", Number: " & Err.Number
		Exit Function
	End If
	
	' array of Win32_ACE objects.
	DACL = wmiSecurityDescriptor.DACL

	For each wmiAce in DACL 
		Set Trustee = wmiAce.Trustee 

		If (cstr(uCase(Trustee.Domain)) = UCase(Domain)) And (cstr(uCase(Trustee.Name)) = UCase(UserName)) Then
			If Cstr(wmiAce.AccessMask) = Mask Then
			    WriteLog "logicalsharesecurity",ShareName,"success",Domain & "\" & UserName & ", Mask: " & wmiAce.AccessMask
			Else
			    WriteLog "logicalsharesecurity",ShareName,"failure",Domain & "\" & UserName & ", Expecting mask: " & Mask & ", Actual mask: " & wmiAce.AccessMask
			End If
			Set Trustee = Nothing
			Set wmiShareSecSetting = Nothing
			Exit Function
		End If
		Set Trustee = Nothing
'		wscript.echo "ACE Type: "        & wmiAce.AceType 
	Next 
	WriteLog "logicalsharesecurity",ShareName,"failure",Domain & "\" & UserName & " has no ACE in DACL"
	Set wmiShareSecSetting = Nothing

End Function

Function IPSubnet(node)
	Set colAdapters = oSWbemServices.ExecQuery("Select * from win32_networkadapterconfiguration where IPEnabled = True")
	
	For Each Adapter In colAdapters
		For x = LBound(Adapter.IPAddress) To UBound(Adapter.IPAddress)
			If IPAddress = Adapter.IPAddress(x) Then
				SubnetMask = SubnetIt(Adapter.IPAddress(x),Adapter.IPSubnet(x))
				Field1 = SubnetMask
				WriteLog "ipsubnet",IPAddress,"success",SubnetMask
				Set colAdapters = Nothing
				Exit Function
			End If
		Next
	Next
	WriteLog "ipsubnet",IPAddress,"failure","Failed to retrieve subnet mask"
	Set colAdapters = Nothing
End Function



Function SubNetIt(Address1, Subnet1) 
 
 	Address = Address1
 	Subnet = Subnet1
 	
     dim addressbytes(4) 
     dim subnetmaskbytes(4) 
 
     i=0 
     period = 1 
     while period<>len( address ) + 2 
           prevperiod=period 
           period = instr( period+1, address, "." ) + 1 
           if period = 1 then period = len( address ) + 2 
           addressbyte = mid( address, prevperiod, period-prevperiod-1 ) 
           addressbytes(i)=addressbyte 
           i=i+1 
  	wend 
 
  i=0 
  period = 1 
  while period<>len( subnet ) + 2 
           prevperiod=period 
           period = instr( period+1, subnet, "." ) + 1 
           if period = 1 then period = len( subnet ) + 2 
           subnetmaskbyte = mid( subnet, prevperiod, period-prevperiod-1 )
           subnetmaskbytes(i)=subnetmaskbyte 
           i=i+1 
 
  wend 
 
  subnet="" 
  for i=0 to 3 
 
           subnet = subnet & (addressbytes(i) AND subnetmaskbytes(i)) & "." 
  next 
  subnet = left( subnet, len(subnet)-1 ) 
  SubnetIt = subnet
 
End Function

Function FreeDiskSpace(Node)
	On Error Resume Next
	Dim Win32_LogicalDisk, objDisk, colDisks

	Set Win32_LogicalDisk = oSWbemServices.ExecQuery("select Name, FreeSpace from Win32_LogicalDisk where Name ='" & Node.GetAttributeNode("driveletter").Value & "'")
	For each objDisk in Win32_LogicalDisk
		field1 = clng(round(objDisk.FreeSpace/1000000,0))
		If clng(round(objDisk.FreeSpace/1000000,0)) <= clng(Node.GetAttributeNode("minimumsize").Value) Then
			WriteLog "freediskspace",objDisk.Name,"failure",cstr(round(objDisk.FreeSpace/1000000,0)) & " < " & Node.GetAttributeNode("minimumsize").Value
		Else
			WriteLog "freediskspace",objDisk.Name,"success",cstr(round(objDisk.FreeSpace/1000000,0))
		End If
		Set Win32_LogicalDisk = Nothing
		Exit Function
	Next
	WriteLog "freediskspace",Node.GetAttributeNode("driveletter").Value,"failure","Drive not found"
    Set Win32_LogicalDisk = Nothing
End Function

Function DriveSize(DriveLetter)
	On Error Resume Next
	Dim Win32_LogicalDisk, objDisk, colDisks
	Set Win32_LogicalDisk = oSWbemServices.ExecQuery("select Name, Size from Win32_LogicalDisk where Name ='" & DriveLetter & "'")
	For each objDisk in Win32_LogicalDisk
		WriteLog "drivesize",objDisk.Name,"success",cstr(round(objDisk.Size/1000000,0))
		Field2 = cstr(round(objDisk.Size/1000000,0))
		Set Win32_LogicalDisk = Nothing
		Exit Function
	Next
	WriteLog "drivesize",DriveLetter,"failure","Drive not found"
	Set Win32_LogicalDisk = Nothing
End Function


Function DomainRole()
	On Error Resume Next
	Dim Win32_ComputerSystem, objItem, colItems, Role
	Set Win32_ComputerSystem = oSWbemServices.ExecQuery("select DomainRole from Win32_ComputerSystem")
	For each objItem in Win32_ComputerSystem
		Select Case CStr(objItem.DomainRole)
			Case "0"
				Role = "Standalone Workstation"
			Case "1" 
				Role = "Member Workstation"
			Case "2"
				Role = "Standalone Server"
			Case "3"
				Role = "Member Server"
			Case "4"
				Role = "Backup Domain Controller"
			Case "5"
				Role = "Primary Domain Controller"
		End Select
		WriteLog "domainrole",CStr(objItem.DomainRole),"success",Role
		Field2 = Role
		Set Win32_ComputerSystem = Nothing
		Exit Function
	Next
	WriteLog "domainrole",,"failure","Unable to retrieve"
	Set Win32_ComputerSystem = Nothing
End Function

Function GetComputerName()
	On Error Resume Next
	Dim Win32_ComputerSystem, objItem, colItems, Role
	Set Win32_ComputerSystem = oSWbemServices.ExecQuery("select Caption from Win32_ComputerSystem")
	For each objItem in Win32_ComputerSystem
		WriteLog "getcomputername",CStr(objItem.Caption),"success",Caption
		Field2 = CStr(objItem.Caption)
		Set Win32_ComputerSystem = Nothing
		Exit Function
	Next
	WriteLog "getcomputername",,"failure","Unable to retrieve"
	Set Win32_ComputerSystem = Nothing
End Function


Sub CreateProcess(Name)
	On Error Resume Next
	Dim Item, Exists
	Const HIDDEN_WINDOW = 12
	
	Set oServices = oSWbemServices.Get("Win32_Process")
	Set oStartup = oSWbemServices.Get("Win32_ProcessStartup")
	set oConfig = oStartup.spawnInstance_

	oConfig.ShowWindow = HIDDEN_WINDOW

	errNo = oServices.Create(Name,Null,oConfig,intProcessID)

	If ErrNo = 0 Then
		WriteLog "createprocess",Name,"success","Created process ID " & intProcessID
	Else
		WriteLog "createprocess",Name,"failure","Failed to create process - " & ErrNo
	End If
	Set oConfig = Nothing
	Set oStartup = Nothing
	Set oServices = Nothing
End Sub

Function LastBootupTime()
	On Error Resume Next
	Dim Win32_OperatingSystem, objItem, colItems, Role
	Set Win32_OperatingSystem = oSWbemServices.ExecQuery("select LastBootupTime from Win32_OperatingSystem")
	For each objItem in Win32_OperatingSystem
		WriteLog "lastbootuptime",FormatDate(objItem.LastBootupTime),"success",""
		Field2 = FormatDate(objItem.LastBootupTime)
		Set Win32_OperatingSystem = Nothing
		Exit Function
	Next
	WriteLog "lastbootuptime",,"failure","Unable to retrieve"
	Set Win32_OperatingSystem = Nothing
End Function

Function Reboot(node)
	On Error Resume Next
	Dim rc
	Dim Win32_OperatingSystem, objItem, colItems, Role
	Set Win32_OperatingSystem = oSWbemServices.ExecQuery("select Name from Win32_OperatingSystem")
	For each objItem in Win32_OperatingSystem
		If node.GetAttributeNode("force").Value = "true" Then
			rc = objItem.Win32ShutDown(6,0)
			If rc = 0 Then
				WriteLog "reboot","Forced reboot","success",""
			Else
				WriteLog "reboot","Forced reboot","failure","rc=" & rc
			End If
		Else
			rc = objItem.Win32ShutDown(2,0)
			If rc = 0 Then
				WriteLog "reboot","Reboot","success",""
			Else
				WriteLog "reboot","Reboot","failure","rc=" & rc
			End If
		End If
		Set Win32_OperatingSystem = Nothing
		Exit Function
	Next
	WriteLog "reboot",,"failure","general failure on Win32_OperatingSystem"
	Set Win32_OperatingSystem = Nothing
End Function

Function Shutdown(node)
	On Error Resume Next
	Dim rc
	Dim Win32_OperatingSystem, objItem, colItems, Role
	Set Win32_OperatingSystem = oSWbemServices.ExecQuery("select Name from Win32_OperatingSystem")
	For each objItem in Win32_OperatingSystem
		If node.GetAttributeNode("force").Value = "true" Then
			rc = objItem.Win32ShutDown(5,0)
			If rc = 0 Then
				WriteLog "shutdown","Forced Shutdown","success",""
			Else
				WriteLog "shutdown","Forced Shutdown","failure","rc=" & rc
			End If
		Else
			rc = objItem.Win32ShutDown(1,0)
			If rc = 0 Then
				WriteLog "shutdown","Shutdown","success",""
			Else
				WriteLog "shutdown","Shutdown","failure","rc=" & rc
			End If
		End If
		Set Win32_OperatingSystem = Nothing
		Exit Function
	Next
	WriteLog "shutdown",,"failure","general failure on Win32_OperatingSystem"
	Set Win32_OperatingSystem = Nothing
End Function

Function FormatDate(DateString)
	' On Error Resume Next
	Dim year, month, Day, time
	If IsNull(DateString) Then
		FormatDate = ""
		Exit Function
	End If
	year = Left(DateString,4)
	month = Mid(DateString,5,2)
	Day = Mid(DateString,7,2)
	time = Mid(DateString,9,2) & ":" & Mid(DateString,11,2) & ":" & Mid(DateString,13,2)
	FormatDate = DateSerial(year,month,day) & " " & CDate(time)
End Function

Sub ServiceSetStartMode(Name,StartMode)
	On Error Resume Next
	Dim oService, blnFound
	blnFound = False
	Set oServices = oSWbemServices.ExecQuery("Select Name, StartMode from win32_service where name='" & Name & "'")
	For Each Item In oServices
		blnFound = True
		rc = Item.ChangeStartMode(StartMode)
		If rc = 0 Then
			WriteLog "servicesetstartmode",Name,"success",StartMode
		Else
			WriteLog "servicesetstartmode",Name,"failure","rc=" & rc
		End If
	Next
	If blnFound = False Then
		WriteLog "servicesetstartmode",Name,"failure","Service does not exist"
	End If
	Set oServices = Nothing
End Sub

Sub ServiceStop(Name)
	On Error Resume Next
	Dim oService, blnFound
	blnFound = False
	Set oServices = oSWbemServices.ExecQuery("Select Name, State from win32_service where name='" & Name & "'")
	For Each Item In oServices
		blnFound = True
		rc = Item.StopService()
		If rc = 0 Then
			WriteLog "servicestop",Name,"success","Service Stopped"
		Else
			WriteLog "servicestop",Name,"failure","rc=" & rc
		End If
	Next
	If blnFound = False Then
		WriteLog "servicestop",Name,"failure","Service does not exist"
	End If
	Set oServices = Nothing
End Sub

Sub ServiceStart(Name)
	On Error Resume Next
	Dim oService, blnFound
	blnFound = False
	Set oServices = oSWbemServices.ExecQuery("Select Name, State from win32_service where name='" & Name & "'")
	For Each Item In oServices
		blnFound = True
		rc = Item.StartService()
		If rc = 0 Then
			WriteLog "servicestart",Name,"success","Service Started"
		Else
			WriteLog "servicestart",Name,"failure","rc=" & rc
		End If
	Next
	If blnFound = False Then
		WriteLog "servicestart",Name,"failure","Service does not exist"
	End If
	Set oServices = Nothing
End Sub

Sub CopyFile(Name,Destination)
	On Error Resume Next

	Name = Replace(Name, "%COMPUTERNAME%", ComputerName)
	Destination = Replace(Destination, "%COMPUTERNAME%", ComputerName)
	If oFSO.FileExists(Name) = False Then
		WriteLog "copyfile",Name,"failure","File does not exist"
		Exit Sub
	End If

'	If oFSO.FolderExists(Destination) = False Then
'		WriteLog "copyfile",Name,"failure","Cannot access folder " & Destination
'		Exit Sub
'	End If

	Err.Clear
	oFSO.CopyFile Name, Destination, True
	If Err.Number <> 0 Then
		WriteLog "copyfile",Name,"failure","Error copying file to '" & Destination & "', Error = " & Err.Number
		Exit Sub
	End If

	WriteLog "copyfile",Name,"success",Destination

End Sub

Sub CopyFolder(Source,Destination)
	On Error Resume Next
	Destination = Replace(Destination, "%COMPUTERNAME%", ComputerName)
	If oFSO.FolderExists(Source) = False Then
		WriteLog "copyfolder",Source,"failure","Folder does not exist"
		Exit Sub
	End If

	If oFSO.FolderExists(Destination) = False Then
		WriteLog "copyfolder",Destination,"failure","Cannot access folder " & Destination
		Exit Sub
	End If

	Err.Clear
	oFSO.CopyFolder Source, Destination, True
	If Err.Number <> 0 Then
		WriteLog "copyfolder",Source,"failure","Error copying folder to '" & Destination & "', Error = " & Err.Number
		Exit Sub
	End If

	WriteLog "copyfolder",Source,"success",Destination
End Sub

Sub CreateFolder(FolderName)
	On Error Resume Next

	Destination = Replace(FolderName, "%COMPUTERNAME%", ComputerName)
	If oFSO.FolderExists(Destination) = True Then
		WriteLog "createfolder",Destination,"success","Folder already exists"
		Exit Sub
	End If

	Err.Clear 
	oFSO.CreateFolder Destination
	If Err.Number <> 0 Then
		WriteLog "createfolder",Destination,"failure","Error creating folder, Err.Number = " & Err.Number & ", err.description=" & Err.Description 
		Exit Sub
	End If

	WriteLog "createfolder",Destination,"success","sweet"
End Sub

Sub DeleteFolder(FolderName)
	On Error Resume Next
	FolderName = Replace(FolderName, "%COMPUTERNAME%", ComputerName)
	If oFSO.FolderExists(FolderName) = False Then
		WriteLog "deletefolder",FolderName,"failure","Folder does not exist"
		Exit Sub
	End If

	Err.Clear
	oFSO.DeleteFolder FolderName, True
	If Err.Number <> 0 Then
		WriteLog "deletefolder",FolderName,"failure","Error deleting folder, Err=" & Err.Number & ", Description=" & Err.description
		Exit Sub
	End If

	WriteLog "deletefolder",FolderName,"success","Folder deleted"

End Sub

Sub DeleteFile(FileName1)
	On Error Resume Next
	Dim FileName
	FileName = Replace(FileName1, "%COMPUTERNAME%", ComputerName)
	If oFSO.FileExists(FileName) = False Then
		WriteLog "deletefile",FileName,"failure","File does not exist"
		Exit Sub
	End If

	Err.Clear
	oFSO.DeleteFile FileName, True
	If Err.Number <> 0 Then
		WriteLog "deletefile",FileName,"failure","Error deleting file, Err=" & Err.Number & ", Description=" & Err.description
		Exit Sub
	End If

	WriteLog "deletefile",FileName,"success","File deleted"

End Sub

Sub Run(node)

	On Error Resume Next
	cmd = Replace(Node.getAttributeNode("cmd").value, "%COMPUTERNAME%", ComputerName)
	
	Ret = oShell.Run (cmd,,Node.getAttributeNode("wait").value)
	if Ret <> 0 Then
		WriteLog "run",cmd,"failure","Return code=" & ret
	Else
		WriteLog "run",cmd,"success",""
	End If

End Sub

Sub PsExec(Name)
	On Error Resume Next
	
	cmd = "psexec.exe \\" & ComputerName & " " & Name
	ret = oshell.Run (cmd,0,True)
	if err.number <> 0 Then
		WriteLog "psexec",Name,"failure","Error Number: " & Err.Number
		exit sub
	End If
	
	if ret <> 0 Then
		WriteLog "psexec",Name,"failure","Return code: " & ret
		Exit Sub
	End If

	WriteLog "psexec",Name,"success",""
	
End Sub

Sub verifydnstocomputername()
	On Error Resume Next
	Dim LocalComputerName, oComputer

	Set oComputer = oSWbemServices.ExecQuery("Select Name from Win32_ComputerSystem")
	For Each Item In oComputer
		LocalComputerName = Item.Name
	Next
	Set oComputer = Nothing

	If UCase(LocalComputerName) = UCase(ComputerName) Then
		WriteLog "verifydnstocomputername",LocalComputerName,"success",""
	Else
		WriteLog "verifydnstocomputername",LocalComputerName & " != " & ComputerName,"failure","DNS Error"
	End If
	Set oComputer = Nothing

End Sub

Sub GetDNSServerSearchOrder(Node)
	On Error Resume Next
	Dim aOld
	Found = False
	
	Set oNetwork = oSWbemServices.ExecQuery("Select * from Win32_NetworkAdapterConfiguration " & Node.getAttributeNode("query").Value)
	SearchOrder = ""
	ServiceName = ""

	For Each Item in oNetwork
		Found = True
		ServiceName = Item.ServiceName
		aOld = Item.DNSServerSearchOrder
		For x = lBound(aOld) to UBound(aOld)
			SearchOrder = SearchOrder & aOld(x) & ", "
		Next
	Next
	If Found = False Then
		WriteLog "getdnsserversearchorder",Node.getAttributeNode("query").Value,"failure","Device does not exist"
	Else
		WriteLog "getdnsserversearchorder",Node.getAttributeNode("query").Value,"success","ServiceName=" & ServiceName & ", DNSServerSearchOrder=" & SearchOrder
		FailureSummary = SearchOrder
	End If

	Set oNetwork = Nothing
End Sub

Sub DHCPSettings(Node)
	On Error Resume Next
	Dim aOld
	Found = False
	DHCPEnabled = ""
	DHCPServer = ""
	ServiceName = ""
	
	Set oNetwork = oSWbemServices.ExecQuery("Select * from Win32_NetworkAdapterConfiguration " & Node.getAttributeNode("query").Value)
	SearchOrder = ""
	For Each oItem in oNetwork
		Found = True
		ServiceName = oItem.ServiceName
		DHCPServer = oItem.DHCPServer
		DHCPEnabled = oItem.DHCPEnabled
	Next
	If Found = False Then
		WriteLog "dhcpsettings",Node.getAttributeNode("query").Value,"failure","Device does not exist"
	Else
		WriteLog "dhcpsettings",Node.getAttributeNode("query").Value,"success","Service Name=" & ServiceName & ", DHCP Enabled=" & DHCPEnabled & ", DHCP Server=" & DHCPServer
		Field2 = DHCPServer
	End If

	Set oNetwork = Nothing
End Sub

Sub macaddress(Node)
	On Error Resume Next
	Dim aOld
	Found = False
	ServiceName = ""
	MAC = ""

	Set oNetwork = oSWbemServices.ExecQuery("Select * from Win32_NetworkAdapterConfiguration " & Node.getAttributeNode("query").Value)
	SearchOrder = ""
	For Each oItem in oNetwork
		Found = True
		ServiceName = oItem.ServiceName
		MAC = oItem.MACAddress
	Next
	If Found = False Then
		WriteLog "macaddress",Node.getAttributeNode("query").Value,"failure","Device does not exist"
	Else
		WriteLog "macaddress",Node.getAttributeNode("query").Value,"success","Service Name=" & ServiceName & ", MACAddress=" & MAC
		Field2 = MAC
	End If

	Set oNetwork = Nothing
End Sub

Sub SetDNSServerSearchOrder(ServiceName, DNSServerSearchOrder)
	On Error Resume Next
	Dim aOld, aNew
	Found = False
	
'	Set oNetwork = oSWbemServices.ExecQuery("Select DNSServerSearchOrder from Win32_NetworkAdapterConfiguration where ServiceName = '" & ServiceName & "' and IPEnabled=true")
	Set oNetwork = oSWbemServices.ExecQuery("Select DNSServerSearchOrder from Win32_NetworkAdapterConfiguration where IPEnabled=true")
	aNew = Split(DNSServerSearchOrder,",")
	For Each Item in oNetwork
		Found = True
		ret = Item.SetDNSServerSearchOrder(aNew)
		If ret = 0 Then
			WriteLog "setdnsserversearchorder",ServiceName,"success","DNSServerSearchOrder=" & DNSServerSearchOrder
		Else
			WriteLog "setdnsserversearchorder",ServiceName,"failure","DNSServerSearchOrder=" & DNSServerSearchOrder & ", Return Code=" & ret
		End If
	Next
	If Found = False Then
		WriteLog "setdnsserversearchorder",ServiceName,"failure","ServiceName does not exist"
	End If		

	Set oNetwork = Nothing
End Sub

Sub SetDNSSuffixSearchOrder(DNSDomainSuffixSearchOrder)
	On Error Resume Next
	Set oNetwork = oSWbemServices.Get("Win32_NetworkAdapterConfiguration")

	aNew = Split(DNSDomainSuffixSearchOrder,",")
	ret = oNetwork.SetDNSSuffixSearchOrder(aNew)
	If ret = 0 Then
		WriteLog "setdnssuffixsearchorder",ServiceName,"success","DNSDomainSuffixSearchOrder=" & DNSDomainSuffixSearchOrder
	Else
		WriteLog "setdnssuffixsearchorder",ServiceName,"failure","DNSDomainSuffixSearchOrder=" & DNSDomainSuffixSearchOrder & ", Return Code=" & ret
	End If
	Set oNetwork = Nothing
End Sub

Sub AppendDNSSuffixSearchOrder(ServiceName, DNSDomainSUffixSearchOrder)
	On Error Resume Next
	Set oNetwork = oSWbemServices.Get("Win32_NetworkAdapterConfiguration")
	Set oInstance = oSWbemServices.ExecQuery("Select DNSDomainSuffixSearchOrder from Win32_NetworkAdapterConfiguration where ServiceName = '" & ServiceName & "' and IPEnabled=true")

	Dim aOld, aNew
	ReDim aOld(-1)
	
	For Each Item In oInstance
		If Not IsNull(Item.DNSDomainSuffixSearchOrder) Then
			For Each entry In Item.DNSDomainSuffixSearchOrder
				ReDim Preserve aOld(UBound(aOld)+1)
				aOld(UBound(aOld)) = entry
			Next
		End If
	Next

	aNew = Split(DNSDomainSuffixSearchOrder,",")

	For x = LBound(aNew) To UBound(aNew)
		Found = False
		For y = LBound(aOld) To UBound(aOld)
			If UCase(aNew(x)) = UCase(aOld(y)) Then	found = True
		Next
		If Found = False Then
			ReDim Preserve aOld(UBound(aOld)+1)
			aOld(UBound(aOld)) = aNew(x)
		End If
	Next
	
	Dim NewDNSDomainSuffixSearchOrder
	For x = LBound(aOld) To UBound(aOld)
		NewDNSDomainSuffixSearchOrder = NewDNSDomainSuffixSearchOrder & aOld(x) & ","
	Next
	
	NewDNSDomainSuffixSearchOrder = Left(NewDNSDomainSuffixSearchOrder,Len(NewDNSDomainSuffixSearchOrder) - 1)
	
	ret = oNetwork.SetDNSSuffixSearchOrder(aOld)
	If ret = 0 Then
		WriteLog "appenddnssuffixsearchorder",ServiceName,"success","DNSDomainSuffixSearchOrder=" & NewDNSDomainSuffixSearchOrder
	Else
		WriteLog "appenddnssuffixsearchorder",ServiceName,"failure","DNSDomainSuffixSearchOrder=" & NewDNSDomainSuffixSearchOrder & ", Return Code=" & ret
	End If
	Set oNetwork = Nothing
	Set oInstance = Nothing
End Sub

Sub SMSGenerateCCR(OutputDirectory)
	On Error Resume Next
	Dim ccrFileName : ccrFileName = g_Temp & ComputerName & ".CCR"
	
	OutputDirectory = AddBackslash(OutputDirectory)

	Set oFile = oFSO.CreateTextFile(ccrFileName,True)
	oFile.WriteLine ""
	oFile.WriteLine "[NT Client Configuration Request]"
	oFile.WriteLine "   Machine Name=" & ComputerName
'	oFile.WriteLine "   IP Address 1=" & IPAddress
'	oFile.WriteLine ""
'	oFile.WriteLine "[IP Address]"
'	oFile.WriteLine "   IP Address 1=" & IPAddress
	oFile.WriteLine ""
	oFile.WriteLine "[IDENT]"
	oFile.WriteLine "    TYPE=Client Config Request File"
	oFile.WriteLine ""

	oFile.Close
	Set oFile = Nothing

	oFSO.CopyFile ccrFileName, OutputDirectory, True
	Err.Clear 
	If Err.Number <> 0 Then
		WriteLog "smsgenerateccr",ccrFileName,"failure","Error copying file to " & OutputDirectory & ", Error=" & Err.Number
		Exit Sub
	End If
	
	Err.Clear
	oFSO.DeleteFile ccrFileName
	If Err.Number <> 0 Then
		WriteLog "smsgenerateccr",ccrFileName,"failure","Error deleting file, Error=" & Err.Number
		Exit Sub
	End If
	
	WriteLog "smsgenerateccr",ccrFileName,"success","OutputDirectory=" & OutputDirectory
End Sub

Sub appendtexttolocalfile(FileName, Text)
	On Error Resume Next
	Const For_Reading = 1
	Const For_Appending = 8
	
	If InStr(FileName,"%COMPUTERNAME%") > 0 Then
		FileName = Replace(Filename,"%COMPUTERNAME%",ComputerName)
	End If
	
	' Check if text is already written to file
	Set oFile = oFSO.OpenTextFile(FileName,For_Reading,False)
	allText = oFile.ReadAll
	oFile.Close
	If InStr(uCase(allText), uCase(Text)) > 0 Then
		WriteLog "appendtexttolocalfile",FileName,"success",Text & " already in file"
		Exit Sub
	End If
	
	Set oFile = oFSO.OpenTextFile(FileName,For_Appending,False)
	If Err.Number <> 0 Then
		WriteLog "appendtexttolocalfile",FileName,"failure","Error opening file: " & Err.Number
		Exit Sub
	End If
	
	oFile.WriteLine ""
	oFile.WriteLine Text
	oFile.Close
	Set oFile = Nothing
	WriteLog "appendtexttolocalfile",FileName,"success",Text & " written to file"
End Sub

Sub issms2003clientinstalled(yesAction, noAction)
	On Error Resume Next
	
	val = GetRegistryKeyValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{FCDC3CDD-F53E-4239-8CA5-BC492942931B}\DisplayName","REG_SZ")
	If val = "SMS Advanced Client" Then
		WriteLog "issms2003clientinstalled","Client is installed","success",""
		if yesAction <> "" Then ProcessVerificationTasks(yesAction)
	Else
		WriteLog "issms2003clientinstalled","Client is not installed","failure",""
		If noAction <> "" Then ProcessVerificationTasks(noAction)
	End If

End Sub

Sub issms2003clientassigned(yesAction, noAction)
	On Error Resume Next
	
	Err.Clear 
	Set oSWbemPolicy = oWbemLocator.ConnectServer (ComputerName, "root\CCM\Policy\Machine")
	If Err.Number <> 0 Then
 		WriteLog "issms2003clientassigned","Client is not assigned","failure",""
 		If noAction <> "" Then ProcessVerificationTasks(noAction)
 		Exit Sub
	End If

	Err.Clear
	Set oClientActions = oSWbemPolicy.ExecQuery("Select ActionID from CCM_ClientActions")
	If Err.Number <> 0 Then
 		WriteLog "issms2003clientassigned","Client is not assigned","failure",""
 		If noAction <> "" Then ProcessVerificationTasks(noAction)
 		Exit Sub
	End If
	
	Dim x 
	x = 0
	For Each Item In oClientActions
		x = x + 1
	Next

	Set oClientActions = Nothing
	If x < 3 Then
 		WriteLog "issms2003clientassigned","Client is not assigned","failure",""
 		If noAction <> "" Then ProcessVerificationTasks(noAction)
 		Exit Sub
	End If

	WriteLog "issms2003clientassigned","Client is assigned","success",""
	If YesAction <> "" Then ProcessVerificationTasks(yesAction)

End Sub


Sub sms_policyretrieved()
	On Error Resume Next
	
	Err.Clear 
	Set oSWbemPolicy = oWbemLocator.ConnectServer (ComputerName, "root\CCM\Policy\Machine")
	Set oSWbemPolicy = GetWMINamespace(ComputerName,"root\CCM\Policy\Machine")
	If Err.Number <> 0 Then
 		WriteLog "sms_policyretrieved","Could not connect to root\CCM\Policy\Machine","failure",""
 		Exit Sub
	End If

	Err.Clear
	Set oClientActions = oSWbemPolicy.ExecQuery("Select ActionID from CCM_ClientActions")
	If Err.Number <> 0 Then
 		WriteLog "sms_policyretrieved","Could not query CCM_ClientActions","failure",""
 		Exit Sub
	End If
	
	Dim x 
	x = 0
	For Each Item In oClientActions
		x = x + 1
	Next

	Set oClientActions = Nothing
	If x < 11 Then
 		WriteLog "sms_policyretrieved","Client action count less than 11","failure",x
 		Exit Sub
	End If

	WriteLog "sms_policyretrieved","Client has retrieved policy","success","Action count = " & x
	Set oSWbemPolicy = Nothing
	Set oClientActions = Nothing

End Sub

Function GetRegistryKeyValue(Name,KeyType)
	On Error Resume Next
	Dim RetVal, KeySplit, KeyBase, KeyName
	
	KeySplit = Split(Name,"\")
	KeyBase = KeySplit(0)
	KeyName = KeySplit(UBound(KeySplit))
	SubKeyName = Mid(Name,Len(KeyBase)+2,len(Name) - (Len(KeyBase) + Len(KeyName) + 2))
	
	' Constants
	Const HKEY_CLASSES_ROOT = &H80000000
	Const HKEY_CURRENT_USER = &H80000001 
	Const HKEY_LOCAL_MACHINE = &H80000002
	Const HKEY_USERS = &H80000003
	Const HKEY_CURRENT_CONFIG = &H80000005
	Const HKEY_DYN_DATA = &H80000006

	Select Case KeyBase
		Case "HKEY_CLASSES_ROOT"
			sKeyBase = HKEY_CLASSES_ROOT
		Case "HKEY_CURRENT_USER"
			sKeyBase = HKEY_CURRENT_USER
		Case "HKEY_LOCAL_MACHINE"
			sKeyBase = HKEY_LOCAL_MACHINE
		Case "HKEY_USERS"
			sKeyBase = HKEY_USERS
		Case "HKEY_CURRENT_CONFIG"
			sKeyBase = HKEY_CURRENT_CONFIG
		Case "HKEY_DYN_DATA"
			sKeyBase = HKEY_DYN_DATA
	End Select

	Select Case KeyType
		Case "REG_SZ"
			retVal = oRegistry.GetStringValue (sKeyBase, SubKeyName,KeyName, RegVal)
		Case "REG_DWORD"
			retVal = oRegistry.GetDWordValue(sKeyBase,SubKeyName,KeyName, RegVal)
	End Select
	
	If retVal <> 0 Then
		GetRegistryKeyValue = vbNullString
		Exit Function
	End If

	GetRegistryKeyValue = RegVal
	
End Function

Sub Quit()
	On Error Resume Next
	' Finish the LogFile
	Dim oOutFile

	WriteLog "variable","timestarted","information",TimeStarted
	WriteLog "variable","networkstatus","information",NetworkStatus
	WriteLog "variable","ipaddress","information",IPAddress
	WriteLog "variable","status","information",status
	WriteLog "variable","field1","information",field1
	WriteLog "variable","field2","information",field2
	WriteLog "variable","field3","information",field3
	WriteLog "variable","successpoints","information",SuccessPoints
	WriteLog "variable","failurepoints","information",FailurePoints

	Status = "Finished"
	WriteLog "quit","Finishing","success","All Done"
	
'	UpdateHeaderInformation(OutputFileName)
	Set stdOut = WScript.stdOut
	if len(LogFile.xml) > 4000 then
		Set oOutFile = oFSO.OpenTextFile (g_currentDirectory & "_" & ComputerName & ".xml",2,True)
		oOutFile.Write LogFile.xml
		oOUtFile.close
		Set oOutFile = Nothing
		stdOut.Write "FILE:" & g_currentDirectory & "_" & ComputerName & ".xml"
	Else
		stdOut.write LogFile.xml
	End if 
	stdOut.Close
	Cleanup
End Sub

Sub fileshowstringaftermatch(filename, searchstring)
	On Error Resume Next
	Const For_Reading = 1
	Dim bFoundMatch : bFoundMatch = True
	If InStr(FileName,"%COMPUTERNAME%") > 0 Then
		FileName = Replace(Filename,"%COMPUTERNAME%",ComputerName)
	End If
	
	' Check if text is already written to file
	Set oFile = oFSO.OpenTextFile(FileName,For_Reading,False)
	do while oFile.AtEndofstream <> True
		line = ofile.readline()
		If InStr(uCase(line), UCase(searchstring)) > 0 Then
			bFoundMatch = True
			location = Replace(Mid(line,InStr(uCase(line), uCase(searchstring)) + Len(searchstring)),"]LOG]!>","")
			location = Replace(location,"'","")
			If InStr(line,"10S01") > 0 Then
				WriteLog "fileshowstringaftermatch",FileName,"success",Location
			Else
				WriteLog "fileshowstringaftermatch",FileName,"failure",Location
			End If 
		End If
	Loop
	oFile.Close

	field1 = Location

	Set oFile = Nothing
End Sub


Sub filecontainsstring(filename, searchstring)
	On Error Resume Next
	Const For_Reading = 1
	If InStr(FileName,"%COMPUTERNAME%") > 0 Then
		FileName = Replace(Filename,"%COMPUTERNAME%",ComputerName)
	End If
	
	' Check if text is already written to file
	Set oFile = oFSO.OpenTextFile(FileName,For_Reading,False)
	allText = oFile.ReadAll
	oFile.Close

	If InStr(uCase(allText), uCase(searchstring)) > 0 Then
		WriteLog "filecontainsstring",FileName,"success",searchstring & " in file"
	Else
		WriteLog "filecontainsstring",FileName,"failure",searchstring & " is not in file"
	End If
	
	Set oFile = Nothing
End Sub

Sub filecontainserrorstring(filename, searchstring)
	On Error Resume Next
	Dim bFoundMatch : bFoundMatch = False
	Const For_Reading = 1
	If InStr(FileName,"%COMPUTERNAME%") > 0 Then
		FileName = Replace(Filename,"%COMPUTERNAME%",ComputerName)
	End If
	' Check if text is already written to file
	Err.clear 
	Set oFile = oFSO.OpenTextFile(FileName,For_Reading,False)
	if err.number <> 0 then
		WriteLog "filecontainserrorstring",FileName,"failure","Problem opening file, err.number=" & err.number & ", err.description=" & err.description
		Exit Sub
	end if
	do while oFile.AtEndofstream <> True
		line = ofile.readline()
		If InStr(uCase(line), uCase(searchstring)) > 0 Then
			if instr(uCase(line), "ERROR RUNNING THE PREVIOUS COMMAND") > 0 Then
				WriteLog "filecontainserrorstring",FileName,"failure",previousline
			End If
			bFoundMatch = True
			WriteLog "filecontainserrorstring",FileName,"failure",line
		End If
		previousline = line

	Loop
	oFile.Close
	Set oFile = Nothing

	If bFoundMatch = False Then
		WriteLog "filecontainserrorstring",FileName,"success","no matches found for string: " & searchstring
	End If
	
End Sub

Sub filereadline(filename, searchstring)
	On Error Resume Next
	Const For_Reading = 1
	Dim oFile, alltext, matchText
	
	matchText = ""
	If InStr(FileName,"%COMPUTERNAME%") > 0 Then
		FileName = Replace(Filename,"%COMPUTERNAME%",ComputerName)
	End If

	Set oFile = oFSO.OpenTextFile(FileName,For_Reading,False)

	Do While oFile.AtEndOfStream <> True
		' Check if text is already written to file
		allText = oFile.ReadLine()
		if InStr(allText,searchString) > 0 Then
			matchText = allText
		End If
	Loop

	If matchText = "" Then
		WriteLog "filereadline",FileName,"failure",searchstring & " is not in file"
	Else
		WriteLog "filereadline",FileName,"success",matchText
	End If

	oFile.Close
	Set oFile = Nothing
End Sub

Function InGroup(GroupName)
	On Error Resume Next
	Dim aMemberOf
	ReDim aMemberOf(-1)

	dn = distinguishedName(ComputerName)
	Set objComputer = GetObject ("LDAP://" & dn)
	objComputer.GetInfo
	aMemberOf = objComputer.GetEx("memberOf")
	If UBound(aMemberOf) < 0 Then
		Set objComputer = Nothing
		WriteLog "ingroup",GroupName,"failure","Computer is not a member of any groups"
		Exit Function
	End If

	For Each Group In aMemberOf
		FirstComma = Instr(Group,",")
		TestGroupName = Mid(Group,4,FirstComma-4)
		If InStr(TestGroupName,GroupName) > 0 Then
			WriteLog "ingroup",TestGroupName,"success",Group
			Set objComputer = Nothing
			Field1 = TestGroupName
			Exit Function
		End If
	Next

	WriteLog "ingroup",GroupName,"failure","Computer is not a member"
	Set objComputer = Nothing
End Function

Function SMSGetAssignedsite
	On Error Resume Next
	Err.Clear
'	Set oSMSClient = GetObject("winmgmts://" & ComputerName & "/root/ccm:SMS_Client")
	Dim ns
	Set ns = GetWMINamespace(ComputerName,"root\ccm")
	Set oSMSClient = ns.get("SMS_Client")
	If Err.NUmber <> 0 Then
		WriteLog "smsgetassignedsite","","failure",Err.Number & " - " & Err.Description
		Exit Function
	End If

	Err.CLear
	Set Result = oSMSClient.ExecMethod_("GetAssignedSite")
	If Err.Number <> 0 Then	
		WriteLog "smsgetassignedsite","","failure",Err.Number & " - " & Err.Description
	End If	
	
	WriteLog "smsgetassignedsite","","success",Result.sSiteCode

	Set Result = Nothing
	Set osMSClient = Nothing
End Function

Function SMSSetAssignedsite(SiteCode)
	On Error Resume Next
	Err.Clear
'	Set oSMSClient = GetObject("winmgmts://" & ComputerName & "/root/ccm:SMS_Client")
	Dim ns
	Set ns = GetWMINamespace(ComputerName,"root\ccm")
	Set oSMSClient = ns.get("SMS_Client")
	If Err.NUmber <> 0 THen
		WriteLog "smssetassignedsite","","failure",Err.Number & " - " & Err.Description
		Exit Function
	End If

	Err.CLear
	Set inParam = oSMSClient.Methods_.Item("SetAssignedSite").inParameters.SpawnInstance_()
	inParam.sSiteCode = SiteCode
	Set Result = oSMSClient.ExecMethod_("SetAssignedSite",inParam)
	If Err.Number <> 0 Then	
		WriteLog "smssetassignedsite","","failure",Err.Number & " - " & Err.Description
	Else
		WriteLog "smssetassignedsite",Result.sSiteCode,"success","New site code is " & SiteCode
	End If	

	Set inParam = Nothing
	Set Result = Nothing
	Set osMSClient = Nothing
End Function



Function SMSResetPolicy
	On Error Resume Next
	Err.Clear
'	Set oSMSClient = GetObject("winmgmts://" & ComputerName & "/root/ccm:SMS_Client")
	Dim ns
	Set ns = GetWMINamespace(ComputerName,"root\ccm")
	Set oSMSClient = ns.get("SMS_Client")
	If Err.NUmber <> 0 THen
		WriteLog "smsresetpolicy","","failure",Err.Number & " - " & Err.Description
		Exit Function
	End If

	Err.CLear
	oSMSClient.ResetPolicy(0)
	If Err.Number <> 0 Then	
		WriteLog "smsresetpolicy","","failure",Err.Number & " - " & Err.Description
	Else
		WriteLog "smsresetpolicy","","success","Policy reset"
	End If	
	Set osMSClient = Nothing
End Function

Function sms_repairclient
	On Error Resume Next
	Err.Clear
'	Set oSMSClient = GetObject("winmgmts://" & ComputerName & "/root/ccm:SMS_Client")
	Dim ns
	Set ns = GetWMINamespace(ComputerName,"root\ccm")
	Set oSMSClient = ns.get("SMS_Client")
	If Err.NUmber <> 0 THen
		WriteLog "sms_repairclient","","failure",Err.Number & " - " & Err.Description
		Exit Function
	End If

	Err.CLear
	oSMSClient.RepairClient
	If Err.Number <> 0 Then	
		WriteLog "sms_repairclient","","failure",Err.Number & " - " & Err.Description
	Else
		WriteLog "sms_repairclient","","success","Client repairing"
	End If	
	Set osMSClient = Nothing
End Function

Function SMSRequestMachinePolicy
	On Error Resume Next
	Err.Clear
'	Set oSMSClient = GetObject("winmgmts://" & ComputerName & "/root/ccm:SMS_Client")
	Dim ns
	Set ns = GetWMINamespace(ComputerName,"root\ccm")
	Set oSMSClient = ns.get("SMS_Client")
	If Err.NUmber <> 0 THen
		WriteLog "smsrequestmachinepolicy","","failure",Err.Number & " - " & Err.Description
		Exit Function
	End If

	Err.CLear
	oSMSClient.RequestMachinePolicy(0)
	If Err.Number <> 0 Then	
		WriteLog "smsrequestmachinepolicy","","failure",Err.Number & " - " & Err.Description
	Else
		WriteLog "smsrequestmachinepolicy","","success","Machine policy retrieval successful"
	End If	
	Set osMSClient = Nothing
End Function


Sub SMSWMIQuery(Node)
	On Error Resume Next

	Query = Replace(Node.GetAttributeNode("query").Value, "%COMPUTERNAME%", ComputerName)

'	Set oServices = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & Node.GetAttributeNode("host").Value & "\root\sms\site_" & Node.GetAttributeNode("sitename").Value)
	Set oServices = GetWMINamespace(Node.GetAttributeNode("host").Value,"root\sms\site_" & Node.GetAttributeNode("sitename").Value)
	If Err.Number <> 0 Then
		WriteLog "smswmiquery","Failed to connect to host " & Node.GetAttributeNode("host").Value,"failure",Err.number & ", " & Err.description
		Field1 = "0"
		Exit Sub
	End If

	Set oResults = oServices.ExecQuery(Query)
'	For Each Result in oResults
'		If Len(Result.Client) < 1 Then
'			Client = "0"
'		Else
'			Client = Result.Client
'		End if
'		WriteLog "smswmiquery",Query,"success",Client
'		Field1 = Client
'	Next

	If oResults.Count=0 Then
		writeLog "smswmiquery",Query,"failure","0"
		field1 = "0"
	Else
		WriteLog "smswmiquery",Query,"success","1"
		Field1 = "1"
	End If

	Set oResults = Nothing
	Set oServices = Nothing	
	
End Sub

Sub Sleep(Node)
	On Error Resume Next
	
	Dim SleepTime
	SleepTime = cInt(Node.GetAttributeNode("time").Value)
	WScript.Sleep SleepTime * 1000
	WriteLog "sleep",SleepTime,"sucess","Slept for " & SleepTime & " seconds."

End Sub

Sub nbtstat_verifyhostname(node)
	On Error Resume Next
	Dim oExec, cmd
	
	cmd = "nbtstat -a " & IPAddress
	Set oExec = oShell.Exec (cmd)
	If Err.Number <> 0 Then
		WriteLog "nbtstat_verifyhostname",cmd,"failure","Error executing: " & Err.Number & ", " & Err.Description
		Exit Sub
	End If
	Do While oExec.Status = 0
		WScript.Sleep 100
	Loop
	StdOutAll = uCase(oExec.StdOut.ReadAll)

'field2 = stdoutall

	Set regEX = New RegExp 
	regex.Pattern = "\s\S.*<00>  GROUP"
	regEx.Ignorecase = True
	regEX.Global = True
	
	Set Matches = regEx.Execute(stdOutAll)
	For Each Match In Matches
		str = trim(replace(Match.Value,"<00>  GROUP",""))
	Next

	If Len(str) > 0 Then
		WriteLog "nbtstat_verifyhostname",cmd,"success",str
	Else
		WriteLog "nbtstat_verifyhostname",cmd,"failure",str
	End If

	Field1 = uCase(str)

	Set oExec = Nothing
	Set regEX = Nothing
	Set Matches = Nothing
End Sub

Sub xml_querynode(node)
	On Error Resume Next
	set m = Root.SelectNodes(node.getAttributeNode("query").Value)
	If m.length > 0 Then ' Success
		Set m = Nothing
		WriteLog "xml_querynode",node.getAttributeNode("query").Value,"success","RecordCount > 0, Processing TrueAction: " & node.getAttributeNode("trueaction").Value
		ProcessVerificationTasks(node.getAttributeNode("trueaction").Value)
	Else ' Failure
		Set m = nothing
		WriteLog "xml_querynode",node.getAttributeNode("query").Value,"success","RecordCount = 0, Processing FalseAction: " & node.getAttributeNode("falseaction").Value
		ProcessVerificationTasks(node.getAttributeNode("falseaction").Value)
	End If

End Sub

Sub radia_enumeratepackages
	On Error Resume Next
	Err.Clear
	Set oFld = oFSO.GetFolder("\\" & ComputerName & "\d$\SERVER\RADIA\Radia Integration Server\etc\rps\RADSTAGE\RADSTAGE\RADSTAGE\ZSERVICE")
	If Err.Number <> 0 Then
		WriteLog "radia_enumeratepackages","Error connectin to radia stage directory","failure",Err.Number & ", " & err.description
		Exit Sub
	End If

	Set oOutFile = oFSO.CreateTextFile("D:\michael\servers\" & ComputerName & ".txt", True)
	
	For Each folder in oFld.SubFolders
		oOutFile.WriteLine ComputerName & vbTab & folder.Name
	Next

	WriteLog "radia_enumeratepackages","","success","Enumeration complete - have a nice day!"

	oOutFile.Close
	Set oOutFile = Nothing
	Set oFld = Nothing
End Sub

Sub sms_addtocollection(Node)
	On Error Resume Next
	Dim lLocator
	Dim lServices

	' Connect to WMI
	Err.Clear
	Set lServices = oWbemLocator.ConnectServer(Node.getAttributeNode("server").value, "root/sms/site_" & Node.getAttributeNode("sitecode").value)
	If Err.Number <> 0 Then
		WriteLog "sms_addtocollection","WMI Connection Error to server: " & Node.getAttributeNode("server").value,"failure",Err.Number & ", " & err.description
		Exit Sub
	End If

	' Get the collection ID
	CollectionID = ""
	Set Collections = lServices.ExecQuery("Select * From SMS_Collection where Name='" & Node.getAttributeNode("collectionname").value & "' ")
	For Each Collection In Collections
			CollectionID = Collection.CollectionID
	Next
	Set Collections = Nothing

	' Test we have a collection
	If CollectionID = "" Then
		WriteLog "sms_addtocollection","Failed to connect to collection","failure",Node.getAttributeNode("collectionname").value
		Exit Sub
	End If

	' Connect to the collection
	Err.Clear
	Set Collection = lServices.Get("SMS_Collection.CollectionID='" & CollectionID & " '")
	If Err.Number <> 0 Then
		WriteLog "sms_addtocollection","Failed to connect to collection","failure",""
		Exit Sub
	End If

	Set CollectionRule = lServices.Get("SMS_CollectionRuleDirect").SpawnInstance_()
	CollectionRule.ResourceClassName = "SMS_R_SYSTEM"

	' Get the resource id of the computer
	Set oSystemSet = lServices.ExecQuery("Select * from SMS_R_System where NetbiosName='" & ComputerName & "' ")
	For Each oSystem in oSystemSet
			ResourceID = oSystem.ResourceID
	Next

	' Test we have a ResourceID
	If ResourceID = "" Then
		WriteLog "sms_addtocollection","Failed to get resource id","failure",""
		Exit Sub
	End If

	CollectionRule.RuleName = ComputerName
	CollectionRule.ResourceID = ResourceID

	'add the rule to the collection
	Collection.AddMembershipRule CollectionRule

	WriteLog "sms_addtocollection","CollectionID=" & CollectionID & ", Collection=" & Node.getAttributeNode("collectionname").value,"success","ResourceID=" & ResourceID

	Set CollectionRule = nothing
	Set Collection = Nothing
	'Update Collection Membership
'	UpdateCollection CollectionID

End Sub

Sub sms_removefromcollection(Node)
	On Error Resume Next
	Dim lLocator
	Dim lServices

	' Connect to WMI
	Err.Clear
	Set lServices = oWbemLocator.ConnectServer(Node.getAttributeNode("server").value, "root/sms/site_" & Node.getAttributeNode("sitecode").value)
	If Err.Number <> 0 Then
		WriteLog "sms_removefromcollection","WMI Connection Error to server: " & Node.getAttributeNode("server").value,"failure",Err.Number & ", " & err.description
		Exit Sub
	End If

	' Get the collection ID
	CollectionID = ""
	Set Collections = lServices.ExecQuery("Select * From SMS_Collection where Name='" & Node.getAttributeNode("collectionname").value & "' ")
	For Each Collection In Collections
			CollectionID = Collection.CollectionID
	Next
	Set Collections = Nothing

	' Test we have a collection
	If CollectionID = "" Then
		WriteLog "sms_removefromcollection","Failed to connect to collection","failure",Node.getAttributeNode("collectionname").value
		Exit Sub
	End If

	' Connect to the collection
	Err.Clear
	Set Collection = lServices.Get("SMS_Collection.CollectionID='" & CollectionID & " '")
	If Err.Number <> 0 Then
		WriteLog "sms_removefromcollection","Failed to connect to collection","failure",""
		Exit Sub
	End If

	' Get the resource id of the computer
	Set oSystemSet = lServices.ExecQuery("Select * from SMS_R_System where NetbiosName='" & ComputerName & "' ")
	For Each oSystem in oSystemSet
			ResourceID = oSystem.ResourceID
	Next
	Set oSystemSet = Nothing

	' Test we have a ResourceID
	If ResourceID = "" Then
		WriteLog "sms_removefromcollection","Failed to get resource id","failure",""
		Exit Sub
	End If

	' Connect to the collection
	if Not (isnull(Collection.CollectionRules)) Then
		Rules = Collection.CollectionRules
		For Each Rule In Rules
			tmpResid = ""
			tmpResid = rule.resourceid
			If tmpResid = ResourceID Then
				Err.Clear
				Collection.DeleteMembershipRule Rule
				If Err.Number <> 0 Then
					WriteLog "sms_removefromcollection","CollectionID=" & CollectionID & ", Collection=" & Node.getAttributeNode("collectionname").value & ", ResourceID=" & ResourceID,"failure","err.number=" & Err.Number & ", err.description=" & Err.Description 
					Exit for
				Else
					WriteLog "sms_removefromcollection","CollectionID=" & CollectionID & ", Collection=" & Node.getAttributeNode("collectionname").value,"success","ResourceID=" & ResourceID
					Exit for
				End If
			End If
		Next
		Set Collection = Nothing
		Exit Sub
	End If
	
	WriteLog "sms_removefromcollection","CollectionID=" & CollectionID & ", Collection=" & Node.getAttributeNode("collectionname").value,"failure","ResourceID=" & ResourceID & " not found in collection"

	Set Collection = Nothing
	'Update Collection Membership
'	UpdateCollection CollectionID

End Sub

Sub sms_rerunadvert(node)
	On Error Resume Next

	strAdvID = node.getAttributeNode("advertid").Value

	' Get the ID of the ScheduledMessage on the target machine
	strSchMsgID = GetAdvSchMsgID(strAdvID)
	If strSchMsgID = "" Then
		WriteLog "sms_rerunadvert","Unable to get ScheduledMessageID","failure",strAdvID
		exit Sub
	End If

	' Then make sure the program can be rerun
	ret = SetRerunBehavior(strAdvID, strOldRerunBehavior, "RerunAlways")
	If ret < 0 then
		WriteLog "sms_rerunadvert","Unable to set RerunBehaviour","failure",strSchMsgID
		exit Sub
	End If

	' Invoke SMS_Client.TriggerSchedule method
	ret = TriggerSchedule(strSchMsgID)
	If ret < 0 then
		WriteLog "sms_rerunadvert","Unable to trigger advertisement","failure",strSchMsgID
		exit Sub
	end If

	' give the client some time to create the execution request
	  wscript.sleep 5000

	' Reset the ADV_RepeatRunBehavior state to the previous value
	ret = SetRerunBehavior(strAdvID, strDummy, strOldRerunBehavior)
	If ret < 0 Then
		WriteLog "sms_rerunadvert","Unable to set RerunBehaviour","failure",strSchMsgID
		exit Sub
	end if

	WriteLog "sms_rerunadvert","Successfully triggered rerun","success",strSchMsgID

End Sub

Function GetAdvSchMsgID(strAdvID)
	On Error Resume Next
	err.clear
    GetAdvSchMsgID=""

    ' Connect to the actual configuration policy via WMI
'    set objNMS = GetObject("winmgmts://" & ComputerName & "/root/ccm/policy/machine/actualconfig")
	Set objNMS = GetWMINamespace(ComputerName,"root\ccm\policy\machine\actualconfig")
    if (Err.number <> 0) Then
        Exit Function
    end If

    ' find the matching ScheduledMessage
    ' Cannot use like operator on machines lower than XP/W2K3, so must first get all
    Set objScheds = objNMS.ExecQuery("select * from CCM_Scheduler_ScheduledMessage")
    if (Err.number <> 0) then
        Exit Function
    end if

    ' Search for match by string comparison
    For each objSched in objScheds
        If Instr(objSched.ScheduledMessageID, strAdvID) > 0 then
           GetAdvSchMsgID = objSched.ScheduledMessageID
           exit For
        End If
    Next

    Set objNMS = Nothing
	Set objScheds = Nothing

End Function

function SetRerunBehavior(strAdvID, strOldRerunBehavior, strRerunBehavior)
	On Error Resume Next
   
'	set objNMS = GetObject("winmgmts://" & ComputerName & "/root/ccm/policy/machine/actualconfig")
	Set objNMS = GetWMINamespace(ComputerName,"root\ccm\policy\machine\actualconfig")
    if (Err.number <> 0) then
        SetRerunBehavior = -1
        Exit Function
    end if

   Set objScheds = objNMS.ExecQuery("select * from CCM_SoftwareDistribution where ADV_AdvertisementID = '" & strAdvID & "' and PRG_DependentPolicy <> True")
   if (Err.number <> 0) then
      SetRerunBehavior = -2
      exit function
   end if

   ' there is only one item in the collection
   if (objScheds.Count <> 1) then
	  SetRerunBehavior = -3
      exit function
   end if

   for each objSched in objScheds
       strOldRerunBehavior = objSched.ADV_RepeatRunBehavior
       objSched.ADV_RepeatRunBehavior = strRerunBehavior
       objSched.Put_ 0
       if (Err.number <> 0) then
          SetRerunBehavior = -4
          exit function
       end if
   Next
   SetRerunBehavior = 0

	Set objNMS = Nothing
	Set objScheds = Nothing
end Function


Function TriggerSchedule(strSchMsgID)

'    set objNMS = GetObject("winmgmts://" & ComputerName & "/root/ccm")
	Set objNMS = GetWMINamespace(ComputerName,"root\ccm")
    if (Err.number <> 0) Then
        TriggerSchedule = -1
        Exit Function
    end if

    dim objSMSClient
    Set objSMSClient = objNMS.Get("SMS_Client")
    if (Err.number <> 0) then
        TriggerSchedule = -2
        Exit Function
    end if

    objSMSClient.TriggerSchedule strSchMsgID
    if (Err.number <> 0) then
        TriggerSchedule = -3
        Exit Function
    end if
    TriggerSchedule = 0

	Set objNMS = Nothing
	Set objSMSClient = Nothing
End Function

Function ad_distinguishedname
	
	dn = distinguishedName(ComputerName)
	if dn = "" Then
		WriteLog "ad_distinguishedname","failed","failure","failed to get dn"
		Field1 = ""
	else
		WriteLog "ad_distinguishedname","","success",dn
		Field1 = dn
	End If

End Function

Function ad_addworkstationtoglobalgroup(node)
	On Error Resume Next
	Dim rootDN
	dn = distinguishedName(ComputerName)
	if dn = "" Then
		WriteLog "ad_addworkstationtoglobalgroup","failed","failure","failed to get computer dn"
		Exit Function
	End If

	rootDN = getRoot(dn)
	groupDN = GlobalGroupDN(rootDN,node.getAttributeNode("name").Value)
	if groupDN = "" Then
		WriteLog "ad_addworkstationtoglobalgroup",dn,"failure","failed to get group dn"
		Exit Function
	End If

	Err.Clear
	Set oGroup = GetObject("LDAP://" & groupDN)
	oGroup.Add("LDAP://" & dn)
	If Err.Number = -2147019886 Then
		WriteLog "ad_addworkstationtoglobalgroup",groupDN,"success",dn & " already a member"
		WriteLog "ad_addworkstationtoglobalgroup",groupDN,"failure",dn & " already a member"
		Field1 = "Already a member"
		Err.Clear
	elseif Err.Number = 0 Then
		WriteLog "ad_addworkstationtoglobalgroup",groupDN,"success",dn
	Else
		WriteLog "ad_addworkstationtoglobalgroup",dn,"failure",groupDN & ", " & Err.Number & ", " & Err.Description
	End If

	Set oGroup = Nothing
End Function

Function ad_addusertoglobalgroup(node)
	On Error Resume Next
	Dim rootDN
	dn = distinguishedNameUser(ComputerName)
	if dn = "" Then
		WriteLog "ad_addusertoglobalgroup","failed","failure","failed to get user dn"
		Exit Function
	End If

	rootDN = getRoot(dn)
	groupDN = GlobalGroupDN(rootDN,node.getAttributeNode("name").Value)
	if groupDN = "" Then
		WriteLog "ad_addusertoglobalgroup",dn,"failure","failed to get group dn"
		Exit Function
	End If

	Err.Clear
	Set oGroup = GetObject("LDAP://" & groupDN)
	oGroup.Add("LDAP://" & dn)
	If Err.Number = -2147019886 Then
		WriteLog "ad_addusertoglobalgroup",groupDN,"success",dn & ", already a member"
		field1 = "existing"
		Err.Clear
	elseif Err.Number = 0 Then
		WriteLog "ad_addusertoglobalgroup",groupDN,"success",dn
		field1 = "added"
	Else
		WriteLog "ad_addusertoglobalgroup",dn,"failure",groupDN & ", " & Err.Number & ", " & Err.Description
		field1 = "error"
	End If

	Set oGroup = Nothing
End Function


Function ad_removeworkstationfromglobalgroup(node)
	On Error Resume Next
	Dim rootDN
	dn = distinguishedName(ComputerName)
	if dn = "" Then
		WriteLog "ad_removeworkstationfromglobalgroup","failed","failure","failed to get computer dn"
		Exit Function
	End If

	rootDN = getRoot(dn)
	groupDN = GlobalGroupDN(rootDN,node.getAttributeNode("name").Value)
	if groupDN = "" Then
		WriteLog "ad_removeworkstationfromglobalgroup",dn,"failure","failed to get group dn"
		Exit Function
	End If

	Err.Clear
	Set oGroup = GetObject("LDAP://" & groupDN)
	oGroup.Remove("LDAP://" & dn)
	If Err.Number = -2147016651 Then
		WriteLog "ad_removeworkstationfromglobalgroup",groupDN,"success",dn & ", not a member"
		Err.Clear
	elseif Err.Number = 0 Then
		WriteLog "ad_removeworkstationfromglobalgroup",groupDN,"success",dn
	Else
		WriteLog "ad_removeworkstationfromglobalgroup",dn,"failure",groupDN & ", " & Err.Number & ", " & Err.Description
	End If

	Set oGroup = Nothing
End Function

function getRoot(dn)
	x = instr(dn,"DC=")
	if x < 0 then
		getRoot = ""
		exit function
	else
		getRoot = mid(dn,x)
	End if

End Function

Function GlobalGroupDN(rootDN, group)
	On Error Resume Next
	Const ADS_SCOPE_SUBTREE = 2
	
	GlobalGroupDN = ""

	Set objConnection = CreateObject("ADODB.Connection")
	Set objCommand =   CreateObject("ADODB.Command")
	objConnection.Provider = "ADsDSOObject"
	objConnection.Open "Active Directory Provider"
	Set objCommand.ActiveConnection = objConnection
	
	objCommand.Properties("Page Size") = 1000
	objCommand.Properties("Chase Referrals") = True
	objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
	objCommand.CommandText = "Select distinguishedName from 'LDAP://" & rootDN & "' where objectClass='group' and name='" & group & "'"
	Set objRecordSet = objCommand.Execute

	If objRecordSet.EOF Then
		Set objConnection = Nothing
		Set objCommand =  Nothing
		Set objRecordSet = Nothing
		Exit Function
	End if 

	objRecordSet.MoveFirst
	Do Until objRecordSet.EOF
		GlobalGroupDN = objRecordSet.Fields("distinguishedName").Value
	    objRecordSet.MoveNext
	Loop

	Set objConnection = Nothing
	Set objCommand =  Nothing
	Set objRecordSet = Nothing
End Function

Function ad_findduplicatecomputerinforest()
	On Error Resume Next
	Const ADS_SCOPE_SUBTREE = 2
	
	Set objRootDSE = GetObject("LDAP://rootDSE")
	strRootDomain = "LDAP://" & objRootDSE.Get("rootDomainNamingContext")
	Set objRootDomain = GetObject(strRootDomain)
	RootName = objRootDomain.Name
	Set objRootDSE = Nothing
	Set objRootDomain = Nothing

	Set objConnection = CreateObject("ADODB.Connection")
	Set objCommand =   CreateObject("ADODB.Command")
	objConnection.Provider = "ADsDSOObject"
	objConnection.Open "Active Directory Provider"
	Set objCommand.ActiveConnection = objConnection
	
	objCommand.Properties("Page Size") = 1000
	objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 

	objCommand.CommandText = "Select cn, distinguishedName from 'GC://" & RootName & ",DC=com' where objectClass='computer' and name='" & ComputerName & "'"

	Set objRecordSet = objCommand.Execute

	if objRecordSet.RecordCount = 1 Then
		WriteLog "ad_findduplicatecomputerinforest","","success","Unique computer name"
	Elseif objRecordSet.RecordCount = 0 Then
		WriteLog "ad_findduplicatecomputerinforest","","failure","Computer not found"
	Else
		WriteLog "ad_findduplicatecomputerinforest","Duplicates detected","failure",objRecordSet.RecordCount
	End If

	Set objConnection = Nothing
	Set objCommand =  Nothing
	Set objRecordSet = Nothing
End Function

Function ad_computerpasswordage()
	On Error Resume Next
	dn = distinguishedName(ComputerName)
	if dn = "" Then
		WriteLog "ad_computerpasswordage","failed","failure","failed to get computer dn"
		Exit Function
	End If

	objDate = getPasswordAge (dn)
	age = DateDiff("d", objDate, Now)

	WriteLog "ad_computerpasswordage",objDate,"success",age
End Function

Function getPasswordAge(dn)
	On Error Resume Next
	
	Const ADS_SCOPE_SUBTREE = 2
	
	getPasswordAge = ""

	Set objConnection = CreateObject("ADODB.Connection")
	Set objCommand =   CreateObject("ADODB.Command")
	objConnection.Provider = "ADsDSOObject"
	objConnection.Open "Active Directory Provider"
	Set objCommand.ActiveConnection = objConnection
	
	objCommand.Properties("Page Size") = 1000
	objCommand.Properties("Chase Referrals") = True
	objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE

	objCommand.Properties("Sort on") = "sAMAccountName"
	objCommand.CommandText = "<LDAP://" & dn & ">;(objectClass=computer);sAMAccountName,pwdLastSet,name,distinguishedname;subtree"

'	objCommand.CommandText = "Select distinguishedName from 'LDAP://" & root & "' where objectClass='group' and name='" & group & "'"
	Set objRecordSet = objCommand.Execute

	If objRecordSet.EOF Then
		Set objConnection = Nothing
		Set objCommand =  Nothing
		Set objRecordSet = Nothing
		Exit Function
	End if 

	objRecordSet.MoveFirst
	Do Until objRecordSet.EOF
		If not isnull(objRecordSet.Fields("distinguishedname")) and objRecordSet.Fields("distinguishedname") <> "" then
			objDate = objRecordSet.Fields("PwdLastSet")
			dtmPwdLastSet = Integer8Date(objDate, lngBias)
			'Field1 = dtmPwdLastSet
			'calculate the current age of the password.
			getPasswordAge = dtmPwdLastSet
			Field2 = DateDiff("d", getPasswordAge, Now)
'			age = DateDiff("d", dtmPwdLastSet, Now)

			'Go to function to make sense of the PwdLastSet value from AD for the machine account.
		End If
		objRecordSet.MoveNext
	Loop

	Set objConnection = Nothing
	Set objCommand =  Nothing
	Set objRecordSet = Nothing
	
End Function

Function Integer8Date(objDate, lngBias)
	' Function to convert Integer8 (64-bit) value to a date, adjusted for
	' local time zone bias.
	Dim lngAdjust, lngDate, lngHigh, lngLow
	lngAdjust = lngBias
	lngHigh = objDate.HighPart
	lngLow = objDate.LowPart
	' Account for bug in IADslargeInteger property methods.
	If lngLow < 0 Then
	lngHigh = lngHigh + 1
	End If
	If (lngHigh = 0) And (lngLow = 0) Then
	lngAdjust = 0
	End If
	lngDate = #1/1/1601# + (((lngHigh * (2 ^ 32)) _
	+ lngLow) / 600000000 - lngAdjust) / 1440
	Integer8Date = CDate(lngDate)
End Function 

Function lookupgroupmembership(node)
	On Error Resume Next

	dn = distinguishedName(ComputerName)
	If dn = "" Then
		WriteLog "lookupgroupmembership","","failure","Can not find computer distinguished name"
		Exit Function
	End If

	Dim aMemberOf
	ReDim aMemberOf(-1)
	
	Set objComputer = GetObject ("LDAP://" & dn)
	objComputer.GetInfo
	aMemberOf = objComputer.GetEx("memberOf")
	If UBound(aMemberOf) < 0 Then
		WriteLog "lookupgroupmembership","","failure","Computer has no group membership"
		Exit Function
	End IF
	
	For Each Group In aMemberOf
		FirstComma = Instr(Group,",")
		GroupName = Mid(Group,4,FirstComma-4)
		
		if ProcessVerificationTasks("groups/group[(@name='" & lCase(GroupName) & "')]") = true then
			WriteLog "lookupgroupmembership",lcase(GroupName),"success",""
		else
			if left(lcase(GroupName),22) = "agg-sms-wks-branchtype" then
				WriteLog "lookupgroupmembership",lcase(GroupName),"success","Nothing targeted at this container group"
			elseif left(lcase(GroupName),7) = "agg-sms" then
				WriteLog "lookupgroupmembership",lcase(GroupName),"failure","No match found for group"
			else
				WriteLog "lookupgroupmembership",lcase(GroupName),"success",""
			end if 
		end if 
		nestedGroupMembership(Group)
	Next
	Set objComputer = Nothing
End Function

Sub ProcessGroupMembership (GroupName)
	On Error Resume Next
	WriteLog "lookupgroupmembership",lcase(GroupName),"success",""
	ProcessVerificationTasks("groups/group[(@name='" & lCase(GroupName) & "')]")
End Sub

Function nestedGroupMembership(GroupDN)
	On Error Resume Next
	Dim aMemberOf
	ReDim aMemberOf(-1)
	
	Set objGroup = GetObject ("LDAP://" & GroupDN)
	objGroup.GetInfo
	aMemberOf = objGroup.GetEx("memberOf")
	If UBound(aMemberOf) < 0 Then
'		WriteLog "nestedgroupmembership",GroupDN,"failure","Group has no group membership"
		Exit Function
	End IF
	
	For Each Group In aMemberOf
		FirstComma = Instr(Group,",")
		GroupName = Mid(Group,4,FirstComma-4)
		
		if ProcessVerificationTasks("groups/group[(@name='" & lCase(GroupName) & "')]") = true then
			WriteLog "nestedgroupmembership",lcase(GroupName),"success",""
		else
			if left(lcase(GroupName),7) = "agg-sms" then
				WriteLog "nestedgroupmembership",lcase(GroupName),"failure","No match found for nested group"
			else
				WriteLog "nestedgroupmembership",lcase(GroupName),"success",""
			end if 
		end if 
		nestedGroupMembership(Group)
	Next
	Set objComputer = Nothing
End Function

Function SMSAdvertisementStatusMessages(Node)
	On Error Resume Next

	ResourceID = ""

'	Query = Replace(Node.GetAttributeNode("query").Value, "%COMPUTERNAME%", ComputerName)

	Set oServices = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & Node.GetAttributeNode("host").Value & "\root\sms\site_" & Node.GetAttributeNode("sitename").Value)
	If Err.Number <> 0 Then
		WriteLog "smswmiquery","Failed to connect to host " & Node.GetAttributeNode("host").Value,"failure",Err.number & ", " & Err.description
'		Field1 = "0"
		Exit Function
	End If

	Set oSystemSet = oServices.ExecQuery("Select ResourceID from SMS_R_System where NetbiosName='" & ComputerName & "'")
	For Each oSystem in oSystemSet
			ResourceID = oSystem.ResourceID
	Next
	Set oSystemSet = Nothing

	If ResourceID = "" Then
		WriteLog "smsadvertisementstatusmessages","No ResourceID","failure",""
		Exit Function
	End If


' SMS_Advertisement.AdvertisementName, SMS_Advertisement.AdvertisementID, SMS_ClientAdvertisementStatus.LastStatusTime, SMS_ClientAdvertisementStatus.LastStatusMessageIDName 
	Query = "select SMS_ClientAdvertisementStatus.LastStatusMessageIDName, SMS_ClientAdvertisementStatus.LastStatusTime from SMS_ClientAdvertisementStatus JOIN SMS_Advertisement on SMS_ClientAdvertisementStatus.AdvertisementID = SMS_Advertisement.AdvertisementID where SMS_ClientAdvertisementStatus.ResourceID = " & ResourceID
'Query = "select LastStatusMessageIDName from SMS_ClientAdvertisementStatus where SMS_ClientAdvertisementStatus.ResourceID = " & ResourceID
	Set oResults = oServices.ExecQuery(Query)
	For Each oResult in oResults
		WriteLog "smsadvertisementstatusmessages","","success",oResult.LastStatusMessageIDName ' & " - " & oResult.LastStatusTime
	Next

	Set oResult = Nothing
	Set oResults = Nothing
	Set oServices = Nothing	
	
End Function

Sub LoggedOnUser(Node)
	On Error Resume Next
	Dim Item, Exists
	
	Set oComputerSystem = oSWbemServices.ExecQuery("Select UserName from Win32_ComputerSystem")
	For Each Item In oComputerSystem
		UserName = Item.UserName
	Next
	Set oComputerSystem = Nothing
	
	If Len(UserName) > 0 Then
		WriteLog "loggedonuser","","success",UserName
	Else
		WriteLog "loggedonuser","","failure",""
	End If
	Field1 = Username
End Sub

Sub ResourceDomain(Node)
	On Error Resume Next
	Dim Item, Exists
	
	Set oComputerSystem = oSWbemServices.ExecQuery("Select Domain from Win32_ComputerSystem")
	For Each Item In oComputerSystem
		Domain = Item.Domain
	Next
	Set oComputerSystem = Nothing
	
	If Len(UserName) > 0 Then
		WriteLog "resourcedomain","","success",Domain
	Else
		WriteLog "resourcedomain","","failure",""
	End If
	Field1 = Domain
End Sub


Sub sms_mpcert(Node)
	On Error Resume Next

	url="http://" & ComputerName & "/sms_mp/.sms_aut?mpcert"
	Set objHTTP = CreateObject("MSXML2.XMLHTTP")

	Err.Clear
	Call objHTTP.Open("GET", url, FALSE)
	If Err.Number <> 0 Then
		WriteLog "sms_mpcert",url,"failure",Err.Number & ", " & Err.Description
		Exit Sub
	End If

	Err.Clear
	objHTTP.Send
	If Err.Number <> 0 Then
		WriteLog "sms_mpcert",url,"failure",Err.Number & ", " & Err.Description
		Exit Sub
	End If

	WriteLog "sms_mpcert",url,"success",objHTTP.Status
	Field1 = objHTTP.Status
	Field2 = objHTTP.StatusText	

	'WScript.Echo(objHTTP.Status)
	'WScript.Echo(objHTTP.StatusText)
	'WScript.Echo( "'" & objHTTP.ResponseText & "'")
End Sub

Sub sms_mplist(Node)
	On Error Resume Next

	url="http://" & ComputerName & "/sms_mp/.sms_aut?mplist"
	Set objHTTP = CreateObject("MSXML2.XMLHTTP")

	Err.Clear
	Call objHTTP.Open("GET", url, FALSE)
	If Err.Number <> 0 Then
		WriteLog "sms_mplist",url,"failure",Err.Number & ", " & Err.Description
		Exit Sub
	End If

	Err.Clear
	objHTTP.Send
	If Err.Number <> 0 Then
		WriteLog "sms_mplist",url,"failure",Err.Number & ", " & Err.Description
		Exit Sub
	End If

	WriteLog "sms_mplist",url,"success",objHTTP.Status
	Field1 = objHTTP.Status
	Field2 = objHTTP.StatusText	

	'WScript.Echo(objHTTP.Status)
	'WScript.Echo(objHTTP.StatusText)
	'WScript.Echo( "'" & objHTTP.ResponseText & "'")
End Sub

Sub RemoteFolderExists(Node)
	On Error Resume Next

	Name = Replace(Node.GetAttributeNode("name").Value, "%COMPUTERNAME%", ComputerName)
	If oFSO.FolderExists(Name) = False Then
		WriteLog "remotefolderexists",Name,"failure","Folder does not exist"
	Else
		WriteLog "remotefolderexists",Name,"success","folder exists"
	End If

End Sub


Sub RemoteFileExists(Node)
	On Error Resume Next

	Name = Replace(Node.GetAttributeNode("name").Value, "%COMPUTERNAME%", ComputerName)
	If oFSO.FileExists(Name) = False Then
		WriteLog "remotefileexists",Name,"failure","File does not exist"
	Else
		WriteLog "remotefileexists",Name,"success","file exists"
	End If

End Sub


Sub DesktopIconExists(Node)
	On Error Resume Next
	Dim Item, Exists
	
	Set oComputerSystem = oSWbemServices.ExecQuery("Select UserName from Win32_ComputerSystem")
	For Each Item In oComputerSystem
		UserName = Item.UserName
	Next
	Set oComputerSystem = Nothing
	
	If Len(UserName) > 0 Then
		' WriteLog "loggedonuser","","success",UserName
	Else
		WriteLog "desktopiconexists","","failure","No user logged on"
		Exit Sub
	End If
	Field1 = Username

	User = split(Username,"\")
	Username = user(1)

	If oFSO.FileExists("\\" & ComputerName & "\d$\Documents and Settings\" & UserName & "\Desktop\" & node.getAttributeNode("filename").Value) = True then
		WriteLog "desktopiconexists","","success","File exists: " & node.getAttributeNode("filename").Value
	Else
		WriteLog "desktopiconexists","","failure","File does not exist: " & node.getAttributeNode("filename").Value
	End If

End Sub


Sub ad_addglobalgrouptolocalgroupwithauth(Node)
	On Error Resume Next
	Err.Clear
	oNetwork.MapNetworkDrive "","\\" & ComputerName & "\ipc$",False,"", ""
	If Err.Number <> 0 Then
		msgbox err.number & vbcrlf & err.description
	End If
	
	Set objGroup = GetObject("WinNT://" & ComputerName & "/Administrators") 
	Set objUser = GetObject("WinNT://" & Node.getAttributeNode("global").value) 
	objGroup.Add(objUser.ADsPath) 

'	oNetwork.RemoveNetworkDrive "K:",True,True
End Sub


Function ADSite(node)
	On Error Resume Next
	Set colAdapters = oSWbemServices.ExecQuery("Select * from win32_networkadapterconfiguration where IPEnabled = True")
	Dim SubnetMask, strDescription
	
	For Each Adapter In colAdapters
		For x = LBound(Adapter.IPAddress) To UBound(Adapter.IPAddress)
			If IPAddress = Adapter.IPAddress(x) Then
				SubnetMask = SubnetIt(Adapter.IPAddress(x),Adapter.IPSubnet(x))
				Subnet = Adapter.IPSubnet(x)
				Set colAdapters = Nothing
			End If
		Next
	Next

	If Len(SubnetMask) < 1 Then
		WriteLog "adsite","","failure","Failed to retrieve subnet id"
		Exit Function
	End IF

	dim sDomain
	Set oComputerSystem = oSWbemServices.ExecQuery("Select Domain from Win32_ComputerSystem")
	For Each Item In oComputerSystem
		sDomain = ucase(Item.Domain)
	Next
	Set oComputerSystem = Nothing
	
	WriteLog "adsite","domain","success",sDomain

	Select Case sDomain
		case "AU.CBAINET.COM"
			strConfigurationNC = "CN=Configuration,DC=cbainet,DC=com"
		case "PBS.CBAINET.COM"
			strConfigurationNC = "CN=Configuration,DC=cbainet,DC=com"
		case "BRANCH1.CBAINET.COM"
			strConfigurationNC = "CN=Configuration,DC=cbainet,DC=com"
		case "BRANCH2.CBAINET.COM"
			strConfigurationNC = "CN=Configuration,DC=cbainet,DC=com"
		case "AUT01.CBAITEST01.COM"
			strConfigurationNC = "CN=Configuration,DC=cbaitest01,DC=com"
		case "PBST01.CBAITEST01.COM"
			strConfigurationNC = "CN=Configuration,DC=cbaitest01,DC=com"
		case "BRANCH1T01.CBAITEST01.COM"
			strConfigurationNC = "CN=Configuration,DC=cbaitest01,DC=com"
		case "AUD01.CBAIDEV01.COM"
			strConfigurationNC = "CN=Configuration,DC=cbaidev01,DC=com"
		case else
			strConfigurationNC = "CN=Configuration,DC=cbainet,DC=com"
	End Select 

'	Set objRootDSE = GetObject("LDAP://RootDSE")
'	strConfigurationNC = objRootDSE.Get("configurationNamingContext")
  Set objConn = CreateObject("ADODB.Connection")
  objConn.Provider = "ADsDSOObject"

	objCommand.Properties("Chase Referrals") = True
'	objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 

	objConn.Properties("User ID") = "aut01\vz56gr"
	objConn.Properties("Password") = "Winter33"
	objConn.Properties("Encrypt Password") = TRUE
	objConn.Properties("ADSI Flag") = 3
 
  objConn.Open "Active Directory Provider"

'  Set objRS = objConn.Execute(strBase & strFilter & strAttrs & strScope)
	 
	SubnetBits = calculateBits(Subnet)

  strMatchedSubnetDN = SubnetMask & "\/" & SubnetBits
  strBase = "<LDAP://cn=subnets,cn=sites," & strConfigurationNC & ">;"
  strFilter = "(&(objectcategory=subnet)" & "(distinguishedName=CN=" & strMatchedSubnetDN & ",cn=subnets,cn=sites," & strConfigurationNC & "));" 
  strAttrs = "name,location,siteObject,Description;"
  strScope = "subtree"


  strMatchedSubnetDN = SubnetMask & "\/" & SubnetBits
  strBase = "<LDAP://IAUNSWT09/cn=subnets,cn=sites," & strConfigurationNC & ">;"
  strFilter = "(&(objectcategory=subnet)" & "(distinguishedName=CN=" & strMatchedSubnetDN & ",cn=subnets,cn=sites," & strConfigurationNC & "));" 
  strAttrs = "name,location,siteObject,Description;"
  strScope = "subtree"

 	Err.Clear
  Set objRS = objConn.Execute(strBase & strFilter & strAttrs & strScope)
	If Err.Number <> 0 Then
		WriteLog "adsite",strSubnetsContainer,"failure","Failed to connect to AD, err.number=" & err.number & ", err.description=" & err.description
		Exit Function
	End If

strSite = ""
strName = ""
strLocation = ""
strDescription = ""


  If objRS.RecordCount > 0 then
    objRS.MoveFirst
    While Not objRS.EOF
      strName = objRS.Fields(0).Value
	  strLocation = objRS.Fields(1).Value
	  strSite = Split(Split(objRS.Fields(2), ",")(0), "=")(1)
	  strDescription = objRS.Fields(3).Value
'	  msgbox strDescription
      objRS.MoveNext
    Wend
  End If

	If strName = "" Then
		WriteLog "adsite",strSubnetsContainer,"failure","Failed to retrieve location"
		WriteLog "adsite","bits","success",SubnetBits
		WriteLog "adsite","subnet","success",Subnet
		WriteLog "adsite","subnetmask","success",SubnetMask
		Field1 = Subnet
		Field2 = SubnetMask
		Field3 = SubnetBits
	else
		WriteLog "adsite","name","success",strName
		WriteLog "adsite","location","success",strLocation
'		msgbox "strDescription: " & strDescription
		WriteLog "adsite","description","success",strDescription
		WriteLog "adsite","site","success",strSite
		WriteLog "adsite","bits","success",SubnetBits
		Field1 = strSite
		Field2 = strLocation
		Field3 = SubnetBits
		
	End If
	objConn.Close
	Set objConn = Nothing

'objSubnetsContainer.Filter = Array("subnet")

'	For Each objSubnet In objSubnetsContainer
'		objSubnet.GetInfoEx Array("siteObject"), 0	
'		strSiteObjectDN = objSubnet.Get("siteObject")
'		strSiteObjectName = Split(Split(strSiteObjectDN, ",")(0), "=")(1)
'		ar = split(objsubnet.Name,"/")
'
'		Name = ""
'		Bits = ""
'		Site = ""
'		Location = ""
'		Description = ""
'
'		Name = Mid(ar(0),4)
'		Bits = ar(1)
'		Site = Replace(strSiteObjectName, "'", "''")
'		Location = Replace(objSubnet.Location, "'", "''")
'		Description = Replace(objSubnet.Description, "'", "''")
'
'		oDBConnection.Execute "INSERT INTO " & TableName & " VALUES ('" & Name & "'," & Bits & ",'" & Site & "'," & "'" & Location & "'," & "'" & Description & "')"
'	Next

End Function

Function calculateBits(mask)
	' On Error Resume Next
	Dim bits : bits = 0

	arr = split(mask,".")
	for x = lBound(arr) to uBound(arr)
		select case arr(x)
			case "128"
				bits = bits + 1
			case "192"
				bits = bits + 2
			case "224"
				bits = bits + 3
			case "240"
				bits = bits + 4
			case "248"
				bits = bits + 5
			case "252"
				bits = bits + 6
			case "254"
				bits = bits + 7
			case "255"
				bits = bits + 8
			case else
				bits = bits + 0
		End Select
	Next

	calculateBits = bits

End Function

Function ad_location(Node)
	On Error Resume Next

	Const ADS_SCOPE_SUBTREE = 2

	dn = distinguishedName(ComputerName)
	if dn = "" Then
		WriteLog "ad_location","failed","failure","failed to get computer dn"
		Exit Function
	End If

	Set objConnection = CreateObject("ADODB.Connection")
	Set objCommand =   CreateObject("ADODB.Command")
	objConnection.Provider = "ADsDSOObject"
	objConnection.Open "Active Directory Provider"
	Set objCommand.ActiveConnection = objConnection
	
	objCommand.Properties("Page Size") = 1000
	objCommand.Properties("Chase Referrals") = True
	objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE

	objCommand.Properties("Sort on") = "sAMAccountName"
	objCommand.CommandText = "<LDAP://" & dn & ">;(objectClass=computer);sAMAccountName,location,name,distinguishedname;subtree"

'	objCommand.CommandText = "Select distinguishedName from 'LDAP://" & root & "' where objectClass='group' and name='" & group & "'"
	Set objRecordSet = objCommand.Execute

	If objRecordSet.EOF Then
		WriteLog "ad_location","failed","failure","object not found in forest"
		Set objConnection = Nothing
		Set objCommand =  Nothing
		Set objRecordSet = Nothing
		Exit Function
	End if 

	objRecordSet.MoveFirst
	Do Until objRecordSet.EOF
		location = objRecordSet.Fields("location")
		objRecordSet.MoveNext
	Loop

	If Len(location) > 0 Then
		WriteLog "ad_location","","success",Location
		Field2 = Location
	Else
		WriteLog "ad_location","","failure",""
		Field2 = Location
	End If

	Set objConnection = Nothing
	Set objCommand =  Nothing
	Set objRecordSet = Nothing
	
End Function


Sub RegistryKeyExistsError(Node)
	On Error Resume Next
	Dim RetVal, KeySplit, KeyBase, KeyName
	
	Name = Node.GetAttributeNode("name").Value
	KeyType = Node.GetAttributeNode("keytype").Value
	KeyValue = Node.GetAttributeNode("keyvalue").Value
	ErrorIfKeyDoesNotExist = Node.GetAttributeNode("errorifkeydoesnotexist").value

	KeySplit = Split(Name,"\")
	KeyBase = KeySplit(0)
	KeyName = KeySplit(UBound(KeySplit))
	SubKeyName = Mid(Name,Len(KeyBase)+2,len(Name) - (Len(KeyBase) + Len(KeyName) + 2))
	
	' Constants
	Const HKEY_CLASSES_ROOT = &H80000000
	Const HKEY_CURRENT_USER = &H80000001 
	Const HKEY_LOCAL_MACHINE = &H80000002
	Const HKEY_USERS = &H80000003
	Const HKEY_CURRENT_CONFIG = &H80000005
	Const HKEY_DYN_DATA = &H80000006

	Select Case KeyBase
		Case "HKEY_CLASSES_ROOT"
			sKeyBase = HKEY_CLASSES_ROOT
		Case "HKEY_CURRENT_USER"
			sKeyBase = HKEY_CURRENT_USER
		Case "HKEY_LOCAL_MACHINE"
			sKeyBase = HKEY_LOCAL_MACHINE
		Case "HKEY_USERS"
			sKeyBase = HKEY_USERS
		Case "HKEY_CURRENT_CONFIG"
			sKeyBase = HKEY_CURRENT_CONFIG
		Case "HKEY_DYN_DATA"
			sKeyBase = HKEY_DYN_DATA
	End Select

	Select Case KeyType
		Case "REG_SZ"
			retVal = oRegistry.GetStringValue (sKeyBase, SubKeyName,KeyName, RegVal)
		Case "REG_DWORD"
			retVal = oRegistry.GetDWordValue(sKeyBase,SubKeyName,KeyName, RegVal)
	End Select
	
	If retVal <> 0 Then
		If uCase(ErrorIfKeyDoesNotExist) = "TRUE" Then
			WriteLog "registrykeyexistserror",KeyBase & "\" & SubKeyName & "\" & KeyName,"failure","Registry key does not exist."
			Exit Sub
		Else
			WriteLog "registrykeyexistserror",KeyBase & "\" & SubKeyName & "\" & KeyName,"success","Registry key does not exist."
			Exit Sub
		End If
	End If

	' Error if the key exists
	WriteLog "registrykeyexistserror",KeyBase & "\" & SubKeyName & "\" & KeyName,"failure",RegVal

End Sub


Sub sms_ccm_ctm_jobstateex_error()
	On Error Resume Next

	Err.Clear 
	Set oSWbemPolicy = GetWMINamespace(ComputerName,"root\CCM\ContentTransferManager")

'	Set oSWbemPolicy = oWbemLocator.ConnectServer (ComputerName, "root\CCM\ContentTransferManager")
	If Err.Number <> 0 Then
 		WriteLog "sms_ccm_ctm_jobstateex_error","Could not connect to root\CCM\ContentTransferManager","failure",""
 		Exit Sub
	End If

	Err.Clear
	Set oClientActions = oSWbemPolicy.ExecQuery("Select ID, ContentID, SourceURL, kBytesTransferred from CCM_CTM_JobStateEx")
	If Err.Number <> 0 Then
 		WriteLog "sms_ccm_ctm_jobstateex_error","Could not query CCM_CTM_JobStateEx","failure",""
 		Exit Sub
	End If
	
	bError = False
	For Each Item In oClientActions
		bError = True
		WriteLog "sms_ccm_ctm_jobstateex_error",Item.ContentID,"failure","SourceURL=" & Item.SourceURL & ", BytesTransferred=" & Item.kBytesTransferred
		field1 = Item.kBytesTransferred
	Next

	If bError = False Then
		WriteLog "sms_ccm_ctm_jobstateex_error","No outstanding ctm jobs","success",""
	End If

	Set oClientActions = Nothing

	Set oSWbemPolicy = Nothing
	Set oClientActions = Nothing

End Sub

Sub peerdp_error()
	On Error Resume Next

	Err.Clear 
	Set oSWbemPolicy = GetWMINamespace(ComputerName,"root\CCM\PeerDPAgent")
'	Set oSWbemPolicy = oWbemLocator.ConnectServer (ComputerName, "root\CCM\ContentTransferManager")
	If Err.Number <> 0 Then
 		WriteLog "peerdp_error","Could not connect to root\CCM\PeerDPAgent","failure",Err.Description & "(" & Hex(Err.number) & ")"
 		Exit Sub
	End If

	Err.Clear
	Set oClientActions = oSWbemPolicy.ExecQuery("Select PackageID, State from CCM_PeerDP_Job where State != 'Succeeded'")
	If Err.Number <> 0 Then
 		WriteLog "peerdp_error","Could not query CCM_PeerDP_Job","failure",""
 		Exit Sub
	End If
	bError = False
	For Each Item In oClientActions
		bError = True
		WriteLog "peerdp_error",Item.PackageID,"failure",Item.State
	Next

	If bError = False Then
		WriteLog "peerdp_error","No outstanding Peer DP Transfers","success",""
	End If

	Set oClientActions = Nothing

	Set oSWbemPolicy = Nothing
	Set oClientActions = Nothing

End Sub

Sub peerdp_status(node)
	On Error Resume Next

	Err.Clear 
	Set oSWbemPolicy = GetWMINamespace(ComputerName,"root\CCM\PeerDPAgent")
'	Set oSWbemPolicy = oWbemLocator.ConnectServer (ComputerName, "root\CCM\ContentTransferManager")
	If Err.Number <> 0 Then
 		WriteLog "peerdp_status","Could not connect to root\CCM\PeerDPAgent","failure",Err.Description & "(" & Hex(Err.number) & ")"
 		Exit Sub
	End If

	minpackagesassigned = cint(Node.GetAttributeNode("minpackagesassigned").Value)

	Err.Clear
	Set oClientActions = oSWbemPolicy.ExecQuery("Select PackageID, State from CCM_PeerDP_Job")
	TotalCount = cint(oClientActions.Count)
	If Err.Number <> 0 Then
 		WriteLog "peerdp_status","Could not query CCM_PeerDP_Job","failure",""
 		Exit Sub
	End If
	If TotalCount > minpackagesassigned Then
		WriteLog "peerdp_status","Total packages assigned to BDP","success",TotalCount
	Else
		WriteLog "peerdp_status","Total packages assigned to BDP < " & minpackagesassigned,"failure","Actual Assigned: " & TotalCount
	End If 
	Set oClientActions = Nothing 

	Set oClientActions = oSWbemPolicy.ExecQuery("Select PackageID, State from CCM_PeerDP_Job where State = 'Succeeded'")
	SuccessfulCount = oClientActions.Count
	If Err.Number <> 0 Then
 		WriteLog "peerdp_status","Could not query CCM_PeerDP_Job","failure",""
 		Exit Sub
	End If
	WriteLog "peerdp_status","Total packages downloaded to BDP","success",SuccessfulCount
	
	TotalProblems = TotalCount - SuccessfulCount
	If totalProblems > 0 Then
		WriteLog "peerdp_status","Total non-downloaded packages to BDP","failure",TotalProblems
	Else
		WriteLog "peerdp_status","Total packages non-downloaded packages to BDP","success",TotalProblems
	End If 

	If bError = False Then
		WriteLog "peerdp_status","No outstanding Peer DP Transfers","success",""
	End If

	Set oClientActions = Nothing

	Set oSWbemPolicy = Nothing
	Set oClientActions = Nothing

End Sub




Sub sms_ccm_executionrequest_error()
	On Error Resume Next

	Err.Clear 
	Set oSWbemPolicy = GetWMINamespace(ComputerName,"root\CCM\SoftMgmtAgent")
'	Set oSWbemPolicy = oWbemLocator.ConnectServer (ComputerName, "root\CCM\SoftMgmtAgent")
	If Err.Number <> 0 Then
 		WriteLog "sms_ccm_executionrequest_error","Could not connect to root\CCM\SoftMgmtAgent","failure",""
 		Exit Sub
	End If

	Err.Clear
	Set oClientActions = oSWbemPolicy.ExecQuery("Select AdvertID, ContentID, ProgramID, State from CCM_ExecutionRequestEx")
	If Err.Number <> 0 Then
 		WriteLog "sms_ccm_executionrequest_error","Could not query CCM_ExecutionRequest","failure",""
 		Exit Sub
	End If
	
	bError = False
	For Each Item In oClientActions
		If Item.State = "WaitingDisabled" Then
			' The program is disabled, don't care!!!
		Else
			bError = True
			WriteLog "sms_ccm_executionrequest_error","ContentID=" & Item.ContentID & ", ProgramID=" & Item.ProgramID,"failure","AdvertID=" & Item.AdvertID & ", State=" & Item.State
		End If 
	Next

	If bError = False Then
		WriteLog "sms_ccm_executionrequest_error","No outstanding execution requests","success",""
	End If

	Set oClientActions = Nothing

	Set oSWbemPolicy = Nothing
	Set oClientActions = Nothing

End Sub

Sub sms_mpproxyinformation()
	On Error Resume Next

	Err.Clear 
	Set oSWbemPolicy = GetWMINamespace(ComputerName,"root\CCM")
' 	Set oSWbemPolicy = oWbemLocator.ConnectServer (ComputerName, "root\CCM")
	If Err.Number <> 0 Then
 		WriteLog "sms_mpproxyinformation","Could not connect to root\CCM","failure",""
 		Exit Sub
	End If

	Err.Clear
	Set oClientActions = oSWbemPolicy.ExecQuery("Select name, sitecode, state from sms_mpproxyinformation")
	If Err.Number <> 0 Then
 		WriteLog "sms_mpproxyinformation","Could not query sms_mpproxyinformation","failure",""
 		Exit Sub
	End If
	
	bError = True
	For Each Item In oClientActions
		If Len(Item.Name) < 1 then
			WriteLog "sms_mpproxyinformation","Name=" & Item.Name & ", SiteCode=" & Item.SiteCode,"success","Proxy management point blank"
		else
			WriteLog "sms_mpproxyinformation","Name=" & Item.Name & ", SiteCode=" & Item.SiteCode,"success","State=" & Item.State
			Field1 = Item.Name
			Field2 = Item.SiteCode
		End If
		bError = False
	Next

	If bError = True Then
		WriteLog "sms_mpproxyinformation","No proxy information","failure",""
	End If

	Set oClientActions = Nothing

	Set oSWbemPolicy = Nothing
	Set oClientActions = Nothing

End Sub

Sub EventLogErrorCount(node)
	On Error Resume Next
	Dim Count, LogFile, Date24
	DateYesterday = DateAdd("h",-24,Now)
	Date24 = Year(DateYesterday) & PadTo2(Month(DateYesterday)) & PadTo2(Day(DateYesterday)) & PadTo2(Hour(DateYesterday)) & PadTo2(Minute(DateYesterday)) & PadTo2(Second(DateYesterday)) & ".000000+600"
	LogFile = node.GetAttributeNode("logfile").Value

	Set Win32_Event = oSWbemServices.ExecQuery("select TimeGenerated, Type, EventCode, Message from win32_ntlogevent where LogFile='" & LogFile & "' and Type='Error' And TimeGenerated > '" & Date24 & "'")
	count = 0
	count = Win32_Event.Count
	If Count > 0 Then
		WriteLog "eventlogerrorcount",LogFile & " logfile error count " & TimeGenerated,"failure",Count
	Else
		WriteLog "eventlogerrorcount",LogFile & " logfile error count","success",Count
	End If 
	Set Win32_Event = Nothing
End Sub 

Sub sms_policy(Node)
	On Error Resume Next

	Policy = node.GetAttributeNode("policyid").Value
	Set objHTTP = CreateObject("MSXML2.XMLHTTP")

	url = "http://" & ComputerName & "/sms_mp/.sms_pol?" & Policy 
	Call objHTTP.Open ("GET", url, FALSE)
	objHTTP.Send

	If objHTTP.StatusText <> "OK" Then
		WriteLog "sms_policy",url,"failure",objHTTP.StatusText & ", " & objHTTP.Status
	else
		WriteLog "sms_policy",url,"success",objHTTP.StatusText & ", " & objHTTP.Status
	End If

	Set objHTTP = Nothing
'	WScript.Echo(objHTTP.Status)
'	WScript.Echo(objHTTP.StatusText)
'	WScript.Echo( "'" & objHTTP.ResponseText & "'")
End Sub


Sub sms_lastadvertstatusmessage(node)
	On Error Resume Next
	Dim bFound : bFound = false
	Dim oconnection : Set oConnection = CreateObject("ADODB.Connection")
	Dim rs : Set rs = CreateObject("ADODB.RecordSet")
	Dim query

	Err.Clear
	oConnection.ConnectionString = "Provider='SQLOLEDB';Data Source='" & node.GetAttributeNode("dbserver").Value & "';Initial Catalog='" & node.GetAttributeNode("dbdatabase").Value & "' ;Integrated Security='SSPI';"
	oConnection.ConnectionTimeout = 10
	oConnection.CommandTimeout = 5
	oConnection.Open
	If Err.Number <> 0 Then
		WriteLog "sms_lastadvertstatusmessage","Error connecting to database","failure","Error " & Err.Description & " - " & Err.Number
		Exit Sub
	End If

	query = "SELECT stat.laststatustime, stat.LastAcceptanceStateName, stat.LastStateName, stat.LastStatusMessageIDName nolock " & _
		"FROM v_Advertisement adv " & _
		"JOIN v_ClientAdvertisementStatus stat ON stat.AdvertisementID = adv.AdvertisementID " & _
		"JOIN v_R_System sys ON stat.ResourceID=sys.ResourceID " & _
		"WHERE sys.Netbios_Name0 = '" & ComputerName & "' " & _
		"and adv.AdvertisementID = '" & node.GetAttributeNode("advertid").Value & "'"

	rs.Open query, oconnection, adOpenStatic, adLockOptimistic
	rs.MoveFirst 
	Do Until rs.EOF
		WriteLog "sms_lastadvertstatusmessage", node.GetAttributeNode("advertid").Value,"success",rs.Fields.Item(0).Value & vbTab & rs.Fields.Item(1).Value & vbTab & rs.Fields.Item(2).Value & vbtab & rs.Fields.Item(3).Value 
		field1 = rs.Fields.Item(1).Value & vbTab & rs.Fields.Item(2).Value & vbtab & rs.Fields.Item(3).Value 
		field2 = rs.Fields.Item(0).Value
		bFound = True
		rs.MoveNext 
	Loop
	
	If bFound = False Then
		WriteLog "sms_lastadvertstatusmessage",node.GetAttributeNode("advertid").Value,"failure","no status message found"
	End If

	rs.Close
	Set rs = Nothing
	oConnection.Close
	Set oConnection = Nothing
End Sub

Sub sms_cacheinfo(node)
	On Error Resume Next
	
	Err.Clear 
	Set oSWbemPolicy = oWbemLocator.ConnectServer (ComputerName, "root\CCM\SoftMgmtAgent")
	If Err.Number <> 0 Then
 		WriteLog "sms_cacheinfo","Could not connect to root\CCM\SoftMgmtAgent","failure",""
 		Exit Sub
	End If

	Err.Clear
	Set oClientActions = oSWbemPolicy.ExecQuery("Select contentid, contentver, contentsize, location from CacheInfo where contentid='" & node.getattributenode("contentid").value & "'")
	If Err.Number <> 0 Then
 		WriteLog "sms_cacheinfo","Could not query CacheInfo","failure",""
 		Exit Sub
	End If
	
	bError = True
	For Each Item In oClientActions
		If Item.contentver <> node.getattributenode("contentver").value Then
			WriteLog "sms_cacheinfo","contentid=" & Item.contentid,"failure","expecting contentver=" & node.getattributenode("contentver").value & ", actual contentver=" & Item.contentver
		Elseif cstr(item.contentsize) <> node.getattributenode("contentsize").value Then
			WriteLog "sms_cacheinfo","contentid=" & Item.contentid,"failure","expecting contentsize=" & node.getattributenode("contentsize").value & ", actual contentsize=" & Item.contentsize
		Else
			WriteLog "sms_cacheinfo","contentid=" & Item.contentid,"success","Location=" & item.Location
		End If
		bError = False
		Field1 = Item.contentver
		'Field2 = cstr(item.contentsize)
	Next

	If bError = True Then
		WriteLog "sms_cacheinfo","contentid=" & node.getattributenode("contentid").value,"failure","Package is not cached"
	End If

	Set oClientActions = Nothing

	Set oSWbemPolicy = Nothing
	Set oClientActions = Nothing
End Sub

Sub wmiobjectsdatahealth
	On Error Resume Next
	Dim Exists, FileDateModified
	
	if oFSO.FileExists("\\" & Computername & "\admin$\system32\wbem\repository\FS\objects.data") = false then
		WriteLog "wmiobjectsdatahealth","\\" & Computername & "\admin$\system32\wbem\repository\FS\objects.data","failure","Can not connect to file"
		field2 = "WMI Cactus Possibly"
		Exit sub
	end if

	err.clear 
	Set oFile = oFSO.GetFile("\\" & Computername & "\admin$\system32\wbem\repository\FS\objects.data")
	if err.number <> 0 then
		WriteLog "wmiobjectsdatahealth","\\" & Computername & "\admin$\system32\wbem\repository\FS\objects.data","failure","Can not connect to file"
		field2 = "WMI Cactus Possibly"
		Exit sub
	End If

	DateLastModified = oFile.DateLastModified
	daysSinceLastUpdate = datediff("d",DateLastModified,now())
	if daysSinceLastUpdate > 7 then
		WriteLog "wmiobjectsdatahealth","\\" & Computername & "\admin$\system32\wbem\repository\FS\objects.data file " & daysSinceLastUpdate & " days old","failure","WMI Cactus - Please Repair"
		field2 = "WMI Cactus"
	else
		WriteLog "wmiobjectsdatahealth","objects.data file last modified" & daysSinceLastUpdate & " days ago.","success",""
	End If

	Set oFile = Nothing
End Sub

Sub win32_ntlogevent(node)
	On Error Resume Next
	
	Set dtmStartDate = CreateObject("WbemScripting.SWbemDateTime")
	dtmStartDate.SetVarDate date() - cint(node.getattributenode("daysago").value), True

	Err.Clear 
	Set oSWbemPolicy = oWbemLocator.ConnectServer (ComputerName, "root\cimv2")
	If Err.Number <> 0 Then
 		WriteLog "win32_ntlogevent","Could not connect to root\cimv2","failure",""
 		Exit Sub
	End If

Dim x : x = 0

	Err.Clear
	Set oClientActions = oSWbemPolicy.ExecQuery("select TimeGenerated, Type, EventCode, Message from win32_ntlogevent where LogFile='System' and EventCode='" & node.getattributenode("eventcode").value & "' and TimeGenerated >= '" & dtmStartDate & "'")
	If Err.Number <> 0 Then
		
 		WriteLog "win32_ntlogevent","Could not query CacheInfo","failure",""
 		Exit Sub
	End If
	
	For Each Item In oClientActions
		if x = 0 then
			field1 = Item.TimeGenerated
		elseif x = 1 then
			field2 = Item.TimeGenerated
		End if
		WriteLog "win32_ntlogevent","EventCode=" & Item.eventcode,"success",Item.TimeGenerated
		x = x + 1
	Next

'	WriteLog "win32_ntlogevent","EventCode=" & Item.EventCode,"success",Item.TimeGenerated

	Set oClientActions = Nothing
	Set oSWbemPolicy = Nothing
	Set oClientActions = Nothing

End Sub

Function sms_dpfreespace(Node)
	On Error Resume Next
	Dim Win32_LogicalDisk, objDisk, colDisks

	' Calculate which drive the DP is on
	
	dpFolder = left(GetDPFolder(),2)
	field1 = dpFolder
	WriteLog "sms_dpfreespace","dpfolder","data",dpFolder
	
	' Work out how much free space on the DP drive.
	Set Win32_LogicalDisk = oSWbemServices.ExecQuery("select Name, FreeSpace, Size from Win32_LogicalDisk where Name ='" & dpFolder & "'")
	For each objDisk in Win32_LogicalDisk
		field2 = clng(round(objDisk.FreeSpace/1000000,0))
		WriteLog "sms_dpfreespace","dpfreespaceinmb","data",cstr(round(objDisk.FreeSpace/1000000,0))
		field3 = clng(round(objDisk.Size/1000000,0))
		WriteLog "sms_dpfreespace","dpsizemb","data",cstr(round(objDisk.Size/1000000,0))
		If clng(round(objDisk.FreeSpace/1000000,0)) <= clng(Node.GetAttributeNode("minimumsize").Value) Then
			WriteLog "sms_dpfreespace",objDisk.Name,"failure",cstr(round(objDisk.FreeSpace/1000000,0)) & " < " & Node.GetAttributeNode("minimumsize").Value
		Else
			WriteLog "sms_dpfreespace",objDisk.Name,"success",cstr(round(objDisk.FreeSpace/1000000,0))
		End If
		Set Win32_LogicalDisk = Nothing
		Exit Function
	Next
	WriteLog "sms_dpfreespace",Node.GetAttributeNode("driveletter").Value,"failure","Drive not found"
    Set Win32_LogicalDisk = Nothing
End Function

Function GetDPFolder()
	On Error Resume Next
	Dim oServices
	GetDPFolder = ""
	
	Set oServices = oSWbemServices.ExecQuery("Select Name, Path from win32_share where name = 'DP$'")
	For Each Item In oServices
		GetDPFolder = Item.Path
	Next

	Set oServices = Nothing
End Function

Function bluetooth(node)
	On Error Resume Next
		err.clear 
		Set oComputerSystem = oSWbemServices.ExecQuery("Select manufacturer, model from Win32_ComputerSystem")
		if err.number <> 0 then
			WriteLog "bluetooth","","failure","err.number=" & err.number & ", err.description=" & err.description
		else
			For Each Item In oComputerSystem
				WriteLog "bluetooth","Manufacturer","success",Item.Manufacturer
				WriteLog "bluetooth","Model","success",Item.Model
'				field1 = item.manufacturer
				field1 = item.model
			Next
		End if
		Set oComputerSystem = Nothing

	Dim objInstance, strPropValue
	err.clear 
	' Set objInstance = GetObject("WinMgmts:{impersonationLevel=impersonate}//" & ComputerName & "/root/Dellomci:Dell_SMBIOSSettings=0")
	Set objInstance = GetWMINamespace (ComputerName,"root/Dellomci:Dell_SMBIOSSettings=0")
	strPropValue = objInstance.Properties_.Item("BluetoothDevices").Value
	If Err.Number <> 0 Then
		WriteLog "bluetooth",Name,"failure",Err.number & "," & Err.Description
		Exit Function
	End If
	
	WriteLog "bluetooth",Name,"success","BlueToothDevices",strPropValue
	Field2 = strPropValue
	
End Function

Function vpro(node)
	On Error Resume Next
		err.clear 
		Set oComputerSystem = oSWbemServices.ExecQuery("Select manufacturer, model from Win32_ComputerSystem")
		if err.number <> 0 then
			WriteLog "vpro","","failure","err.number=" & err.number & ", err.description=" & err.description
		else
			For Each Item In oComputerSystem
				WriteLog "vpro","Manufacturer","success",Item.Manufacturer
				WriteLog "vpro","Model","success",Item.Model
'				field1 = item.manufacturer
				field1 = item.model
			Next
		End if
		Set oComputerSystem = Nothing

	Dim objInstance, strPropValue
	err.clear 
	' Set objInstance = GetObject("WinMgmts:{impersonationLevel=impersonate}//" & ComputerName & "/root/Dellomci:Dell_SMBIOSSettings=0")
	Set objInstance = GetWMINamespace (ComputerName,"root/Dellomci:Dell_IntelvProSettings=0")
	strPropValue = objInstance.Properties_.Item("MEFWMajorVersion").Value
	If Err.Number <> 0 Then
		WriteLog "vpro",Name,"failure",Err.number & "," & Err.Description
		Exit Function
	End If
	
	WriteLog "vpro",Name,"success","MEFWMajorVersion",strPropValue
	Field2 = strPropValue
	
End Function

Function serialasset(node)
		err.clear 
		Set oComputerSystem = oSWbemServices.ExecQuery("select serialnumber, smbiosassettag from Win32_SystemEnclosure")
		if err.number <> 0 then
			WriteLog "serialasset","","failure","err.number=" & err.number & ", err.description=" & err.description
		Else
			For Each Item In oComputerSystem
				WriteLog "serialasset","serialnumber","success",Item.serialnumber
				WriteLog "serialasset","smbiosassettag","success",Item.smbiosassettag
				field1 = item.serialnumber
				field2 = item.smbiosassettag
			Next
		End if
		Set oComputerSystem = Nothing
End Function

Function GetWMINamespace(strComputerName,namespace)
	On Error Resume Next
	' Logic for SST
	Dim UserName, Password
	
	' Try integrated authentication first
	If g_IntegratedAuthentication = True Then
		Err.Clear 
		Set oNameSpace = oWbemLocator.ConnectServer (strComputerName, namespace)
		If Err.Number = 0 Then
			Set GetWMINamespace = oNameSpace
			g_IntegratedAuthentication = True
			Exit Function
		Else
			g_IntegratedAuthentication = False
		End If
	End If

	If Len(g_LastSuccessfulUsername) > 0 Then
		' We have successfully connected before 
		Err.Clear
		Set oNameSpace = oWbemLocator.ConnectServer (strComputerName, namespace,g_LastSuccessfulUsername,g_LastSuccessfulPassword,,,,128)
		If Err.Number = 0 Then
			Set GetWMINamespace = oNameSpace
			g_IntegratedAuthentication = False
			Exit Function
		End If
	End If
	
	' Loop the recordset and attempt specific authentication
	For each Node in oDataXML.selectSingleNode("data/users").ChildNodes 
		Username = node.GetAttributeNode ("name").value
		Password = node.GetAttributeNode ("password").value
		if ucase(Username) = "TOPTIGER" then
			Username = ComputerName & "\toptiger"
		end if 
		If Len(Username) > 0 And Len(Password) > 0 Then
			Err.Clear
			Set oNameSpace = oWbemLocator.ConnectServer (strComputerName, namespace,Username,Password)
			If Err.Number = 0 Then
				Set GetWMINamespace = oNameSpace
				g_IntegratedAuthentication = False
				g_LastSuccessfulUsername = Username
				g_LastSuccessfulPassword = Password
				Exit Function
			End If
		End If
	Next

	g_IntegratedAuthentication = True
	Set GetWMINamespace = Nothing
	
End Function

Sub sms_cacheinfo1(node)
	On Error Resume Next
	
	Err.Clear 
	Set oSWbemPolicy = oWbemLocator.ConnectServer (ComputerName, "root\CCM\SoftMgmtAgent")
	If Err.Number <> 0 Then
 		WriteLog "sms_cacheinfo1","Could not connect to root\CCM\SoftMgmtAgent","failure",""
 		Exit Sub
	End If

	Err.Clear
	Set oClientActions = oSWbemPolicy.ExecQuery("Select contentid, contentver, contentsize, location from CacheInfo where contentid='" & node.getattributenode("contentid").value & "' AND contentver='" & node.getattributenode("contentver").value & "'")
	If Err.Number <> 0 Then
 		WriteLog "sms_cacheinfo1","Could not query CacheInfo","failure",""
 		Exit Sub
	End If
	
	bError = True
	For Each Item In oClientActions
		If Item.contentver <> node.getattributenode("contentver").value Then
			WriteLog "sms_cacheinfo1","contentid=" & Item.contentid,"failure","expecting contentver=" & node.getattributenode("contentver").value & ", actual contentver=" & Item.contentver
		Elseif cstr(item.contentsize) <> node.getattributenode("contentsize").value Then
			WriteLog "sms_cacheinfo1","contentid=" & Item.contentid,"failure","expecting contentsize=" & node.getattributenode("contentsize").value & ", actual contentsize=" & Item.contentsize
		Else
			WriteLog "sms_cacheinfo1","contentid=" & Item.contentid,"success","Location=" & item.Location
		End If
		bError = False
		Field1 = Item.contentver
		'Field2 = cstr(item.contentsize)
	Next

	If bError = True Then
		WriteLog "sms_cacheinfo1","contentid=" & node.getattributenode("contentid").value,"failure","Package is not cached"
	End If

	Set oClientActions = Nothing

	Set oSWbemPolicy = Nothing
	Set oClientActions = Nothing
End Sub

Function PadTo2(stuff)
	On Error Resume Next
	If Len(stuff) = 0 Then
		PadTo2 = "00"
	Elseif Len(stuff) = 1 Then
		PadTo2 = "0" & Stuff
	Else
		PadTo2 = stuff
	End If 

End Function 

Function GetSoftwareDistribution
	On Error Resume Next
	Dim tblText, Item, Win32_QFE, oRequestedConfig
	Dim State,RunStartTime
	
	strUserName
	strUserDomain

	If strUserName <> "" Then
		Set objRecordSet = QueryActiveDirectory("Select distinguishedName from '" & ConvertDomainToQuery(strUserDomain) & "' where objectClass='user' and samAccountname='" & strUserName & "'")
		objRecordSet.MoveFirst
		Set oUser = GetObject("LDAP://" & objRecordSet.Fields("distinguishedName").Value)
		objMemberOf = oUser.GetEx("MemberOf")
		arrsid = oUser.Get("ObjectSid")
		strSidHex = OctetToHexStr(arrSid)
		strSidDec = HexStrToDecStr(strSidHex)
		UserSid = strSidDec
		UserSidModified = Replace(UserSid,"-","_")
	End If

	GetStatusInformation(UserSid)

	Target = "Machine"

	Set oRequestedConfig = GetWMINamespace("root\CCM\Policy\Machine\RequestedConfig")

'	Set oRequestedConfig = oWbemLocator.ConnectServer (strComputerName, "root\CCM\Policy\Machine\RequestedConfig")
	' Query Win32_QuickFixEngineering Information
	tblText = "<table id=""tblSoftwareDistribution"" width='100%' border='0' cellspacing='1' cellpadding='2' OnClick=""sortColumn(window.event)"">" & vbcrlf & _
		"<THEAD>" & VbCRLF & _
		"<tr>" & vbcrlf & _
		"<td align='center' id='SoftwareDistributionColum1' class='TableHeading' type=""CaseInsensitiveString"" height='12'>Package</td>" & vbcrlf & _
		"<td align='center' class='TableHeading' type=""CaseInsensitiveString"" height='12'>Program</td>" & vbcrlf & _
		"<td align='center' class='TableHeading' type=""CaseInsensitiveString"" height='12'>Version</td>" & vbcrlf & _
		"<td align='center' class='TableHeading' type=""CaseInsensitiveString"" height='12'>Source Version</td>" & vbcrlf & _
		"<td align='center' class='TableHeading' type=""CaseInsensitiveString"" height='12'>Rerun Behaviour</td>" & vbcrlf & _
		"<td align='center' class='TableHeading' type=""CaseInsensitiveString"" height='12'>PolicySource</td>" & vbcrlf & _
		"<td align='center' class='TableHeading' type=""CaseInsensitiveString"" height='12'>PolicyID</td>" & vbcrlf & _
		"<td align='center' class='TableHeading' type=""CaseInsensitiveString"" height='12'>Target</td>" & vbcrlf & _
		"<td align='center' class='TableHeading' type=""CaseInsensitiveString"" height='12'>State</td>" & vbcrlf & _
		"<td align='center' class='TableHeading' type=""CaseInsensitiveString"" height='12'>Run Start Time</td>" & vbcrlf & _
		"</tr>" & _
		"</THEAD>" & _
		"<TBODY>"

	Set Win32_QFE = oRequestedConfig.ExecQuery("select PKG_Name, PRG_HistoryLocation, PKG_PackageID, PKG_SourceVersion, PRG_ProgramName, PKG_Version, PolicySource, PolicyID, ADV_RepeatRunBehavior from CCM_SoftwareDistribution")
	For Each Item In Win32_QFE
		State = ""
		RunStartTime = ""
		If Item.PRG_HistoryLocation = "Machine" Then
			GetStuffed Item.PKG_PackageID & ":" & Item.PRG_ProgramName,True,State,RunStartTime
		Else
			GetStuffed Item.PKG_PackageID & ":" & Item.PRG_ProgramName,False,State,RunStartTime
		End If
		
		tblText = tblText & _
			"<tr>" & vbcrlf & _
			"<td align='center' class='TableDetails' height='12'>" & Item.PKG_Name & "</td>" & vbcrlf & _
			"<td align='center' class='TableDetails' height='12'>" & Item.PRG_ProgramName & "</td>" & vbcrlf & _
			"<td align='center' class='TableDetails' height='12'>" & Item.PKG_Version & "</td>" & vbcrlf & _
			"<td align='center' class='TableDetails' height='12'>" & Item.PKG_SourceVersion & "</td>" & VbCrLf & _
			"<td align='center' class='TableDetails' height='12'>" & Item.ADV_RepeatRunBehavior & "</td>" & VbCrLf & _
			"<td align='center' class='TableDetails' height='12'>" & Item.PolicySource & "</td>" & vbcrlf & _
			"<td align='center' class='TableDetails' height='12'>" & Item.PolicyID & "</td>" & vbcrlf & _
			"<td align='center' class='TableDetails' height='12'>" & Target & "</td>" & vbcrlf & _
			"<td align='center' class='TableDetails' height='12'>" & State & "</td>" & vbcrlf & _
			"<td align='center' class='TableDetails' height='12'>" & RunStartTime & "</td>" & vbcrlf & _
			"</tr>"
	Next
	Set Win32_QFE = Nothing
	Set oRequestedConfig = Nothing

	' Get user information
	If strUserName <> "" Then
		Target = "User"
		NameSpace = "root\CCM\Policy\" & UserSidModified & "\RequestedConfig"
		Set oRequestedConfig = GetWMINamespace("root\CCM\Policy\" & UserSidModified & "\RequestedConfig")

		Set Win32_QFE = oRequestedConfig.ExecQuery("select PKG_Name, PRG_HistoryLocation, PKG_PackageID, PKG_SourceVersion, PRG_ProgramName, PKG_Version, PolicySource, PolicyID, ADV_RepeatRunBehavior from CCM_SoftwareDistribution")
		For Each Item In Win32_QFE
			State = ""
			RunStartTime = ""
			If Item.PRG_HistoryLocation = "Machine" Then
				GetStuffed Item.PKG_PackageID & ":" & Item.PRG_ProgramName,True,State,RunStartTime
			Else
				GetStuffed Item.PKG_PackageID & ":" & Item.PRG_ProgramName,False,State,RunStartTime
			End If
			tblText = tblText & _
				"<tr>" & vbcrlf & _
				"<td align='center' class='TableDetails' height='12'>" & Item.PKG_Name & "</td>" & vbcrlf & _
				"<td align='center' class='TableDetails' height='12'>" & Item.PRG_ProgramName & "</td>" & vbcrlf & _
				"<td align='center' class='TableDetails' height='12'>" & Item.PKG_Version & "</td>" & vbcrlf & _
				"<td align='center' class='TableDetails' height='12'>" & Item.PKG_SourceVersion & "</td>" & vbcrlf & _
				"<td align='center' class='TableDetails' height='12'>" & Item.ADV_RepeatRunBehavior & "</td>" & vbcrlf & _
				"<td align='center' class='TableDetails' height='12'>" & Item.PolicySource & "</td>" & vbcrlf & _
				"<td align='center' class='TableDetails' height='12'>" & Item.PolicyID & "</td>" & vbcrlf & _			
				"<td align='center' class='TableDetails' height='12'>" & Target & "</td>" & vbcrlf & _
				"<td align='center' class='TableDetails' height='12'>" & State & "</td>" & vbcrlf & _
				"<td align='center' class='TableDetails' height='12'>" & RunStartTime & "</td>" & vbcrlf & _
				"</tr>"
		Next
		Set Win32_QFE = Nothing
		Set oRequestedConfig = Nothing

	End If

    tblText = tblText & "</TBODY></table>" & vbcrlf
	document.getElementById("sysSoftwareDistribution").innerHTML = tblText
	document.getElementById("aSoftwareDistribution").innerHTML = "SMS Software Distribution"
	window.document.getElementById("SoftwareDistributionColum1").click()
End Function

Function sccm_policyreceived(node)
	On Error Resume Next
	Dim tblText, Item, Win32_QFE, oRequestedConfig
	Dim State,RunStartTime
	Dim bFound : bFound = False
	Dim PolicyID 
	Dim PKG_Name : PKG_Name = ""
	PolicyID = 	node.getattributenode("policyid").value

	err.clear 
	Set oRequestedConfig = GetWMINamespace(ComputerName,"root\ccm\Policy\Machine\RequestedConfig")
	If Err.Number <> 0 Then
 		WriteLog "sccm_policyreceived","Could not connect to root\CCM\Policy\Machine\RequestedConfig","failure","(" & err.number & ")" & err.description 
 		Exit Function
	End If
	
	err.clear 
	Set Win32_QFE = oRequestedConfig.ExecQuery("select * from CCM_SoftwareDistribution where policyid = '" & PolicyID & "'")
	If Err.Number <> 0 Then
 		WriteLog "sccm_policyreceived","Could not query CCM_SoftwareDistribution","failure",""
 		Exit function
	End If

	For Each Item In Win32_QFE
		bFound = True
		PKG_Name = Item.PKG_Name
		WriteLog "sccm_policyreceived",PolicyID,"success","Policy Received for Package: " & PKG_Name
	Next
	if bFound = False then
		WriteLog "sccm_policyreceived",PolicyID,"failure","Policy has not been received."
	end if 
	Set Win32_QFE = Nothing
	Set oRequestedConfig = Nothing

End Function

Sub sms_ccm_softwaredistribution_exists(node)
	On Error Resume Next

	advertisementid = node.GetAttributeNode("advertisementid").Value

	Err.Clear 
	Set oSWbemPolicy = GetWMINamespace(ComputerName,"root\CCM\Policy\Machine\ActualConfig")

'	Set oSWbemPolicy = oWbemLocator.ConnectServer (ComputerName, "root\CCM\Policy\Machine\ActualConfig")
	If Err.Number <> 0 Then
 		WriteLog "sms_ccm_softwaredistribution_exists","Could not connect to root\CCM\Policy\Machine\ActualConfig","failure",""
 		Exit Sub
	End If

	Err.Clear
	Set oClientActions = oSWbemPolicy.ExecQuery("Select ADV_AdvertisementID, Pkg_Name from CCM_SoftwareDistribution where ADV_AdvertisementID='" & advertisementid & "'")
	If Err.Number <> 0 Then
 		WriteLog "sms_ccm_softwaredistribution_exists","Could not query CCM_SoftwareDistribution","failure",""
 		Exit Sub
	End If
	
	bFound = False
	For Each Item In oClientActions
		bFound = True
		WriteLog "sms_ccm_softwaredistribution_exists",Item.ADV_AdvertisementID,"success",Item.Pkg_Name
	Next

	If bFound = False Then
		WriteLog "sms_ccm_softwaredistribution_exists","Policy has not been retrieved","failure",advertisementid
	End If

	Set oClientActions = Nothing

	Set oSWbemPolicy = Nothing
	Set oClientActions = Nothing

End Sub

Function QueryActiveDirectory(Query)
	On Error Resume Next

	Const ADS_SCOPE_SUBTREE = 2
	Dim objConnection, objCommand

	Set objConnection = CreateObject("ADODB.Connection")
	Set objCommand =   CreateObject("ADODB.Command")

	objConnection.Provider = "ADsDSOObject"
	objConnection.Open "Active Directory Provider"
	Set objCOmmand.ActiveConnection = objConnection

	objCommand.CommandText = Query
	objCommand.Properties("Page Size") = 10
	objCommand.Properties("Timeout") = 20
	objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE
	objCommand.Properties("Cache Results") = False
	objCommand.Properties("Chase Referrals") = True

	Set QueryActiveDirectory = objCommand.Execute
	
	Set objCommand = Nothing
	Set objConnection = Nothing
End Function

Function ConvertDomainToQuery(DomainName)
	On Error Resume Next
	Dim oArray, TempString, x
	
	If Len(DomainName) = 0 Then
		Exit Function
	End If
	oArray = Split(DomainName,".")
	For x = LBound(oArray) To UBound(oArray)
		TempString = TempString & "DC=" & oArray(x) & ","
	Next
	ConvertDomainToQuery = "GC://" & Mid(TempString,1,Len(TempString)-1)
End Function

Sub failedsystemapplications()
	On Error Resume Next
	Dim CountFailed : CountFailed = 0
	
	' --------------------------  Get system history
'	ReDim agHistoryPackageID(-1)
'	ReDim agHistoryState(-1)
'	ReDim agHistoryRunStartTime(-1)

	keyPath = "SOFTWARE\Microsoft\SMS\Mobile Client\Software Distribution\Execution History\System"
	oRegistry.EnumKey HKEY_LOCAL_MACHINE,keyPath,aSystemPackageIDs

	For Each PackageID In aSystemPackageIDs
		oRegistry.EnumKey HKEY_LOCAL_MACHINE,keyPath & "\" & PackageID,aStrangeNumbers
		For Each StrangeNumber In aStrangeNumbers
			State = ""
			RunStartTime = ""
			ProgramID = ""
'			oRegistry.GetStringValue HKLM,keyPath & "\" & PackageID & "\" & StrangeNumber, "_ProgramID",ProgramID
'			oRegistry.GetStringValue HKLM,keyPath & "\" & PackageID & "\" & StrangeNumber, "_RunStartTime",RunStartTime
			oRegistry.GetStringValue HKEY_LOCAL_MACHINE,keyPath & "\" & PackageID & "\" & StrangeNumber, "_State",State
			If State = "Failure" Then
				WriteLog "failedsystemapplications",PackageId,"failure",State
				CountFailed = CountFailed + 1
			End If
			
'			ReDim preserve agHistoryPackageID(UBound(agHistoryPackageID)+1)
'			ReDim preserve agHistoryState(UBound(agHistoryState)+1)
'			ReDim preserve agHistoryRunStartTime(UBound(agHistoryRunStartTime)+1)
			
'			agHistoryPackageID(UBound(agHistoryPackageID)) = PackageID & ":" & ProgramID
'			agHistoryState(UBound(aghistoryState)) = State
'			agHistoryRunStartTime(UBound(agHistoryRunStartTime)) = RunStartTime
		Next
	Next
	
	If CountFailed > 0 Then
		WriteLog "failedsystemapplications","Failed system programs detected in SMS execution history","failure",CountFailed
	Else
		WriteLog "failedsystemapplications","No failed system programs detected in SMS execution history","success","0"
	End If 

End Sub 

Sub LogicalDeviceError()
	On Error Resume Next
	Dim oStuff, bFound
	bFound = False
	Set oStuff = oSWbemServices.ExecQuery("Select Name, Status, PNPDeviceID from CIM_LogicalDevice where status='Error' and LastErrorCode <> '<empty>'")

	For Each Item In oStuff
		bFound = True
		WriteLog "logicaldeviceerror","Device Error: " & Item.Name & ", Status=" & Item.Status,"failure",Item.PNPDeviceID
	Next

	If bFound = False Then
		WriteLog "logicaldeviceerror","All devices are functional","success",""
	End If 
	Set oStuff = Nothing
End Sub

Sub osd_schedulebuild(node)
	On Error Resume Next
	Dim ADSServer, Status, UnknownComputer
	

	Dim PartitionProfile, SCCMServer, NetBIOSName, MigrationDate, NetBIOSDomain, MACAddress, AssetTag, SerialNumber, ReferenceUser, ReferenceWorkstation, ChangeNumber , Comment, Contact, TaskSequence, TargetDomainName, SysprepProfileName
	Dim DNSDomain, tempOU, gStatus
	
	SCCMServer = node.getattributenode("server").value
	Timezone = oUserDataXML.selectsinglenode("data/variable[(@name='timezone')]").getattributenode("value").text
	TaskSequence = oUserDataXML.selectsinglenode("data/variable[(@name='tasksequence')]").getattributenode("value").text
	DNSDOmain = oUserDataXML.selectsinglenode("data/variable[(@name='dnsdomain')]").getattributenode("value").text
	NetBIOSDomain = Left(DNSDOmain,InStr(DNSDOmain,".")-1)
	NetBIOSName = ComputerName
	MACAddress = oUserDataXML.selectsinglenode("data/variable[(@name='macaddress')]").getattributenode("value").text
	ReferenceWorkstation = oUserDataXML.selectsinglenode("data/variable[(@name='referenceworkstation')]").getattributenode("value").text
	ReferenceUser = oUserDataXML.selectsinglenode("data/variable[(@name='referenceuser')]").getattributenode("value").text
	AssetTag = oUserDataXML.selectsinglenode("data/variable[(@name='assettag')]").getattributenode("value").text
	SerialNumber = oUserDataXML.selectsinglenode("data/variable[(@name='serialnumber')]").getattributenode("value").text
	ChangeNumber = oUserDataXML.selectsinglenode("data/variable[(@name='changenumber')]").getattributenode("value").text
	Comment = oUserDataXML.selectsinglenode("data/variable[(@name='comment')]").getattributenode("value").text
	ScheduledBy = oNetwork.UserName & "@" & oNetwork.UserDomain & ".com"
	MandatoryTime = oUserDataXML.selectsinglenode("data/variable[(@name='mandatorytime')]").getattributenode("value").text

	If Len(MandatoryTime) < 1 Then
		MandatoryTime = "18/02/2010 2:16:27 PM"
	End If 

	' Verify we have a valid mandatory time
	If IsDate(MandatoryTime) = False Then
		WriteLog "osd_schedulebuild","Schedule Build (" & SCCMServer & ")","failure","Invalid Mandatory Date"
		Exit Sub
	End If

	MandatoryTime1 = CDate(MandatoryTime)

	' Verify we have a valid unique identifier
	If (Len(MACAddress) < 1 and Len(AssetTag) < 1 And Len(SerialNumber) < 1) then
		WriteLog "osd_schedulebuild","Schedule Build (" & SCCMServer & ")","failure","Invalid unique identifier"
		Exit Sub
	End If

	' Bit messy
	SCCMSiteCode = GetSiteCodeForServer(SCCMServer)

	' Calculate the central server of the selected server
	Select Case SCCMServer
		Case "IAUNSW456.AU.CBAINET.COM"
			SCCMCentralServer = "IAUNSW456.AU.CBAINET.COM"
			SCCMCentralDatabase = "SMS_PRD"
		Case "IAUNSW457.AU.CBAINET.COM"
			SCCMCentralServer = "IAUNSW456.AU.CBAINET.COM"
			SCCMCentralDatabase = "SMS_PRD"
		Case "IAUQLD064.AU.CBAINET.COM"
			SCCMCentralServer = "IAUNSW456.AU.CBAINET.COM"
			SCCMCentralDatabase = "SMS_PRD"
		Case "IAUSA035.AU.CBAINET.COM"
			SCCMCentralServer = "IAUNSW456.AU.CBAINET.COM"
			SCCMCentralDatabase = "SMS_PRD"
		Case "IAUWA040.AU.CBAINET.COM"
			SCCMCentralServer = "IAUNSW456.AU.CBAINET.COM"
			SCCMCentralDatabase = "SMS_PRD"
		Case "IAUNSWT72.AUT01.CBAITEST01.COM"
			SCCMCentralServer = "IAUNSWT72.AUT01.CBAITEST01.COM"
			SCCMCentralDatabase = "SMS_PRE"
		Case "IAUNSWT73.AUT01.CBAITEST01.COM"
			SCCMCentralServer = "IAUNSWT72.AUT01.CBAITEST01.COM"
			SCCMCentralDatabase = "SMS_PRE"
		Case Else
			SCCMCentralServer = "IAUNSW456.AU.CBAINET.COM"
			SCCMCentralDatabase = "SMS_PRD"
	End Select

	' Is the build already scheduled in the EDS_OSDBuildData table on the central
	if IsBuildAlreadyScheduled(MACAddress,AssetTag,SerialNumber,SCCMCentralServer,SCCMCentralDatabase) Then 
		WriteLog "osd_schedulebuild","Schedule Build (" & SCCMCentralServer & ")","failure","Build already scheduled"
		Exit Sub
	End If 

	' Set up a connection to the local provider.
	Set swbemconnection = owbemLocator.ConnectServer(SCCMServer , "root\sms\site_" + SCCMSiteCode)

	' Is this an unknown computer
	' If it is then we may not need to create collection/advertisement
	If Len(MACAddress) > 0 Then
		if IsUnknownComputer (swbemconnection,MACAddress) Then 
			UnknownComputer = True
		Else
			UnknownComputer = True
		End If 
	Else
		UnknownComputer = False
	End If 

	' Get the comptuers resourceID
	ResourceID = GetResourceID(swbemconnection,NetBIOSName)
	
	If Len(ResourceID) < 1 Then
		' Unknown Computer
		UnknownComputer = True
	End If 

	If UnknownComputer = False Then
		
		' What's the task sequence package id
		TaskSequenceID = GetOSDTaskSequenceID(swbemconnection,TaskSequence)
		If Len(TaskSequenceID) < 1 Then
			WriteLog "osd_schedulebuild","Schedule Build (" & SCCMCentralServer & ")","failure","Could not determine the Task Sequence ID."
			Exit Sub 
		End If 
		' Create the collection
		CollectionName = "OSD." & NetBIOSName
		CollectionComment = "Scheduled By: " & ScheduledBy
		ownedByThisSite = True
		ParentCollectionID = GetCollectionID(swbemconnection,"Active OSD Deployments")
		CreateStaticCollection swbemconnection, ParentCollectionID, CollectionName, CollectionComment, true, "SMS_R_System", ResourceID
	
		' Get the ID of the new collection
		CollectionID = GetCollectionID(swbemconnection,CollectionName)
	
		' Create collection variables so that client can locate EDS_OSDBuildData SQL Table
		CreateCollectionVariable swbemconnection, "OSDDatabaseServer",SCCMCentralServer,False,CollectionID,5
		CreateCollectionVariable swbemconnection, "OSDDatabaseName",SCCMCentralDatabase,False,CollectionID,5
	
		' Create the advertisement
		AdvertisementName = "OSD." & NetBIOSName
		AdvertisementComment = "Scheduled by: " & ScheduledBy 
		AdvertisementFlags = "49676288" 
		newAdvertisementStartOfferDateTime = ConvertToWMIDate(MandatoryTime1)
		newAdvertisementStartOfferEnabled = True
		AdvertisementMandatoryTime = MandatoryTime1
		CreateAdvertisement swbemconnection, CollectionID, TaskSequenceID , "*", AdvertisementName, newAdvertisementComment, AdvertisementFlags, newAdvertisementStartOfferDateTime, newAdvertisementStartOfferEnabled


		' Get the ID of the new advertisement
		AdvertisementID = GetOSDAdvertisementID(swbemconnection,AdvertisementName)
	
		' Create the mandatory For the advertisement
		AddSchedTokenOneOffMandatory swbemconnection,AdvertisementMandatoryTime,AdvertisementID
		
	
		' Move the advertisement to the appropriate folder
		MoveAdvertisementToFolder swbemconnection,AdvertisementID, "Active OSD Deployments"

		WriteLog "osd_schedulebuild","Schedule Build (" & ADSServer & ")","success","AdvertisementID: " & AdvertisementID
		WriteLog "osd_schedulebuild","Schedule Build (" & ADSServer & ")","success","TaskSequenceID: " & TaskSequenceID
		
	
	End If ' Unknwon computer

	Status = "Ready"
	' Write the data to the SQL table
	ConnectionString = "Provider='SQLOLEDB';Data Source='" & SCCMCentralServer & "';Initial Catalog='" & SCCMCentralDatabase & "' ;Integrated Security='SSPI';"
	Query = "exec dbo.eds_OSDSetBuildData " & _
	"'" & NetBIOSName & "'," & _
	"'" & DNSDOmain  & "'," & _
	"'" & TimeZone & "'," & _
	"'" & MACAddress  & "'," & _
	"'" & AssetTag  & "'," & _
	"'" & SerialNumber  & "'," & _
	"'" & ReferenceUser  & "'," & _
	"'" & ReferenceWorkstation  & "'," & _
	"'" & ChangeNumber  & "'," & _
	"'" & Comment  & "'," & _
	"'" & cstr(UnknownComputer) & "',"  & _
	"'" & ScheduledBy  & "'," & _
	"'" & Status  & "'," & _
	"'" & SCCMServer  & "'," & _
	"'" & AdvertisementID  & "'"
	DBQuery Query,ConnectionString
	
	WriteLog "osd_schedulebuild","Schedule Build (" & ADSServer & ")","success","UnknownComputer: " & CStr(UnknownComputer)
	WriteLog "osd_schedulebuild","Schedule Build (" & ADSServer & ")","success","Build successfully scheduled"

	Set swbemconnection = Nothing

	
End Sub 

Function ConvertToWMIDate(strDate)
	On Error Resume Next
	
	Dim StrYear, strDay, strMinute 
    'Convert from a standard date time to wmi date
    '4/18/2005 11:30:00 AM = 2005041811300.000000+*** 
    StrYear = year(strDate)
    strMonth = month(strDate)
    strDay = day(strDate)
    strHour = hour(strDate)
    strMinute = minute(strDate)

    'Pad single digits with leading zero
    if len(strmonth) = 1 then strMonth = "0" & strMonth
    if len(strDay) = 1 then strDay = "0" & strDay
    if len(strHour) = 1 then strHour = "0" & strHour
    if len(strMinute) = 1 then strMinute = "0" & strMinute
    ConvertToWMIDate = strYear & strMonth & strDay & strHour & strMinute & "00.000000+***"
end Function


Sub osd_cancelbuild(node)
	On Error Resume Next
	Dim ADSServer, Status
	DIm MACADDRESS

	ADSServer = node.getattributenode("server").value

	DNSDomain = oUserDataXML.selectsinglenode("data/variable[(@name='dnsdomain')]").getattributenode("value").text
	OSIMAGE = oUserDataXML.selectsinglenode("data/variable[(@name='osimage')]").getattributenode("value").text
'	PARTITIONPROFILE = oUserDataXML.selectsinglenode("data/variable[(@name='partitionprofile')]").getattributenode("value").text
	PRIMARYUSER = oUserDataXML.selectsinglenode("data/variable[(@name='primaryuser')]").getattributenode("value").text
	ReferenceWorkstation = oUserDataXML.selectsinglenode("data/variable[(@name='referenceworkstation')]").getattributenode("value").text
	SYSPREPPROFILE = oUserDataXML.selectsinglenode("data/variable[(@name='sysprepprofile')]").getattributenode("value").text
	MACADDRESS = oUserDataXML.selectsinglenode("data/variable[(@name='macaddress')]").getattributenode("value").text
	ASSETTAG = oUserDataXML.selectsinglenode("data/variable[(@name='assettag')]").getattributenode("value").text
	SERIALNUMBER = oUserDataXML.selectsinglenode("data/variable[(@name='serialnumber')]").getattributenode("value").text
	BATCHDESCRIPTION = oUserDataXML.selectsinglenode("data/variable[(@name='description')]").getattributenode("value").text
	
	if CancelBuild (ADSServer, ComputerName, MACADDRESS, ASSETTAG, SERIALNUMBER) = True Then 
		WriteLog "osd_cancelbuild","Cancel Build (" & ADSServer & ")","success","Build successfully cancelled"
	Else
		WriteLog "osd_cancelbuild","Cancel Build (" & ADSServer & ")","failure","Error cancelling build"
	End If 
	
End Sub 

Function CancelBuild (SCCMCentralServer, ComputerName, MACAddress, AssetTag, SerialNumber)
	On Error Resume Next
	Dim MigrationDate, BatchName, batchID, contact, NetBIOSDomain, PUID, strXMLData

	CancelBuild = False
'	if CheckADSAccess(ADSServer) = False Then
'		WriteLog "osd_cancelbuild","Cancel Build (" & ADSServer & ") Failed","failure","Populate a username and password under Preferences\Usernames."
'		Exit Function
'	End If
	
	' Derived values
	MigrationDate = Day(Now()) & "/" & Month(Now()) & "/" & Year(Now())
	BatchName = ComputerName
	batchID = ComputerName
	' Contat
	if len(g_WSUser) > 0 then
		aArray = split(g_WSUser,"\")
		contact = aArray(1) & "@" & aArray(0) & ".com"
	else
		contact = oNetwork.UserName & "@" & oNetwork.UserDomain & ".com"
	End If

	' PUID
	If Len(MACAddress) > 0 and instr(macaddress,"*") < 0 Then
		PUID = ComputerName
	elseif instr(Macaddress,"*") > 0 then
		PUID = replace(MACAddress,"*","")
	Elseif Len(AssetTag) > 0 Then
		PUID = AssetTag
	Elseif Len(SerialNumber) > 0 Then
		PUID = SerialNumber
	Else
		WriteLog "osd_cancelbuild","Cancel Build (" & SCCMCentralServer & ") Failed","failure","Could not determine PUID."
		Exit Function
	End If

	Select Case SCCMCentralServer
		Case "IAUNSW456.AU.CBAINET.COM"
			SCCMCentralDatabase = "SMS_PRD"
			SCCMSiteCode = "PRD"
		Case "IAUNSWT72.AUT01.CBAITEST01.COM"
			SCCMCentralDatabase = "SMS_PRE"
			SCCMSiteCode = "PRE"
		Case Else
		
	End Select

	' Set up a connection to the local provider.
	Set swbemconnection = owbemLocator.ConnectServer(SCCMCentralServer , "root\sms\site_" + SCCMSiteCode)

	' What's the Advertisement Name
	AdvertisementName = "OSD." & ComputerName

	' Get the ID of the new advertisement
	AdvertisementID = GetOSDAdvertisementID(swbemconnection,AdvertisementName)

	' What's the colleciton Name
	CollectionName = "OSD." & ComputerName

	CollectionID = GetCollectionID(swbemconnection,CollectionName)

	' Delete advertisement
	DeleteAdvertisement swbemconnection,AdvertisementID

	' Delete Collection
	DeleteCollection swbemconnection, CollectionID


	if IsBuildAlreadyScheduled(macaddress,AssetTag,SerialNumber,SCCMCentralServer,SCCMCentralDatabase) = True Then 
		ConnectionString = "Provider='SQLOLEDB';Data Source='" & SCCMCentralServer & "';Initial Catalog='" & SCCMCentralDatabase & "' ;Integrated Security='SSPI';"
		Query = "exec dbo.eds_OSDCancelBuildData '" & MACaddress & "','" & AssetTag & "','" & SerialNumber & "'"
		DBQuery Query,ConnectionString
		WriteLog "osd_cancelbuild","Cancel Build (" & SCCMCentralServer & ")","success",""
		CancelBuild = True
	Else
		WriteLog "osd_cancelbuild","Cancel Build (" & SCCMCentralServer & ")","success","No build was scheduled."
		CancelBuild = True
	End If 


End Function

Function ScheduleBuild (ADSServer, ComputerName, DNSDomain, OSPackageName, PartitionProfile, PrimaryUser, ReferenceWorkstation, SysprepProfileName, macaddress, AssetTag, SerialNumber, BatchDescription, ByRef Status)
	On Error Resume Next
	Dim MigrationDate, BatchName, batchID, contact, NetBIOSDomain, PUID, strXMLData

	ScheduleBuild = False
	if CheckADSAccess(ADSServer) = False Then
		WriteLog "osd_schedulebuild","Schedule Build (" & ADSServer & ") Failed","failure","Populate a username and password under Preferences\Usernames."
		Exit Function
	End If

	' Derived values
	MigrationDate = Day(Now()) & "/" & Month(Now()) & "/" & Year(Now())
	BatchName = ComputerName
	batchID = ComputerName
	' Contact
	contact = oNetwork.UserName & "@" & oNetwork.UserDomain & ".com"
	NetBIOSDomain = Left(DNSDOmain,InStr(DNSDOmain,".")-1)

	' PUID
	If Len(MACAddress) > 0 Then
		PUID = ComputerName
	Elseif Len(AssetTag) > 0 Then
		PUID = AssetTag
	Elseif Len(SerialNumber) > 0 Then
		PUID = SerialNumber
	Else
		WriteLog "osd_schedulebuild","Schedule Build (" & ADSServer & ") Failed","failure","Could not determine PUID."
		g_Errors = True
		Exit Function
	End If

	' Delete existing batch
	strXMLData = DeleteExistingBatchXML (batchName,batchID,ComputerName, PUID)
	If Not callWEBMethod("PerformXmlOperations","xmlCommands=" + strXMLData, responseXML,ADSServer) Then
		If ErrorInXML(responseXML) = True Then
			' Don't care
'	 		document.getElementById("aRebuildStatus").innerHTML = "Eror cancelling build."
'			Exit Sub
		End If
	End If

	' Delete existing data
	strXMLData = DeleteExistingDataXML (batchName,batchID,ComputerName, PUID)
	If Not callWEBMethod("PerformXmlOperations","xmlCommands=" + strXMLData, responseXML,ADSServer) Then
		' Don't care
' 		MsgBox "Error" & VbCrLf & responseXML
' 		document.getElementById("aRebuildStatus").innerHTML = "Eror scheduling build."
' 		Exit Sub 
	End If

	' Generate the import data
	strXMLData = GetComputerImportDataXML(ComputerName, NetBIOSDomain, DNSDomain, MACAddress, AssetTag, SerialNumber, PrimaryUser, ReferenceWorkstation,PUID)
	If Not callWEBMethod("PerformXmlOperations","xmlCommands=" + strXMLData, responseXML,ADSServer) Then
		WriteLog "osd_schedulebuild","Schedule Build (" & ADSServer & ") Failed on Generate Import Data","failure",responseXML
		g_Errors = True
		Exit Function
 	Else
		If ErrorInXML(responseXML) = True Then
			WriteLog "osd_schedulebuild","Schedule Build (" & ADSServer & ") Failed on Generate Import Data","failure",responseXML
			g_Errors = True
			Exit Function
		End If
	End If

	' Create the association
	strXMLData = GetComputerAssociationXML(ComputerName,PUID)
	If Not callWEBMethod("PerformXmlOperations","xmlCommands=" + strXMLData, responseXML,ADSServer) Then
		WriteLog "osd_schedulebuild","Schedule Build (" & ADSServer & ") Failed on Create Association","failure",responseXML
		g_Errors = True
		Exit Function
 	Else
		if ErrorInXML(responseXML) = True Then
			WriteLog "osd_schedulebuild","Schedule Build (" & ADSServer & ") Failed on Create Association","failure",responseXML
			g_Errors = True
			Exit Function
		End If
	End If

	' Create the batch
	strXMLData = GetNewBatchXML(batchName,batchID,BatchDescription,Contact,MigrationDate,OSPackageName,NetBIOSDomain, SysprepProfileName,PartitionProfile)
	If Not callWEBMethod("PerformXmlOperations","xmlCommands=" + strXMLData, responseXML,ADSServer) Then
		WriteLog "osd_schedulebuild","Schedule Build (" & ADSServer & ") Failed on Create Batch","failure",responseXML
		g_Errors = True
		Exit Function
 	Else
		if ErrorInXML(responseXML) = True Then
			WriteLog "osd_schedulebuild","Schedule Build (" & ADSServer & ") Failed on Create Batch","failure",responseXML
			g_Errors = True
			Exit Function
		End If
	End If 

	' Add the compuer to the batch
	strXMLData = AddComputerToBatchXML(batchName,PUID)
	If Not callWEBMethod("PerformXmlOperations","xmlCommands=" + strXMLData, responseXML,ADSServer) Then
		WriteLog "osd_schedulebuild","Schedule Build (" & ADSServer & ") Failed on Add Computer to Batch","failure",responseXML
		g_Errors = True
		Exit Function
 	Else
		if ErrorInXML(responseXML) = True Then
			WriteLog "osd_schedulebuild","Schedule Build (" & ADSServer & ") Failed on Add Computer to Batch","failure",responseXML
			g_Errors = True
			Exit Function
		End If
	End If

	' Release the batch
	strXMLData = GetReleaseBatchXML(batchName, batchID)
	If Not callWEBMethod("PerformXmlOperations","xmlCommands=" + strXMLData, responseXML,ADSServer) Then
		WriteLog "osd_schedulebuild","Schedule Build (" & ADSServer & ") Failed on Release Batch","failure",responseXML
		g_Errors = True
		Exit Function
 	Else
		if ErrorInXML(responseXML) = True Then
			WriteLog "osd_schedulebuild","Schedule Build (" & ADSServer & ") Failed on Release Batch","failure",responseXML
			g_Errors = True
			Exit Function
		End If
	End If

	' Write an audit log
	if len(g_WSUser) > 0 Then
		uName = g_WSUser
	Else
		uName = oNetwork.UserDomain & "\" & oNetwork.UserName 
	End If
	AuditBuildSchedule uName,"ScheduleBuild",ComputerName,adsserver

	ScheduleBuild = True
End Function

Sub AuditBuildSchedule(uName,Action,NetBIOSName,adsserver)
	On Error Resume Next
	Dim ConnectionString, Query, RS

	ConnectionString = "Provider='SQLOLEDB';Data Source='" & adsServer & "';Initial Catalog='ADS' ;Integrated Security='SSPI';"
	Query = "exec dbo.USP_CBAAudit '" & uName & "','" & Action & "','" & NetBIOSName & "'"
'	MsgBox "ConnectionString: " & ConnectionString
'	MsgBox "Query: " & Query
	
	Set rs = DBQuery (Query,ConnectionString)
	Set RS = Nothing
End Sub

' Return a recordset
Function DBQuery(query,ConnectionString)
	On Error Resume Next
	Dim oConnection, oRecordSet
	Dim bConnected : bConnected = False
	Dim dbServer : dbServer = ""

	Set oConnection = CreateObject("ADODB.Connection")
	Set oRecordSet = CreateObject("ADODB.RecordSet")

	oConnection.ConnectionString = ConnectionString
	oConnection.ConnectionTimeout = 10
	oConnection.CommandTimeout = 20

	' Just try connecting to the database
	Err.Clear
	oConnection.Open
	If Err.Number = 0 Then
		bConnected = True
	Else
		' calculate the DB server just incase we need to map a drive
		xx = InStr(ucase(connectionstring),"DATA SOURCE=")
		yy = InStr(xx+13,ConnectionString,"'")
		If xx > 0 Then
			dbServer = Mid(connectionstring,xx+13,yy-(xx+13))
		End If
	End If

	' If the username for WSUer is set then try connecting with this
	if bConnected = False Then
		Err.Clear
		If Len(g_DBUser) > 0 Then
			oConnection.ConnectionString = ConnectionString & " Network Library='dbnmpntw'; "
			oConnection.Open
			If Err.Number = 0 Then
				bConnected = True
			Else
				g_DBUser = ""
				g_DBPass = ""
			End If
		End If 
	End If

	If bConnected = False then
		If InStr(oConnection.ConnectionString,"Network Library") < 1 Then
			oConnection.ConnectionString = ConnectionString & " Network Library='dbnmpntw'; "
		End If
		MapNetworkDrive "", "\\" & dbServer & "\IPC$", False, g_DBUser, g_DBPass
		oConnection.Open
		If Err.Number = 0 Then
			bConnected = True
		Else
			g_DBUser = ""
			g_DBPass = ""
		End If
	End If

	if bConnected = False Then
		WriteLog "dbquery","Error connecting to SQL server","failure","ConnectionString: " & ConnectionString 
		dbquery = ""
		Exit Function
	End If

	Err.Clear 
	oRecordSet.Open Query, oConnection, adOpenStatic, adLockOptimistic
	If Err.Number <> 0 Then
		WriteLog "dbquery","Error executing SQL query","failure","Query: " & Query 
		DBQuery = ""
		Exit Function
	End If

	If oRecordSet.State = 0 Then
		Set oRecordSet = Nothing
		Set oConnection = Nothing 
		WriteLog "dbquery","No records returned","success","Query: " & Query 
		Set DBQuery = Nothing
		Exit Function
	End If 	
	
	If oRecordSet.RecordCount < 1 Then
		WriteLog "dbquery","No records returned","success","Query: " & Query 
		Exit Function
	End If 

	x = 0
	oRecordSet.MoveFirst
	If oRecordSet.EOF Then
		WriteLog "dbquery","No records returned","success","Query: " & Query 
		DBQuery = "No records found"
		Exit Function
	End If

	set DBQuery = oRecordSet
	' oConnection.Close
	Set oRecordSet = Nothing
	Set oConnection = Nothing
	
	WriteLog "dbquery","Query executed successfully","success","Query: " & Query 

End Function

Sub MapNetworkDrive (letter, path, username, password)
	On Error Resume Next
	Err.Clear
	oNetwork.MapNetworkDrive letter, path, False, username, password 
	If Err.Number <> 0 Then
		MsgBox "error mapping drive" & Err.Number & ", " & Err.Description 
		Err.Clear 
	End If
End Sub

Function ErrorInXML(xml)
	On Error Resume Next
	ErrorInXML = False
	Set mick = xml.DocumentElement.selectNodes("OperationResult")
	For Each Node In mick
		If node.selectSingleNode("Status").Text = "Failure" Then
			Set Errors = node.selectNodes("Errors")
			For Each er In Errors
				WriteLog "callwebmethod","ErrorInXML","failure","ERROR: " & er.selectSingleNode("string").Text
				ErrorInXML = True
				Exit Function 
			Next
		End If
	Next

End Function

Function CheckADSAccess(ADSServer)
	On Error Resume Next
	Dim nodes, node, responseXML

	CheckADSAccess = False
	If callWEBMethod("GetSysprepProfileList","", responseXML, ADSServer) = False Then
		For each Node in oDataXML.selectSingleNode("data/users").ChildNodes 
			Username = node.GetAttributeNode ("name").value
			Password = node.GetAttributeNode ("password").value
			if ucase(Username) = "TOPTIGER" Then
				Username = ComputerName & "\toptiger"
			elseIf Len(Username) > 0 And Len(Password) > 0 Then
'				msgbox Username & vbcrlf & Password
				g_WSUser = Username
				g_WSPass = Password
				Err.Clear
				If callWEBMethod("GetSysprepProfileList","", responseXML, ADSServer) Then
					g_WSUser = Username
					g_WSPass = Password
					CheckADSAccess = True
					Exit Function
				End If
			End If
		Next
	Else
		CheckADSAccess = True
		Exit Function 
	End If
		
	CheckADSAccess = False
End Function

Sub SaveToFile (fname,data)
	On Error Resume Next

	Set oFile = oFSO.OpenTextFile (fname,2,True)
	oFile.writeline data
	ofile.close
	set oFile = Nothing

End Sub

Function QueryBuild (ADSServer, ComputerName, DNSDomain, OSPackageName, PartitionProfile, PrimaryUser, ReferenceWorkstation, ByRef ReferenceWorkstationExists, SysprepProfileName, MACAddress, AssetTag, SerialNumber, BatchDescription, Schedule, ByRef Status)
	On Error Resume Next
	Dim MigrationDate, BatchName, batchID, contact, NetBIOSDomain, PUID, strXMLData
	Dim ConnectionString, Query, RS

	' Default values
	Status = "Error querying build"

	' Derived values
	MigrationDate = Day(Now()) & "/" & Month(Now()) & "/" & Year(Now())
	BatchName = ComputerName
	batchID = ComputerName
	contact = oNetwork.UserName & "@" & oNetwork.UserDomain & ".com"
	NetBIOSDomain = Left(DNSDOmain,InStr(DNSDOmain,".")-1)

	' PUID
	If Len(MACAddress) > 0 Then
		PUID = ComputerName
	Elseif Len(AssetTag) > 0 Then
		PUID = AssetTag
	Elseif Len(SerialNumber) > 0 Then
		PUID = SerialNumber
	Else
		SetStatus "iLastError","Could not determine PUID"
		Status = "Could not determine PUID"
		g_Errors = True
		Exit Function
	End If

	' Check for reference file
	if ReferenceFileExists (ReferenceWorkstation,ADSServer) = True then
		ReferenceWorkstationExists = "True"
	Else
		ReferenceWorkstationExists = "False"
	End If

	ConnectionString = "Provider='SQLOLEDB';Data Source='" & ADSServer & "';Initial Catalog='ADS' ;Integrated Security='SSPI';"
	Query = "exec dbo.USP_CBASearchMigrationStatus '" & ComputerName & "'"
	Set rs = DBQuery (Query,ConnectionString)
	
	Select Case rs.Fields("MigrationStatus").value
		Case "6"
			Status = "Failed: An OSD phase ended in error (" & rs.Fields("MigrationStatus").value & ")"
		Case "4"
			Status = "DeployStarted: OS deployment in progress (" & rs.Fields("MigrationStatus").value & ")"
		Case "5"
			Status = "Successful: OS deployment completed successfully (" & rs.Fields("MigrationStatus").value & ")"
		Case "3"
			Status = "ReadyForDeploy: Waiting for OS deployment to start (" & rs.Fields("MigrationStatus").value & ")"
		Case "2"
			Status = "ReadyForScheduler: Waiting on the scheduler (" & rs.Fields("MigrationStatus").value & ")"
		Case "7"
			Status = "Stalled: An OSD phase timed-out (" & rs.Fields("MigrationStatus").value & ")"
		Case Else
			Status = "Unknown: (" & rs.Fields("MigrationStatus").value & ")"
	End Select 
	Set RS = Nothing

End Function

Function ReferenceFileExists(name,adsserver)
	On Error Resume Next
	
	ReferenceFileExists = False
	
	If oFSO.FileExists ("\\" & adsserver & "\executionhistory$\" & name & ".xml") Then
		ReferenceFileExists = True
	Else
		ReferenceFileExists = False
	End If
End Function

Function callWEBMethod(strMethodCall,strMethodArgs,byRef responseXML, ADSServer)
	On Error Resume Next
	Dim strADSURL,objXMLhttp

	callWEBMethod = False

	strADSURL = "http://" & ADSServer & "/ADSWebService/ADSWebMethods.asmx"
	Err.Clear 
   	Set objXMLhttp =  CreateObject("Microsoft.XMLHTTP")
	If Err.Number Then
		WriteLog "callwebmethod","CreateObject","failure","Cannot Create XMLHTTP: 0x" + Hex(Err.Number) + " - " + Err.Description
		Exit Function
	end If

	if len(g_WSUser) > 0 Then 
		' Mick'
'			msgbox g_WSUser & vbcrlf & g_WSPass
			Err.Clear 
		    Call objXMLhttp.open("POST",strADSURL + "/" + strMethodCall, False,g_WSUser,g_WSPass)
			If Err.Number <> 0 Then
				WriteLog "callwebmethod","CreateObject","failure","Cannot Post To ADS: 0x" + Hex(Err.Number) + " - " + Err.Description
				Exit Function
			End If
	Else
		Err.Clear
	    Call objXMLhttp.open("POST",strADSURL + "/" + strMethodCall, False)
		If Err.Number <> 0 Then
			WriteLog "callwebmethod","CreateObject","failure","Cannot Post To ADS: 0x" + Hex(Err.Number) + " - " + Err.Description
			Exit Function
		end If
	End If 
	objXMLhttp.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
	objXMLhttp.setRequestHeader "Content-Length",CStr(Len(strMethodArgs))
	objXMLhttp.setRequestHeader "Host",ADSServer

	Err.Clear
	If Len (strMethodArgs) > 0 Then
		objXMLhttp.send(strMethodArgs)
		If Err.Number Then
			WriteLog "callwebmethod","CreateObject","failure","ERROR: Cannot Create SEND: 0x" + Hex(Err.Number) + " - " + Err.Description
			Exit Function
		End If
	Else
		objXMLhttp.send()
	End If
	If (instr(objXMLhttp.statusText,"OK") =< 0) Then 
		responseXML = objXMLhttp.responseText
		callWEBMethod = False
		Exit Function
	End If
	Set responseXML = objXMLhttp.responseXML
	If Err.Number Then
		WriteLog "callwebmethod","CreateObject","failure","ERROR: failure to get responseXML: 0x" + + Hex(Err.Number) + " - " + Err.Description
		responseXML = objXMLhttp.responseText
		Exit Function
	End If
	callWEBMethod = True
End Function 

Function GetComputerAssociationXML(NetBIOSName, PUID)
	On Error Resume Next
		GetComputerAssociationXML = "<?xml version=""1.0"" encoding=""UTF-8""?>" & VbCrLf & _
		"<ADSWebServiceOperations version=""1.0"">" & vbcrlf & _
		"<operation action=""Add"">" & vbcrlf & _
	    "    <ComputerAssociation version=""1.0"" id=""" & PUID & """>" & VbCrLf & _
	    "        <Attributes>" & vbcrlf & _
	    "            <Attribute Name=""TargetPUID"" Value=""" & PUID & """ />" & vbcrlf & _
	    "            <Attribute Name=""MigrationType"" Value=""New"" />" & vbcrlf & _
	    "        </Attributes>" & vbcrlf & _
	    "    </ComputerAssociation>" & vbcrlf & _
	    "</operation>" & vbcrlf & _
		"</ADSWebServiceOperations>"
End Function

Function GetReleaseBatchXML(batchName, batchID)
	On Error Resume Next
	GetReleaseBatchXML = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbcrlf & _
		"<ADSWebServiceOperations version=""1.0"">" & vbcrlf & _
	    "<operation action=""Execute"">" & vbcrlf & _
	    "    <ReleaseMigrationBatch BatchName=""" & batchName & """ version=""1.0"" id=""" & batchID & """ />" & vbcrlf & _
	    "</operation>" & vbcrlf & _
		"</ADSWebServiceOperations>"
End Function

Function GetComputerImportDataXML(NetBIOSName, NetBIOSDomain, DNSDomain, MACAddress, AssetTag, SerialNumber, PrimaryUser, ReferenceWorkstation,PUID)
	On Error Resume Next
	GetComputerImportDataXML = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbcrlf & _
		"<ADSWebServiceOperations version=""1.0"">" & vbcrlf & _
		"<operation action=""Add"">" & vbcrlf & _
		"    <ComputerImport ComputerPUID=""" & PUID & """ version=""1.0"" id=""" & NetBIOSName & """>" & vbcrlf & _
		"        <Attributes>" & vbcrlf & _
		"            <Attribute Name=""NetBIOSName"" Value=""" & NetBIOSName & """ />" & vbcrlf & _
		"            <Attribute Name=""NetBIOSDomain"" Value=""" & NetBIOSDomain & """ />" & vbcrlf & _
		"            <Attribute Name=""DNSDomain"" Value=""" & DNSDomain & """ />" & vbcrlf & _
		"            <Attribute Name=""MACAddress"" Value=""" & MACAddress & """ />" & vbcrlf & _
		"            <Attribute Name=""AssetTag"" Value=""" & AssetTag & """ />" & vbcrlf & _
		"            <Attribute Name=""SerialNumber"" Value=""" & SerialNumber & """ />" & vbcrlf & _
		"            <Attribute Name=""UUID"" Value="""" />" & vbcrlf & _
		"            <Attribute Name=""CPUSpeed"" Value=""3000"" />" & vbcrlf & _
		"            <Attribute Name=""PrimaryUser"" Value=""" & PrimaryUser & """ />" & vbcrlf & _
		"            <Attribute Name=""HD1Capacity_GB"" Value=""40"" />" & vbcrlf & _
		"            <Attribute Name=""Memory_MB"" Value=""1024"" />" & vbcrlf & _
		"            <Attribute Name=""ComputerType"" Value=""desktop"" />" & vbcrlf & _
		"            <Attribute Name=""AllowHDFormat"" Value=""true"" />" & vbcrlf & _
		"            <Attribute Name=""Custom1"" Value=""" & ReferenceWorkstation & """ />" & vbcrlf & _
		"        </Attributes>" & vbcrlf & _
		"    </ComputerImport>" & VbCrLf & _
		"</operation>" & vbcrlf & _
		"</ADSWebServiceOperations>"
End Function

Function DeleteExistingBatchXML(batchName,batchID,NetBIOSName, PUID)
	On Error Resume Next
    DeleteExistingBatchXML = "<?xml version=""1.0"" encoding=""UTF-8""?>" & VbCrLf & _
		"<ADSWebServiceOperations version=""1.0"">" & vbcrlf & _
	    "<operation action=""Delete"">" & vbcrlf & _
	    "    <MigrationBatch Name=""" & batchName & """ version=""1.0"" id=""" & batchID & """ />" & VbCrLf & _
	    "    <ComputerAssociation version=""1.0"" id=""" & NetBIOSName & """>" & vbcrlf & _
	    "        <Attributes>" & vbcrlf & _
	    "            <Attribute Name=""TargetPUID"" Value=""" & PUID & """ />" & vbcrlf & _
	    "        </Attributes>" & vbcrlf & _
	    "    </ComputerAssociation>" & vbcrlf & _
	    "</operation>" & vbcrlf & _
	    "</ADSWebServiceOperations>"
End Function

Function DeleteExistingDataXML(batchName,batchID,NetBIOSName, PUID)
	On Error Resume Next
    DeleteExistingDataXML = "<?xml version=""1.0"" encoding=""UTF-8""?>" & VbCrLf & _
		"<ADSWebServiceOperations version=""1.0"">" & vbcrlf & _
	    "<operation action=""Delete"">" & vbcrlf & _
	    "    <MigrationBatch Name=""" & batchName & """ version=""1.0"" id=""" & batchID & """ />" & VbCrLf & _
	    "    <ComputerAssociation version=""1.0"" id=""" & NetBIOSName & """>" & vbcrlf & _
	    "        <Attributes>" & vbcrlf & _
	    "            <Attribute Name=""TargetPUID"" Value=""" & PUID & """ />" & vbcrlf & _
	    "        </Attributes>" & vbcrlf & _
	    "    </ComputerAssociation>" & vbcrlf & _
	    "    <ComputerImport ComputerPUID=""" & PUID & """ version=""1.0"" id=""" & NetbiosName & """ />" & vbcrlf & _
	    "</operation>" & vbcrlf & _
	    "</ADSWebServiceOperations>"
End Function

Function DeleteExistingBatchDataXML(batchName,batchID)
	On Error Resume Next
    DeleteExistingBatchDataXML = "<?xml version=""1.0"" encoding=""UTF-8""?>" & VbCrLf & _
		"<ADSWebServiceOperations version=""1.0"">" & vbcrlf & _
	    "<operation action=""Delete"">" & vbcrlf & _
	    "    <MigrationBatch Name=""" & batchName & """ version=""1.0"" id=""" & batchID & """ />" & VbCrLf & _
	    "</operation>" & vbcrlf & _
	    "</ADSWebServiceOperations>"
End Function

Function AddComputerToBatchXML(batchName,PUID)
    On Error Resume Next
    AddComputerToBatchXML = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbcrlf & _
		"<ADSWebServiceOperations version=""1.0"">" & vbcrlf & _
	    "<operation action=""Execute"">" & VbCrLf & _
	    "    <AddComputersToBatch BatchName=""" & batchName & """ OnlyOneExpected=""true"" version=""1.0"" id=""" & batchName & """>" & VbCrLf & _
	    "        <Criteria>" & VbCrLf & _
	    "            <Criterion AttributeGroup=""Computer"" AttributeName=""Provider unique ID"" Operator=""Equal"" AttributeValue=""" & PUID & """/>" & VbCrLf & _
	    "        </Criteria>" & VbCrLf & _
	    "    </AddComputersToBatch>" & VbCrLf & _
	    "</operation>" & VbCrLf & _
	    "</ADSWebServiceOperations>"
End Function

Function GetAssocationXML(PUID)
    On Error Resume Next
    GetAssocationXML = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbcrlf & _
		"<ADSWebServiceOperations version=""1.0"">" & vbcrlf & _
		"<operation action=""Add"">" & VbCrLf & _
	    "    <ComputerAssociation version=""1.0"" id=""ComputerAssociationAdd1"">" & VbCrLf & _
	    "        <Attributes>" & VbCrLf & _
	    "            <Attribute Name=""TargetPUID"" Value=""" & PUID & """ />" & VbCrLf & _
	    "            <Attribute Name=""MigrationType"" Value=""New"" />" & VbCrLf & _
	    "        </Attributes>" & VbCrLf & _
	    "    </ComputerAssociation>" & VbCrLf & _
	    "</operation>" & VbCrLf & _
	    "</ADSWebServiceOperations>"

End Function

Function GetNewBatchXML(batchName,batchID,desc,Contact,MigrationDate,OSPackageName,TargetDomainName, SysprepProfileName,PartitionProfileName)
	On Error Resume Next
	GetNewBatchXML = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbcrlf & _
		"<ADSWebServiceOperations version=""1.0"">" & vbcrlf & _
		"  <operation action=""Add"">" & vbcrlf & _
		"    <MigrationBatch Name=""" & batchName & """ version=""1.0"" id=""" & BatchID & """>" & VbCrLf & _
		"      <Attributes>" & vbcrlf & _
		"        <Attribute Name=""SysprepProfileName"" Value=""" & SysprepProfileName & """ />" & vbcrlf & _
		"        <Attribute Name=""OSPackageName"" Value=""" & OSPackageName & """ />" & vbcrlf & _
		"        <Attribute Name=""AllowApplicationOverride"" Value=""False"" />" & vbcrlf & _
		"        <Attribute Name=""AllowComputerRoleOverride"" Value=""True"" />" & vbcrlf & _
		"        <Attribute Name=""Name"" Value=""" & batchName & """ />" & vbcrlf & _
		"        <Attribute Name=""TargetDomainName"" Value=""" & TargetDomainName & """ />" & vbcrlf & _
		"        <Attribute Name=""PartitionProfileName"" Value=""" & PartitionProfileName & """ />" & vbcrlf & _
		"        <Attribute Name=""Applications"" Value="""" />" & vbcrlf & _
		"        <Attribute Name=""MigrationDate"" Value=""" & MigrationDate & """ />" & vbcrlf & _
		"        <Attribute Name=""Description"" Value=""" & desc & """ />" & vbcrlf & _
		"        <Attribute Name=""NotificationProfileName"" Value=""Empty Profile"" />" & vbcrlf & _
		"        <Attribute Name=""ComputerRoleName"" Value=""Basic Computer"" />" & vbcrlf & _
		"        <Attribute Name=""Contact"" Value=""" & Contact & """ />" & vbcrlf & _
		"        <Attribute Name=""StateMigrationProfileName"" Value=""No User State Migration"" />" & vbcrlf & _
		"      </Attributes>" & vbcrlf & _
		"      <Children PreserveExistingChildObjects="""" />" & vbcrlf & _
		"    </MigrationBatch>" & vbcrlf & _
		"  </operation>" & vbcrlf & _
		"</ADSWebServiceOperations>"

End Function

Function PromptForFile(defaultFilename)
	On Error Resume Next
	Err.Clear
	cDialog.Filter="CSV Files (*.csv)|*.csv"
	cDialog.CancelError = False
	cDialog.FileName = defaultFilename
	cDialog.ShowOpen()
	If Err.Number <> 0 Then
		Exit Function
	End If

	PromptForFile = cDialog.FileName
End Function

Function btnExit_OnClick()
	On Error Resume Next
	Cleanup
	window.close()

End Function

Function addBackslash(sString)
	On Error Resume Next

	If Right(sString,1) = "\" Then
		addBackslash = sString
	Else
		addBackslash = sString & "\"
	End If
End Function

Sub smsguid
	On Error Resume Next
	
	Err.Clear 
	Set oSWbemPolicy = GetWMINamespace(ComputerName,"root\ccm")
	If Err.Number <> 0 Then
 		WriteLog "smsguid","Could not connect to root\CCM","failure",""
 		Exit Sub
	End If

	Err.Clear
	Set oClientActions = oSWbemPolicy.ExecQuery("Select * from CCM_Client")
	If Err.Number <> 0 Then
 		WriteLog "smsguid","Could not query CCM_Client","failure",""
 		Exit Sub
	End If
	
	For Each Obj In oClientActions
		WriteLog "smsguid","","success",obj.ClientID
		Field1 = obj.ClientID
	Next
	Set oClientActions = Nothing
	Set oSWbemPolicy = Nothing
End Sub

Sub SetDatabaseValue(Node)
	On Error Resume Next
	dbserver = Node.GetAttributeNode("dbserver").Value
	dbdatabase = Node.GetAttributeNode("dbdatabase").Value
	dbtable = Node.GetAttributeNode("dbtable").Value
	dbkey = Node.GetAttributeNode("dbkey").value
	dbupdatefield = Node.GetAttributeNode("dbupdatefield").value
	dbfieldvalue = Node.GetAttributeNode("dbfieldvalue").value

	Set oConnection = CreateObject("ADODB.Connection")
	Set oRecordSet = CreateObject("ADODB.RecordSet")
	oConnection.ConnectionString = ConnectionString
	oConnection.ConnectionTimeout = 30
	oConnection.CommandTimeout = 60
	oConnection.ConnectionString = "Provider='SQLOLEDB';Data Source='" & dbserver & "';Initial Catalog='" & dbdatabase & "' ;Integrated Security='SSPI';"
	Err.Clear
	oConnection.Open
	If Err.Number <> 0 Then
		WriteLog "setdatabasevalue","","failure","Error " & Err.Description & " - " & Err.Number 
		Exit sub
	End If
	
	' SetDatabaseValue dbserver="IAUNSWT72' dbdatabase="admin" dbtable="sccmmigration" dbkey="name" dbupdatefield="status" dbfieldvalue="Investigating" />
	Query = "Update " & dbtable & " set " & dbupdatefield & " = '" & dbfieldvalue & "' where " & dbkey & " = '" & ComputerName & "'"

	Err.Clear 	
	oRecordSet.Open Query, oConnection, adOpenStatic, adLockOptimistic
	If Err.Number <> 0 Then
		WriteLog "setdatabasevalue","","failure","Error " & Err.Description & " - " & Err.Number 
	Else
		WriteLog "setdatabasevalue","","success",Query
	End If 
	
	oConnection.Close
	Set oRecordSet = Nothing
	Set oConnection = Nothing
End Sub

Function GetSiteCodeForServer(SCCMServer)
	
	Select Case SCCMServer
		Case "IAUNSW456.AU.CBAINET.COM"
			GetSiteCodeForServer = "PRD"
		Case "IAUNSW457.AU.CBAINET.COM"
			GetSiteCodeForServer = "NSW"
		Case "IAUQLD064.AU.CBAINET.COM"
			GetSiteCodeForServer = "QLD"
		Case "IAUSA035.AU.CBAINET.COM"
			GetSiteCodeForServer = "SAU"
		Case "IAUWA040.AU.CBAINET.COM"
			GetSiteCodeForServer = "WAU"
		Case "IAUNSWT72.AUT01.CBAITEST01.COM"
			GetSiteCodeForServer = "PRE"
		Case "IAUNSWT73.AUT01.CBAITEST01.COM"
			GetSiteCodeForServer = "TST"
	End Select 

End Function

Function IsBuildAlreadyScheduled(MACAddress,AssetTag,SerialNumber,DatabaseServer,DatabaseName)
	On Error Resume Next
	Dim Query, oRecordSet
	IsBuildAlreadyScheduled = True
	Query = "exec dbo.EDS_OSDGetBuildData '" & MACAddress & "','" & AssetTag & "','" & SerialNumber & "'"
	Err.Clear 
	Set oRecordSet = DBQuery(Query,"Provider='SQLOLEDB';Data Source='" & DatabaseServer & "';Initial Catalog='" & DatabaseName & "' ;Integrated Security='SSPI';")
	If Err.Number <> 0 Then
		' We get an error if the returned recordset is not valid
		IsBuildAlreadyScheduled = False
		Exit Function
	End If 
	oRecordSet.MoveFirst
	If oRecordSet.EOF Then
		IsBuildAlreadyScheduled = False
	End If
	Set oRecordSet = Nothing
End Function 

' Return a recordset
Function DBQuery(query,ConnectionString)
	On Error Resume Next
	Dim oConnection, oRecordSet
	Dim bConnected : bConnected = False
	Dim dbServer : dbServer = ""

	Set oConnection = CreateObject("ADODB.Connection")
	Set oRecordSet = CreateObject("ADODB.RecordSet")
	oConnection.ConnectionString = ConnectionString
	oConnection.ConnectionTimeout = 30
	oConnection.CommandTimeout = 60

	' Just try connecting to the database
	Err.Clear
	oConnection.Open
	oRecordSet.Open Query, oConnection, adOpenStatic, adLockOptimistic

	If Err.Number = 0 Then
		bConnected = True
	Else
		' calculate the DB server just incase we need to map a drive
		xx = InStr(ucase(connectionstring),"DATA SOURCE=")
		yy = InStr(xx+13,ConnectionString,"'")
		If xx > 0 Then
			dbServer = Mid(connectionstring,xx+13,yy-(xx+13))
		End If
	End If

	' If the username for WSUer is set then try connecting with this
	if bConnected = False then
		Err.Clear
		If Len(g_DBUser) > 0 Then
'			oConnection.ConnectionString = "Network Library='dbnmpntw';" & ConnectionString
			oConnection.Open
			oRecordSet.Open Query, oConnection, adOpenStatic, adLockOptimistic
			If Err.Number = 0 Then
				bConnected = True
			Else
				g_DBUser = ""
				g_DBPass = ""
			End If
		End If 
	End If
	If bConnected = False Then
		promptForDBUsername
'		If InStr(oConnection.ConnectionString,"Network Library") < 1 Then
'			msgbox "Connection string before: " & oConnection.ConnectionString
'			' oConnection.ConnectionString = ConnectionString & "; Network Library='dbnmpntw'; "
'			msgbox "Connection string after: " & oConnection.ConnectionString
'		End If

oNetwork.MapNetworkDrive "","\\" & dbServer & "\IPC$",False,g_DBUser, g_DBPass
'		MapNetworkDrive "", "\\" & dbServer & "\ipc$", g_DBUser, g_DBPass

		oConnection.Open
		oRecordSet.Open Query, oConnection, adOpenStatic, adLockOptimistic
		If Err.Number = 0 Then
			bConnected = True
		Else
			g_DBUser = ""
			g_DBPass = ""
		End If
	End If

	if bConnected = False then
		dbquery = ""
		Exit Function
	End If

	Err.Clear 
	If Err.Number <> 0 Then
		DBQuery = ""
		Exit Function
	End If

	x = 0
	oRecordSet.MoveFirst
	If oRecordSet.EOF Then
		DBQuery = "No records found"
		Exit Function
	End If

	set DBQuery = oRecordSet
	' oConnection.Close
	Set oRecordSet = Nothing
	Set oConnection = Nothing

End Function

Function IsUnknownComputer(Connection,MACAddress)
	On Error Resume Next
    Set settings = connection.ExecQuery ("Select * From SMS_G_System_NETWORK_ADAPTER_CONFIGURATION Where MACAddress= '" & MACAddress & "'")

	IsUnknownComputer = False 
    If settings.Count = 0 Then
		IsUnknownComputer = True
    End If  

	Set settings = Nothing 
End Function 

Sub CreateCollectionVariable( connection, name, value, mask, collectionId, precedence)

    Dim collectionSettings
    Dim collectionVariables
    Dim collectionVariable
    Dim Settings
    
    ' See if the settings collection already exists. if it doesn't, create it.
    Set settings = connection.ExecQuery ("Select * From SMS_CollectionSettings Where CollectionID = '" & collectionID & "'")
   
    If settings.Count = 0 Then
        ' Wscript.Echo "Creating collection settings object"
        Set collectionSettings = connection.Get("SMS_CollectionSettings").SpawnInstance_
        collectionSettings.CollectionID = collectionId
        collectionSettings.Put_
    End If  
    
    ' Get the collection settings object.
    Set collectionSettings = connection.Get("SMS_CollectionSettings.CollectionID='" & collectionId &"'" )
   
    ' Get the collection variables.
    collectionVariables=collectionSettings.CollectionVariables
    
    ' Create and populate a new collection variable.
    Set collectionVariable = connection.Get("SMS_CollectionVariable").SpawnInstance_
    collectionVariable.Name = name
    collectionVariable.Value = value
    collectionVariable.IsMasked = mask
    
    ' Add the new collection variable.
    ReDim Preserve collectionVariables (UBound (collectionVariables)+1)
    Set collectionVariables(UBound(collectionVariables)) = collectionVariable
    
    collectionSettings.CollectionVariables=collectionVariables
    
    collectionSettings.Put_
    
 End Sub   


Sub DeleteCollection(connection, collectionIDToDelete)
    On Error Resume Next
    
    ' Get the specific collection instance to delete.
    Set collectionToDelete = connection.Get("SMS_Collection.CollectionID='" & collectionIDToDelete & "'")
    
    ' Delete the collection.
    collectionToDelete.Delete_
    
End Sub

Sub DeleteAdvertisement(connection, AdvertisementID )
    On Error Resume Next
    
    ' Get the specific collection instance to delete.
    Set Advertisement = connection.Get("SMS_Advertisement.AdvertisementID='" & AdvertisementID & "'")
    
    ' Delete the collection.
    Advertisement.Delete_
    
End Sub

Function GetResourceID(oSMSProvider,ComputerName)
	On Error Resume Next
	Dim Computers, oComputer 

	ResourceIDs = ""
	GetResourceID = ""
	Set Computers = oSMSProvider.ExecQuery ("select * from SMS_G_System_SYSTEM where name = '" & ComputerName & "'")
	If Computers.count = 0 Then
'		WriteLog "No such collection with name: " & CollectionName 
		Exit Function
	End If
	
	For Each oComputer in Computers
		ResourceIDs = ResourceIDs & "," & oComputer.resourceid 
		GetResourceID = oComputer.resourceid 
	Next
	
	GetResourceID = ResourceIDs
	
	Set Computers = Nothing

End Function 

Function GetCollectionID(oSMSProvider,CollectionName)
	On Error Resume Next
	Dim Collections, oCollection
	
	GetCollectionID = ""
	Set Collections = oSMSProvider.ExecQuery ("select * from SMS_Collection where name = '" & CollectionName & "'")
	If Collections.count = 0 Then
'		WriteLog "No such collection with name: " & CollectionName 
		Exit Function
	End If
	
	For Each oCollection in Collections
		GetCollectionID = oCollection.CollectionID
	Next
	
	Set Collections = Nothing

End Function 


Function GetSiteCodeForServer(SCCMServer)
	
	Select Case SCCMServer
		Case "IAUNSW456.AU.CBAINET.COM"
			GetSiteCodeForServer = "PRD"
		Case "IAUNSW457.AU.CBAINET.COM"
			GetSiteCodeForServer = "NSW"
		Case "IAUQLD064.AU.CBAINET.COM"
			GetSiteCodeForServer = "QLD"
		Case "IAUSA035.AU.CBAINET.COM"
			GetSiteCodeForServer = "SAU"
		Case "IAUWA040.AU.CBAINET.COM"
			GetSiteCodeForServer = "WAU"
		Case "IAUNSWT72.AUT01.CBAITEST01.COM"
			GetSiteCodeForServer = "PRE"
		Case "IAUNSWT73.AUT01.CBAITEST01.COM"
			GetSiteCodeForServer = "TST"
	End Select 

End Function 

Function CreateAdvertisement(connection, existingCollectionID, existingPackageID, existingProgramName, newAdvertisementName, newAdvertisementComment, newAdvertisementFlags, newAdvertisementStartOfferDateTime, newAdvertisementStartOfferEnabled)

    ' Create the new advertisement object.
    Set newAdvertisement = connection.Get("SMS_Advertisement").SpawnInstance_
    
    ' Populate the advertisement properties.
    newAdvertisement.CollectionID = existingCollectionID
    newAdvertisement.PackageID = existingPackageID
    newAdvertisement.ProgramName = existingProgramName
    newAdvertisement.AdvertisementName = newAdvertisementName
    newAdvertisement.Comment = newAdvertisementComment
    newAdvertisement.AdvertFlags = newAdvertisementFlags
    newAdvertisement.PresentTime = newAdvertisementStartOfferDateTime
    newAdvertisement.PresentTimeEnabled = newAdvertisementStartOfferEnabled
    newAdvertisement.RemoteClientFlags = "2088"
    ' Save the new advertisement and properties.
    newAdvertisement.Put_ 
    
End Function 

Function GetOSDAdvertisementID(oSMSProvider,AdvertisementName)
	On Error Resume Next
	Dim Advertisements, oAdvertisement
	
	GetOSDAdvertisementID = ""
	Set Advertisements = oSMSProvider.ExecQuery ("select * from SMS_Advertisement where AdvertisementName = '" & AdvertisementName & "'")
	If Advertisements.count = 0 Then
'		MsgBox "No such advertisement with name: " & AdvertisementName 
		Exit Function
	End If
	
	For Each oAdvertisement in Advertisements 
		GetOSDAdvertisementID = oAdvertisement.AdvertisementID
	Next
	
	Set Advertisements = Nothing

End Function 

Function GetOSDTaskSequenceID(oSMSProvider,TaskSequenceName)
	On Error Resume Next
	Dim TaskSequences, oTaskSequence 
	GetOSDTaskSequenceID = ""
	Set TaskSequences = oSMSProvider.ExecQuery ("select * from SMS_TaskSequencePackage where Name= '" & TaskSequenceName & "'")
	If TaskSequences.count = 0 Then
'		MsgBox "No such advertisement with name: " & AdvertisementName 
		Exit Function
	End If
	
	For Each oTaskSequence in TaskSequences
		GetOSDTaskSequenceID = oTaskSequence.PackageID
	Next
	
	Set TaskSequences = Nothing

End Function 


Function AddSchedTokenOneOffMandatory(oSMSProvider,AssignedSchedule, AdvertisementID)
	On Error Resume Next
	Dim AdvertArray(0), oAdvert 

	Set oAdvert=oSMSProvider.Get ("SMS_Advertisement.AdvertisementID='" & AdvertisementID & "'")
	Set SchedToken = oSMSProvider.Get("SMS_ST_NonRecurring").SpawnInstance_()
	SchedToken.StartTime = ConvertToWMIDate(AssignedSchedule) ' "20080101230000.000000+***"
	Set AdvertArray(0) = SchedToken
	oAdvert.AssignedSchedule = AdvertArray
	
	oAdvert.AssignedScheduleEnabled = True
	oAdvert.AssignedScheduleIsGMT = False
	oAdvert.Put_
	
	Set oAdvert = Nothing
	Set SchedToken = Nothing 

End Function 

Sub MoveAdvertisementToFolder (oSMSProvider,AdvertisementID, FolderName)
	On Error Resume Next

	Dim FolderID, oFolder, ret, aAdvertID
	Dim SourceFolderID : SourceFolderID = 0 ' Assume the root
	
	FolderID = getFolderID(oSMSProvider,FolderName)
	If FolderID = "" Then
'		MsgBox "Couldn't find Folder ID, nothing to do"
		Exit Sub
	End If

'	MsgBox "Moving AdvertisementID: " & AdvertisementID & " to FolderID: " & DestinationFolderID
	
	Set oFolder = oSMSProvider.Get("SMS_ObjectContainerItem")
	aAdvertID = Array(AdvertisementID)
	Err.Clear 
	ret = oFolder.MoveMembers (aAdvertID, SourceFolderID, FolderID, 3)
	If Err.Number <> 0 Then
'		MsgBox "Error moving advertisement to folder, err.number = " & Err.Number & ", err.description = " & Err.Description 
		Exit Sub
	End If

	If ret <> 0 Then
'		MsgBox "Error moving AdvertID: " & AdvertisementID & " to folder: " & DestinationFolderID & " ReturnCode=" & ret
	Else
'		MsgBox "Success moving AdvertID: " & AdvertisementID & " to folder: " & DestinationFolderID
	End If
	
	Set oFolder = Nothing

End Sub

Function GetFolderID(oSMSProvider,FolderName)
	On Error Resume Next
	
	Dim oContainers
	GetFolderID = ""
	Err.Clear 
	Set oContainers = oSMSProvider.ExecQuery("select * from SMS_ObjectContainerNode where ObjectType = 3 and Name = '" & FolderName & "'")
	If Err.number <> 0 Then
		MsgBox  "Error querying folder container node, err.number = " & Err.Number & ", err.description = " & Err.Description 
		Exit Function
	End If

	If IsNull(oContainers) Then
		MsgBox "Error: No folders returned: " & PkgiD
	End If

	For Each Container In oContainers
		GetFolderID = Container.ContainerNodeID 
	Next
	
	Set oContainers = Nothing
End Function


Sub CreateStaticCollection(connection, existingParentCollectionID, newCollectionName, newCollectionComment, ownedByThisSite, resourceClassName, ResourceIDs)

    ' Create the collection.
    Set newCollection = connection.Get("SMS_Collection").SpawnInstance_
    newCollection.Comment = newCollectionComment
    newCollection.Name = newCollectionName
    newCollection.OwnedByThisSite = ownedByThisSite
    
    ' Save the new collection and save the collection path for later.
    Set collectionPath = newCollection.Put_    
    
   ' Define to what collection the new collection is subordinate.
   ' IMPORTANT: If you do not specify the relationship, the new collection will not be visible in the console. 
    Set newSubCollectToSubCollect = connection.Get("SMS_CollectToSubCollect").SpawnInstance_
    newSubCollectToSubCollect.parentCollectionID = existingParentCollectionID
    newSubCollectToSubCollect.subCollectionID = CStr(collectionPath.Keys("CollectionID"))
    
    ' Save the subcollection information.
    newSubCollectToSubCollect.Put_

    ' Get the collection.
    Set newCollection = connection.Get(collectionPath.RelPath)
    
    aResourceID = Split(resourceIDs,",")
    For x = LBound(aResourceID) To UBound(aResourceID)
    	If Len(aResourceID(x)) > 0 then
		    ' Create the direct rule.
		    Set newDirectRule = connection.Get("SMS_CollectionRuleDirect").SpawnInstance_
		    newDirectRule.ResourceClassName = resourceClassName
		    newDirectRule.ResourceID = aResourceID(x)
		    
		    ' Add the new query rule to a variable.
		    Set newCollectionRule = newDirectRule
		    
		    ' Add the rules to the collection.
		    newCollection.AddMembershipRule newCollectionRule
			
		    ' Call RequestRefresh to initiate the collection evaluator. 
		    newCollection.RequestRefresh False
		    
		    Set newCollectionRule = Nothing
		    Set newDirectRule = Nothing
		End If
	Next   
End Sub

Sub filecontainsstringinlastlines(Node)
	On Error Resume Next
	LastLines = cint(Node.GetAttributeNode("lastlines").Value)
	searchstring = Node.GetAttributeNode("searchstring").Value
	filename = Node.GetAttributeNode("filename").Value
	Dim lines()
	ReDim lines(Lastlines)
	
	Const For_Reading = 1
	If InStr(FileName,"%COMPUTERNAME%") > 0 Then
		FileName = Replace(Filename,"%COMPUTERNAME%",ComputerName)
	End If
	
	' Check if text is already written to file
	If oFSO.FileExists(filename) = False Then
		WriteLog "filecontainsstringinlastlines",FileName,"failure","File does not exist"
		Exit sub
	End If 
	
	Err.Clear 
	Set oFile = oFSO.OpenTextFile(FileName,For_Reading,False)
	If err.Number <> 0 Then
		WriteLog "filecontainsstringinlastlines",FileName,"failure","Error opening file"
		Exit Sub
	End If 
	
	x = 0
	Do While oFile.AtEndOfStream <> True
		Lines(x) = oFile.ReadLine()
		x = x + 1
		If x = LastLines Then
			x = 0
		End If
	Loop 
	oFile.Close
	
	If x = 0 Then 
		x = LastLines - 1
	Else
		x = x - 1
	End If
	
	bFound = False
	For y = 0 To LastLines - 1
		If InStr(lines(y),searchstring) > 0 Then
			bFound = True
			exit for 
		End If 
	Next

	If bFound = True Then
		WriteLog "filecontainsstringinlastlines",FileName,"success",Left(lines(y),80)
		field1 = left(lines(y),50)
	Else
		WriteLog "filecontainsstringinlastlines",filename,"failure",searchstring & " is not in last " & LastLines & " lines." & left(lines(UBound(lines)-1),50)
	End If
	
	Set oFile = Nothing
End Sub

Sub filecontainsstringinlastlineserror(Node)
	On Error Resume Next
	LastLines = cint(Node.GetAttributeNode("lastlines").Value)
	searchstring = Node.GetAttributeNode("searchstring").Value
	filename = Node.GetAttributeNode("filename").Value
	Dim lines()
	ReDim lines(Lastlines)
	
	Const For_Reading = 1
	If InStr(FileName,"%COMPUTERNAME%") > 0 Then
		FileName = Replace(Filename,"%COMPUTERNAME%",ComputerName)
	End If
	
	' Check if text is already written to file
	If oFSO.FileExists(filename) = False Then
		WriteLog "filecontainsstringinlastlineserror",FileName,"failure","File does not exist"
		Exit sub
	End If 
	
	Err.Clear 
	Set oFile = oFSO.OpenTextFile(FileName,For_Reading,False)
	If err.Number <> 0 Then
		WriteLog "filecontainsstringinlastlineserror",FileName,"failure","Error opening file"
		Exit Sub
	End If 
	
	x = 0
	Do While oFile.AtEndOfStream <> True
		Lines(x) = oFile.ReadLine()
		x = x + 1
		If x = LastLines Then
			x = 0
		End If
	Loop 
	oFile.Close
	
	If x = 0 Then 
		x = LastLines - 1
	Else
		x = x - 1
	End If
	
	bFound = False
	For y = 0 To LastLines - 1
		If InStr(lines(y),searchstring) > 0 Then
			bFound = True
			exit for
		End If 
	Next

	If bFound = True Then
		WriteLog "filecontainsstringinlastlineserror",FileName,"failure",Left(lines(y),80)
		field1 = left(lines(y),30)
	Else
		WriteLog "filecontainsstringinlastlineserror",FileName,"success",searchstring & " is not in last " & lastLines & " lines."
	End If
	
	Set oFile = Nothing
End Sub

Sub prestagebdp(Source)
	On Error Resume Next
	

	if ofso.FolderExists ("\\" & ComputerName & "\f$\smspkgf$\") = true then
		Destination = "\\" & ComputerName & "\f$\smspkgf$\"
	elseif ofso.FolderExists ("\\" & ComputerName & "\e$\smspkge$\") = true then
		destination = "\\" & ComputerName & "\e$\smspkge$\"
	elseif ofso.FolderExists ("\\" & ComputerName & "\g$\smspkgg$\") = true then
		destination = "\\" & ComputerName & "\g$\smspkgg$\"
	else
		WriteLog "prestagebdp",Source,"failure","Server is not a BDP"
		exit sub 
	end if 

	If oFSO.FolderExists(Source) = False Then
		WriteLog "prestagebdp",Source,"failure","Folder does not exist"
		Exit Sub
	End If

	If oFSO.FolderExists(Destination) = False Then
		WriteLog "prestagebdp",Destination,"failure","Cannot access folder " & Destination
		Exit Sub
	End If

	Err.Clear
	oFSO.CopyFolder Source, Destination, True
	If Err.Number <> 0 Then
		WriteLog "prestagebdp",Source,"failure","Error copying folder to '" & Destination & "', Error = " & Err.Number
		Exit Sub
	End If

	WriteLog "prestagebdp",Source,"success",Destination
End Sub


