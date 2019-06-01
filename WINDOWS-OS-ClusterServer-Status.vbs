'=========================================================================================================
'NAME:        SVAC_WINDOWS-OS-ClusterServer-Status.vbs
'
'DESCRIPTION: This script will first determines if a server is on a cluster. If so, 
'             it'll check the state of various cluster related objects, including: nodes
'             resource groups, resources, networks and network interfaces.
'INPUT:       None
'OUTPUT:      Output Echo to HPSA
'
'MODIFICATION LOG: 
'04/19/2010	 Garry Xu 			Created.
'06/09/2010  Dean Tondreau      Modified:  Removed OS caption & trim messages for easier readability.
'09/30/2010  Dean Tondreau		Rewrote code to report only issues & resolve "state may not exist" return
'==========================================================================================================
Option Explicit

'==============================================
' VARIABLE DEFINITIONS
'==============================================
Dim objWMIService 
Dim colOperatingSystems
Dim objOS
Dim colRunningServices
Dim objService
Dim colItems
Dim objItem
Dim nItems
Dim strComputerName
Dim stateValue

strComputerName = "."

'==============================================
' DETERMINE CLUSTER SERVICE ENALBED ON SERVER
'==============================================
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputerName & "\root\cimv2")
Set colRunningServices = objWMIService.ExecQuery("SELECT * FROM Win32_Service WHERE DisplayName = 'Cluster Service'")
nItems = colRunningServices.Count

If nItems > 0 Then
	Call Verify_Cluster_Status()
Else
	Wscript.Echo "Cluster Status Check: Skipped"
End If

WScript.Quit(0)

'==============================================
' INITIATES CLUSTER STATE VERIFICATION
'==============================================
Function Verify_Cluster_Status()
	On Error Resume Next
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\mscluster")
	Set colItems = objWMIService.ExecQuery("SELECT * FROM MSCluster_Node")
	If Err.Number <> 0 Then
		WScript.Echo "Cluster Status Check: Unable to retreive Cluster Node Status"
	Else
		For Each objItem in colItems
			stateValue = objItem.State
			SELECT CASE stateValue
				CASE 0
					Wscript.Echo "Cluster Status Check: " & objItem.Name & " || Node State: OK"
				CASE 1	
					Wscript.Echo "Cluster Status Check: " & objItem.Name & " || Node State: DOWN"
				CASE 2	
					Wscript.Echo "Cluster Status Check: " & objItem.Name & " || Node State: PAUSED"
				CASE 3
					Wscript.Echo "Cluster Status Check: " & objItem.Name & " || Node State: JOINING"
				CASE Else
					WScript.Echo "Cluster Status Check: " & objItem.Name & " || Node State: UNKNOWN"
			END SELECT
		Next
	End If

	Set colItems = objWMIService.ExecQuery("SELECT * FROM MSCluster_ResourceGroup")
	If Err.Number <> 0 Then
		Wscript.Echo "Cluster Status Check: Unable to retreive Cluster Resource Group Status"
	Else
		For Each objItem in colItems
			stateValue = objItem.State
			SELECT CASE stateValue
				CASE 0
					WScript.Echo "Cluster Status Check: " & objItem.Name & " || Resource Group Status = Online"
				CASE 1	
					Wscript.Echo "Cluster Status Check: " & objItem.Name & " || Resource Group Status = Offline"
				CASE 2
					Wscript.Echo "Cluster Status Check: " & objItem.Name & " || Resource Group Status = Failed"
				CASE 3
					Wscript.Echo "Cluster Status Check: " & objItem.Name & " || Resource Group Status = Partial Online"
				CASE 4	
					Wscript.Echo "Cluster Status Check: " & objItem.Name & " || Resource Group Status = Pending"
				CASE Else
					Wscript.Echo "Cluster Status Check: " & objItem.Name & " || Resource Group Status = UNKNOWN"
			END SELECT
		Next
	End If
	
	Set colItems = objWMIService.ExecQuery("SELECT * FROM MSCluster_Resource")
	If Err.Number <> 0 Then
		Wscript.Echo "Cluster Status Check: Unable to retreive Cluster Resource Status"
	Else
		For Each objItem in colItems
			stateValue = objItem.State
			SELECT CASE stateValue
				CASE 1
					Wscript.Echo "Cluster Status Check: " & objItem.Name & " || Resource Status = Initializing"
				CASE 2
					WScript.Echo "Cluster Status Check: " & objItem.Name & " || Resource Status = Online"
				CASE 3	
					Wscript.Echo "Cluster Status Check: " & objItem.Name & " || Resource Status = Offline"
				CASE 4
					Wscript.Echo "Cluster Status Check: " & objItem.Name & " || Resource Status = Failed"
				CASE 128
					Wscript.Echo "Cluster Status Check: " & objItem.Name & " || Resource Status = Pending"
				CASE 129	
					Wscript.Echo "Cluster Status Check: " & objItem.Name & " || Resource Status = Online Pending"
				CASE 130
					Wscript.Echo "Cluster Status Check: " & objItem.Name & " || Resource Status = Offline Pending"
				CASE Else
					Wscript.Echo "Cluster Status Check: " & objItem.Name & " || Resource Status = UNKNOWN"
			END SELECT
		Next
	End If
	
	Set colItems = objWMIService.ExecQuery("SELECT * FROM MSCluster_Network")
	If Err.Number <> 0 Then
		Wscript.Echo "Cluster Status Check: Unable to retreive Cluster Network Status"
	Else
		For Each objItem in colItems
			stateValue = objItem.State
			SELECT CASE stateValue
				CASE 0
					Wscript.Echo "Cluster Status Check: " & objItem.Name & " || Network Status = Unavailable"
				CASE 1
					Wscript.Echo "Cluster Status Check: " & objItem.Name & " || Network Status = Down"
				CASE 2
					Wscript.Echo "Cluster Status Check: " & objItem.Name & " || Network Status = Partitioned"				
				CASE 3	
					Wscript.Echo "Cluster Status Check: " & objItem.Name & " || Network Status = Up"
				CASE Else	
					Wscript.Echo "Cluster Status Check: " & objItem.Name & " || Network Status = UNKNOWN"
			END SELECT
		Next
	End If
	
	Set colItems = objWMIService.ExecQuery("SELECT * FROM MSCluster_NetworkInterface")
	If Err.Number <> 0 Then
		Wscript.Echo "Cluster Status Check: Unable to retreive Cluster Network Interface Status"
	Else
		For Each objItem in colItems
			stateValue = objItem.State
			SELECT CASE stateValue
				CASE 0
					Wscript.Echo "Cluster Status Check: " & objItem.Name & " || Network Interface Status = Unavailable"
				CASE 1
					Wscript.Echo "Cluster Status Check: " & objItem.Name & " || Network Interface Status = Failed"
				CASE 2
					Wscript.Echo "Cluster Status Check: " & objItem.Name & " || Network Interface Status = Unreachable"	
				CASE 3		
					Wscript.Echo "Cluster Status Check: " & objItem.Name & " || Network Interface Status = Up"
				CASE Else	
					Wscript.Echo "Cluster Status Check: " & objItem.Name & " || Network Interface Status = UNKNOWN"
			END SELECT
		Next
	End If
End Function
