Option Explicit
'=================================================================================================================
'* Copyright 2010 Hewlett-Packard Development Company, L.P.
'*
'* Script Name: WINDOWS-OS-Stop_Services.vbs
'* Purpose: Windows Script to Stop the services on the machine
'* Notes: - This script is used to stop the parameterised services  on specific servers
'* 		  - /COMPUTERNAMES and /SERVICES are the mandatory parameters
'* 		  - The Script willbe stored centrally in MSSR and HPSA.
'* Usage 1: Cscript.exe WINDOWS-OS-Stop_Services.vbs /COMPUTERNAMES:computer1,computer2 /SERVICES:spooler,bits 
'* Usage 2: Cscript.exe WINDOWS-OS-Stop_Services.vbs /COMPUTERNAMES:computer2 /SERVICES:c:\temp\serviceslist.txt
'*
'* Modification Log:
'* ---Date---- Author Version Change-------------------------------------
'* 12-Sep-2014: Balaji Sethuraman 1.0 Created
'==================================================================================================================
'==================
'* Global constants
'==================
Const FOR_READING=1
Const FOR_WRITING=2
Const FOR_APPENDING=8
Const STRFOLDER = "C:\Scripts"
Const LOG_FILE = "WINDOWS-OS-Stop_Services.Log"
Const ECHOLOG = True
Const PAT_TXTFILE    = "^.*([tT][xX][tT]).*?$"
'==================
'* Global Variables
'==================
Dim g_strComputer
Dim g_objNetwork
Dim g_objFSO
Dim g_strDoLogFolder
Dim g_strlogfile
Dim g_objShell
Dim g_objArgsN
Dim g_StrComputerNames
Dim g_strServices

'==================
'* Global Objects
'==================
Set g_objNetwork = WScript.CreateObject("Wscript.network")
Set g_objFSO = CreateObject("Scripting.FileSystemObject")
Set g_objShell = CreateObject("WScript.Shell")
Set g_objArgsN = Wscript.Arguments.Named

'====================
'* Global Declaration
'====================
g_strDoLogFolder = STRFOLDER & "\Logs"
g_strlogfile = g_strDoLogFolder & "\" & LOG_FILE
g_strComputer = g_objNetwork.ComputerName
g_StrComputerNames = g_objArgsN("COMPUTERNAMES")
g_strServices = g_objArgsN("SERVICES")

Main

'=================================================
'* Name: Subroutine  Main
'* Purpose: This is the Core subroutine of Script
'* Input: N/A
'* Output: Completion of the Script 
'* Return: N/A
'=================================================
Sub Main
	Dim arrComputers
	Dim strComputer
	Dim objWMIService
	Dim colServiceList
	Dim strService
	Dim arrServices
	Dim objService
	Dim errReturn
	Dim arrOfServices
'	On Error Resume Next
		If Not g_objFSO.FolderExists(STRFOLDER) Then g_objFSO.CreateFolder(STRFOLDER)
		If Not g_objFSO.FolderExists(g_strDoLogFolder) Then g_objFSO.CreateFolder(g_strDoLogFolder)
		'If g_objFSO.FileExists(g_strlogfile) Then g_objFSO.DeleteFile(g_strlogfile)
		DoLog "INFO: Script Execution Started"
		DoLog "INFO: Script Executing from the server: " & g_objNetwork.ComputerName
		DoLog "INFO: User:" & g_objNetwork.UserName
		fnParseCommandlineParameters()
			arrComputers = Split(g_StrComputerNames,",")
			For Each strComputer In arrComputers
				DoLog "INFO: Connecting to " & strComputer
				Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
				
				arrServices = Split(g_strServices,",")
				For Each strService In arrServices
					Set colServiceList = objWMIService.ExecQuery("Select * from Win32_Service where DisplayName='"& strService &"'")
						DoLog "INFO: Stopping Service: " & strService
						For Each objService in colServiceList
						    errReturn = objService.StopService()
						    Select Case errReturn
						    	Case 0 DoLog "SUCCESS: Service " & strService & " Stopped"
						    	Case 1 DoLog "FAILED: The request is not supported":reportError
						    	Case 2 DoLog "FAILED: The user did not have the necessary access.":reportError
						    	Case 3 
						    		   DoLog "FAILED: The service cannot be stopped because other services that are running are dependent on it"
						    		   DoLog "Dependant Services: " & fnGetDependantServices(objWMIService, objService.Name)
						    		   reportError
						    End Select 
						Next
				Next 
			Next 
		DoLog "INFO: Script Completed"
		DoLog "---------------------------------------------------"
End Sub  
'============================================
'* Name: Subroutine  DoLog
'* Purpose: TO write log files
'* Input: Text Entry
'* Output: Echos the log when ECHOLOG = True 
'* Return: Write logs to the file
'============================================
Sub DoLog(strEntry)
	Dim objDoLogFile
		On Error Goto 0
		Err.Clear
		If Not g_objFSO.FileExists(g_strlogfile) Then
				set objDoLogFile= g_objFSO.OpenTextFile(g_strlogfile, FOR_WRITING, True)
		Else
				Set objDoLogFile= g_objFSO.OpenTextFile(g_strlogfile, FOR_APPENDING, True)
		End If
		If ECHOLOG = True Then WScript.Echo Now & ": " & strEntry
		objDoLogFile.WriteLine Now & ": " & strEntry
		objDoLogFile.Close
		Set objDoLogFile= Nothing
End Sub  
'==================================
'* Name: Function  reportError
'* Purpose: To log Error and quit
'* Input: N/A
'* Output: N/A
'* Return: Quits the script
'==================================
Function reportError
	 Dolog "ERROR: Script interrupted due to error"
	 Err.Clear
	 WScript.Quit
End Function
'==========================================================
'* Name: Function  fnGetDependantServices
'* Purpose: To build and return array of Service names
'* Input: WMI Connection and Display Name of Service
'* Output: N/A
'* Return: Array of Dependant Names seperated by Comma ","
'==========================================================
Function fnGetDependantServices (ByVal objWMI,ByVal Service)
	Dim colDependantServiceList
	Dim arrDependantNames()
	Dim objDepenantService, j
	j = 0
		Set colDependantServiceList = objWMI.ExecQuery("Associators of " _
		    & "{Win32_Service.Name='"& Service &"'} Where " _
		        & "AssocClass=Win32_DependentService " & "Role=Antecedent" )
		For Each objDepenantService In colDependantServiceList
			Redim Preserve arrDependantNames(j)
			arrDependantNames(j) = objDepenantService.DisplayName
			j = j + 1
		Next 
		fnGetDependantServices = Join(arrDependantNames,",")
End Function
'=======================================================================
'* Name: Function TestPattern
'* Purpose: Take a string and a regular expression patern as parameters
'*          return true is sString match sPatern, otherwise return false
'* Input: string and a regular expression patern
'* Output: N/A
'* Return: true is sString match sPatern, otherwise return false
'=======================================================================
Function TestPattern(sString, sPatern)
   Dim oRegEx                            'Object declaration
   Set oRegEx = New RegExp               'Create the RegExp object
   oRegEx.Pattern = sPatern 'Assign the patern to the RegExp object
   oRegEx.IgnoreCase = True              'This function always ignore the case
   TestPattern = oRegEx.Test(sString)    'Directly assign the function result
                                         'based on the test
   Set oRegEx = Nothing                  'Clean the object
End Function
'===================================================
'* Name: Function  fnGetAllServicesTxtFile
'* Purpose: To build array of names from text file
'* Input: Text file
'* Output: N/A
'* Return: Array of Names seperated by Comma ","
'===================================================
Function fnGetAllServicesTxtFile(ByVal InputFile)
Dim ObjInputFile
Dim arrServiceTxt(), i
i = 0
	Set ObjInputFile = g_objFSO.OpenTextFile(InputFile,FOR_READING,True)
		Do Until ObjInputFile.AtEndOfStream
			Redim Preserve arrServiceTxt(i)
			arrServiceTxt(i) = ObjInputFile.ReadLine
			i = i + 1
		Loop
fnGetAllServicesTxtFile = Join(arrServiceTxt,",")
End Function 
'===============================================
'* Name: Function  fnParseCommandlineParameters
'* Purpose: To Parse command line parameters
'* Input: N/A
'* Output: writes logs on obtained paramters
'* Return: N/A
'===============================================
Function fnParseCommandlineParameters
Dim bCmdCheck : bCmdCheck = True
Dim arrOfServices
		' Check if help files are called
		If g_objArgsN.Exists("?") Then
			Dolog "Switch /? used, Calling Help Files"
			Call Help
		End If
		
		If g_objArgsN.Exists("COMPUTERNAMES") And (g_StrComputerNames <> "") Then
			Call DoLog("INFO: Received parameters: /COMPUTERNAMES:<" & g_StrComputerNames &">")
		Else
			Call DoLog("ERROR: Invalid parameters received: /COMPUTERNAMES:<" & g_StrComputerNames &">") 
			bCmdCheck = False
		End If 
		
		' Check if Txt file is specified if exist parse services array from text file
		If TestPattern(g_objArgsN("SERVICES"),PAT_TXTFILE) Then
			Call DoLog("INFO: Received valid parameters: /SERVICES:<" & g_strServices &">")
			arrOfServices = fnGetAllServicesTxtFile(g_strServices)
			g_strServices = ""
			g_strServices = arrOfServices
		Else 
			If g_objArgsN.Exists("SERVICES") And g_strServices <> "" Then
				Call DoLog("INFO: Received parameters: /SERVICES:<" & g_strServices &">")
			Else 
				Call DoLog("ERROR: Invalid parameters received: /SERVICES:<" & g_strServices &">")
				bCmdCheck = False 
			End If 
		End If
		
		If Not bCmdCheck Then
			DoLog("ERROR: Atleast one switch or switch value is incorrect")
			Call Help 
		End If 

End Function 
'=======================================
'* Name: Subroutine HELP
'* Purpose: To Display the Help files
'* Input: Script with /? as switch
'* Output: Displays the Help Files
'* Return: Help files
'=======================================
Sub Help
	Dolog "Running Help Files."
	Wscript.Echo "****************************************************************************************"
	Wscript.Echo "* NAME: WINDOWS-OS-Stop_Services.vbs                                                   *"
	Wscript.Echo "* PURPOSE: Windows Script to Stop the services on the machine                          *"
	Wscript.Echo "* SWITCHES:  /COMPUTERNAMES:<computer1,computer2> - Gets an Array of Computers         *"
	Wscript.Echo "*            /SERVICES:<service1,service2> - Gets an Array of Services to Stop         *"
	Wscript.Echo "*            /SERVICES:c:\temp\Input.txt - Gets services from input file, line by line *"
	Wscript.Echo "* USAGE1 :<Script Name>.vbs /COMPUTERNAMES:computer1,computer2 /SERVICES:spooler,bits  *"
	Wscript.Echo "* USAGE1 :<Script Name>.vbs /COMPUTERNAMES:computer2 /SERVICES:c:\temp\serviceslist.txt*"
	Wscript.Echo "* HELP   :<Script Name>.vbs /?                                                         *"
	Wscript.Echo "****************************************************************************************"
	WScript.Quit
End Sub
