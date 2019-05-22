Option Explicit

' NAME: WINDOWS-OS-Disk_Cleanup-1381.vbs
' PURPOSE: Move orphaned files
'          Move or clear temp files
'          Delete user PA* profiles
'          Delete log files
'
' INPUTS: [/SERVERLIST:path_to_list] [/IIS:iis_age] [/OTHERS:others_age][/HELP]
' PARAMETERS:
'   path_to_list - optional, file containing list of servers to process
'                - default, look at current folder for serverlist.txt
'   iis_age      - optional, file age of IIS logs in days to check
'                - default, 30 days
'   others_age   - optional, file age of all logs in days to check
'                - default, 360 days
'   /HELP        - display this help
'
' USAGE:
'   < script >
'     - Run against serverlist.txt
'   < script > /IIS:40
'     - Change default 30 days for IIS logs
'   < script > /SERVERLIST:"c:\temp\serverlist.txt"
'     - Run against c:\temp\serverlist.txt
'   < script >
'     - Display this help screen"
'
' NOTE:
'   Increase HPSA script output to 9999 KB
'
' AUTHOR(s): Co, Harris <harris.co@hpe.com>
' DATE WRITTEN:  Jul 2015
' MODIFICATION HISTORY:  Jul 2015 - initial release
'

REM <FromDriveC_7212015>

REM >> WINDOWS 2003 System Drive
REM Move orphaned files (files in C:\ without folders) and installer files/folders to another drive and name it <FromDriveC>
REM Move or clear Temp/Tempo/TMP/Software/Drivers/Install folders
REM Delete User Profiles in C:\Documents and Settings older than 2015
REM     Only starting with PA should be deleted
REM Delete log files older than 2015 (search for *.log in C:\)
REM Delete memory dump files and old files (search for *.dmp and *.old in C:\)
REM Delete IIS log files older than 30 days (C:\WINDOWS\system32\LogFiles\W3SVC1  | HTTPERR)

REM >> Windows 2008 System Drive
REM Move orphaned files (files in C:\ without folders) and installer files/folders to another drive and name it <FromDriveC>
REM Move or clear Temp/Tempo/TMP/Software/Drivers/Install folders
REM Delete User Profiles in C:\Users
REM     Only starting with PA should be deleted
REM Delete log files older than 2015 (search for *.log in C:\)
REM Delete memory dump files and old files (search for *.dmp and *.old in C:\)
REM Delete IIS log files older than 30 days (C:\inetpub\logs\LogFiles\W3SVC1 | HTTPERR)

REM 1.  Move orphaned files (files in C:\ without folders) and installer files/folders to another drive and name it <FromDriveC>
REM     a.  Which drive? C
REM     b.  Location of installer files/folders? C:/Software or Softwares; C:/Install
REM 2.  Move or clear Temp/Tempo/TMP/Software/Drivers/Install folders
REM     Delete User Profiles in C:\Documents and Settings older than 2015
REM     a.  Criteria for move or clear?  How to determine if it is a move or clear? If the other drive can still accommodate the size, then it’s ‘move’ but if no other drive to transfer it to, then it should be cleared
REM     b.  Specify folder for Temp (is it system temp?) Same name. Move them all to a folder like <FromDriveC>
REM     c.  Specify folder for Tempo
REM     d.  Specify folder for TMP (is it system tmp?)
REM     e.  Specify folder for Software
REM     f.  Specify folder for Drivers
REM     g.  Specify folder for Install Folders

REM ***The folders of Item bcdefg are all located under Drive C:

On Error Resume Next

Const cReturnCode = 0
Const TIMEOUT_MINUTES = 30

Dim ParamSERVERLIST
Dim ParamOLDIIS
Dim ParamOLDOTHER

Dim DiagnosedFlag:      DiagnosedFlag   = False

'containers
Dim ServerList:         Set ServerList  = CreateObject("Scripting.Dictionary")

Dim TotalServersInFile
Dim TotalServersProcessed:  TotalServersProcessed = 0

Dim startTime
Dim BACKUPFOLDERNAME


DisplayHeader

CheckParameter

Run

DoExit

Private Sub Run
Dim DriveLetter
Dim server

    On Error Resume Next

    For Each server In ServerList.Keys
        WScript.Echo "Processing " & server
        DriveLetter = GetDriveWithHighestFreeSpace(server)
        If CreateBackupFolder(server, DriveLetter) Then

            If RemoteFolderExists(server, DriveLetter) Then
                WScript.Echo Space(4) & "Backup folder successfully created."
                MoveOrphanedFiles server
                MoveTEMP server
                DeleteDMPOLDFiles server, ParamOLDOTHER
                DeleteLOGFiles server, ParamOLDOTHER
                DeleteIISLogFiles server, ParamOLDIIS
                DeleteOldProfiles server, ParamOLDOTHER

                TotalServersProcessed = TotalServersProcessed + 1
            End If
        End If

        WScript.Echo
        WScript.Echo
    Next
End Sub

Private Function GetServerListFromFile
Const FORREADING = 1
Const ADVARCHAR  = 200
Const MAXCHARACTERS = 255

Dim obj, item
Dim temp

    Set temp = CreateObject("ADOR.RecordSet")
    temp.Fields.Append "Item", ADVARCHAR, MAXCHARACTERS
    temp.Open

    TotalServersInFile = 1

    If Not FileExists(ParamSERVERLIST) Then
        WScript.Echo ParamSERVERLIST & " does not exists."
    Else
        Set obj = CreateObject("Scripting.FileSystemObject").OpenTextFile(ParamSERVERLIST, FORREADING)
        Do While Not obj.AtEndOfStream
            item = Trim(obj.ReadLine)
            If item <> "" Then
                temp.AddNew
                temp("Item") = UCase(item)
                temp.Update
            End If
        Loop

        obj.Close
        Set obj = Nothing
    End If

    If Err.Number <> 0 Then
        WScript.Echo ParamSERVERLIST & " cannot be read"
        DiagnosedFlag = True
    Else
        TotalServersInFile = temp.RecordCount
        WScript.Echo ParamSERVERLIST & " contains " & TotalServersInFile & " servers"
        If TotalServersInFile = 0 Then
            WScript.Echo "Please list down server names in " & ParamSERVERLIST
            DiagnosedFlag = True
        End If
    End If

    Sort temp, ServerList

    temp.Close
    Set temp = Nothing

    Err.Clear
End Function

Private Sub Sort(source, ByRef list)
Dim item
Dim bit

    source.Sort = "Item"
    source.MoveFirst

    Do Until source.EOF
        item = source("Item")
        If Not list.Exists(item) Then
            list.Add item, ""
        End If
        source.MoveNext
    Loop
End Sub

Private Function FileExists(file)
    FileExists = CreateObject("Scripting.FileSystemObject").FileExists(file)
End Function

Private Sub DisplayHeader
    WScript.Echo "RBA script stdout"
    WScript.Echo "WFAN="""
    startTime = Now
End Sub

'-------------------------
Private Sub MoveOrphanedFiles(server)
    RunRemoteProcess server, "MOVE /Y C:\* " & BACKUPFOLDERNAME
End Sub

Private Sub MoveTEMP(server)
    'Temp/Tempo/TMP/Software/Drivers/Install
    If RunRemoteProcess(server, "MOVE /Y C:\Drivers\* " & BACKUPFOLDERNAME) Then
        RunRemoteProcess server, "DEL /F /Q C:\Drivers\* "
    End If

    If Not RunRemoteProcess(server, "MOVE /Y C:\Software\* " & BACKUPFOLDERNAME) Then
        RunRemoteProcess server, "DEL /F /Q C:\Software\* "
    End If

    If Not RunRemoteProcess(server, "MOVE /Y " & """" & "C:\Install Folders\* " & """" & BACKUPFOLDERNAME) Then
        RunRemoteProcess server, "DEL /F /Q " & """" & "C:\Install Folders\* " & """"
    End If

    If Not RunRemoteProcess(server, "MOVE /Y C:\temp\* " & BACKUPFOLDERNAME) Then
        RunRemoteProcess server, "DEL /F /Q C:\temp\* "
        'RunRemoteProcess server, "RD /S /Q C:\temp\* "
        'RunRemoteProcess server, "MKDIR C:\temp "
    End If

    If Not RunRemoteProcess(server, "MOVE /Y C:\Tempo\* " & BACKUPFOLDERNAME) Then
        RunRemoteProcess server, "DEL /F /Q C:\Tempo\* "
        'RunRemoteProcess server, "RD /S /Q C:\Tempo\* "
    End If

    If Not RunRemoteProcess(server, "MOVE /Y C:\TMP\* " & BACKUPFOLDERNAME) Then
        RunRemoteProcess server, "DEL /F /Q C:\TMP\* "
    End If
End Sub

Private Function CreateBackupFolder(server, drive)
    WScript.Echo Space(2) & "Creating Backup Folder"
    BACKUPFOLDERNAME = drive & ":\FromDriveC_" & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & Year(Now)
    CreateBackupFolder = RunRemoteProcess(server, "MKDIR " & BACKUPFOLDERNAME)
End Function

Private Function RemoteFolderExists(server, drive)
Dim objWMIService
Dim colItems
Dim OK

    On Error Resume Next

    OK = False
    WScript.Echo Space(2) & "Verifying Backup Folder"

    Set objWMIService = GetObject("winmgmts:\\" & server & "\root\cimv2")
    If Err.Number <> 0 Then
        WScript.Echo Space(4) & "Error: " & Err.Description
    Else
        Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Directory where drive='" & drive & ":" & "' and name='" & Replace(BACKUPFOLDERNAME, "\", "\\") & "'  ")

        OK = (colItems.Count <> 0)
    End If

    RemoteFolderExists = OK
End Function

Private Sub DeleteDMPOLDFiles(server, daysOLD)
Dim ModifiedDate
Dim objWMIService
Dim colItems
Dim objItem
Dim cmd
Dim COUNT

    On Error Resume Next

    If daysOLD = "" Then
        daysOLD = 360
    End If

    WScript.Echo Space(2) & "Executing DeleteDMPOLDFiles older than " & daysOLD & " days"

    Set objWMIService = GetObject("winmgmts:\\" & server & "\root\cimv2")
    If Err.Number <> 0 Then
        WScript.Echo Space(4) & "Error: " & Err.Description
    Else
        ModifiedDate = DateAdd("d", -daysOLD, Now)
        'format is YYYYMMDDHHMMSS.000000+000
        ModifiedDate = Year(ModifiedDate) & AddZero(Month(ModifiedDate)) & AddZero(Day(ModifiedDate)) & "000000.000000+000"

        Set colItems = objWMIService.ExecQuery("SELECT Name, LastModified, FileSize FROM CIM_DataFile " & _
                                 "WHERE (Extension = 'old' OR Extension = 'dmp') AND LastModified <= '" & ModifiedDate & "' AND " & _
                                 "Drive='C:'")

        If colItems.Count > 0 Then
            COUNT = colItems.Count
            For Each objItem In colItems
                cmd = "DEL /F /Q " & """" & objItem.Name & """"
                RunRemoteProcess server, cmd
            Next

            'Check
            Set colItems = objWMIService.ExecQuery("SELECT Name, LastModified, FileSize FROM CIM_DataFile " & _
                                 "WHERE (Extension = 'old' OR Extension = 'dmp') AND LastModified <= '" & ModifiedDate & "' AND " & _
                                 "Drive='C:'")
            WScript.Echo
            WScript.Echo Space(4) & "Summary:"
            WScript.Echo Space(6) & "Found: " & COUNT
            WScript.Echo Space(6) & "Not Deleted: " & colItems.Count
            If COUNT = colItems.Count Then
                WScript.Echo Space(4) & "WARNING: No files deleted, please check permissions"
            End If
            WScript.Echo
        Else
            WScript.Echo Space(4) & colItems.Count & " items."
        End If
    End If

    Err.Clear
End Sub

Private Sub DeleteLOGFiles(server, daysOLD)
Dim ModifiedDate
Dim objWMIService
Dim colItems
Dim objItem
Dim cmd
Dim COUNT

    On Error Resume Next

    If daysOLD = "" Then
        daysOLD = 360
    End If

    WScript.Echo Space(2) & "Executing DeleteLOGFiles older than " & daysOLD & " days"

    Set objWMIService = GetObject("winmgmts:\\" & server & "\root\cimv2")
    If Err.Number <> 0 Then
        WScript.Echo Space(4) & "Error: " & Err.Description
    Else
        ModifiedDate = DateAdd("d", -daysOLD, Now)
        'format is YYYYMMDDHHMMSS.000000+000
        ModifiedDate = Year(ModifiedDate) & AddZero(Month(ModifiedDate)) & AddZero(Day(ModifiedDate)) & "000000.000000+000"

        Set colItems = objWMIService.ExecQuery("SELECT Name, LastModified, FileSize FROM CIM_DataFile " & _
                                 "WHERE Extension = 'log' AND LastModified <= '" & ModifiedDate & "' AND " & _
                                 "Drive='C:'")

        If colItems.Count > 0 Then
            COUNT = colItems.Count
            For Each objItem In colItems
                cmd = "DEL /F /Q " & """" & objItem.Name & """"
                RunRemoteProcess server, cmd
            Next

            'Check
            Set colItems = objWMIService.ExecQuery("SELECT Name, LastModified, FileSize FROM CIM_DataFile " & _
                                 "WHERE Extension = 'log' AND LastModified <= '" & ModifiedDate & "' AND " & _
                                 "Drive='C:'")
            WScript.Echo
            WScript.Echo Space(4) & "Summary:"
            WScript.Echo Space(6) & "Found: " & COUNT
            WScript.Echo Space(6) & "Not Deleted: " & colItems.Count
            If COUNT = colItems.Count Then
                WScript.Echo Space(4) & "WARNING: No files deleted, please check permissions"
            End If
            WScript.Echo
        Else
            WScript.Echo Space(4) & colItems.Count & " items."
        End If
    End If

    Err.Clear
End Sub

Private Sub DeleteIISLogFiles(server, daysOLD)
Dim ModifiedDate
Dim objWMIService
Dim colItems
Dim objItem
Dim cmd
Dim COUNT

    On Error Resume Next

    If daysOLD = "" Then
        daysOLD = 30
    End If

    ModifiedDate = DateAdd("d", -daysOLD, Now)
    ModifiedDate = Year(ModifiedDate) & AddZero(Month(ModifiedDate)) & AddZero(Day(ModifiedDate))& "000000.000000+000"

    WScript.Echo Space(2) & "Executing DeleteIISLogFiles older than " & daysOLD & " days"

    Set objWMIService = GetObject("winmgmts:\\" & server & "\root\cimv2")
    If Err.Number <> 0 Then
        WScript.Echo Space(4) & "Error: " & Err.Description
    Else
        Set colItems = objWMIService.ExecQuery("SELECT * FROM CIM_DataFile " & _
                       "WHERE Drive='C:' And (" & _
                       "Path='\\inetpub\\logs\\LogFiles\\W3SVC1\\' OR " & _
                       "Path='\\inetpub\\logs\\LogFiles\\HTTPERR\\' OR " & _
                       "Path='\\WINDOWS\\system32\\LogFiles\\W3SVC1\\' OR " & _
                       "Path='\\WINDOWS\\system32\\LogFiles\\HTTPERR\\') AND " & _
                       "LastModified <= '" & ModifiedDate & "'")

        If colItems.Count > 0 Then
            COUNT = colItems.Count
            For Each objItem In colItems
                cmd = "DEL /F /Q " & """" & objItem.Name & """"
                RunRemoteProcess server, cmd
            Next

            'Check
            Set colItems = objWMIService.ExecQuery("SELECT * FROM CIM_DataFile " & _
                       "WHERE Drive='C:' And (" & _
                       "Path='\\inetpub\\logs\\LogFiles\\W3SVC1\\' OR " & _
                       "Path='\\inetpub\\logs\\LogFiles\\HTTPERR\\' OR " & _
                       "Path='\\WINDOWS\\system32\\LogFiles\\W3SVC1\\' OR " & _
                       "Path='\\WINDOWS\\system32\\LogFiles\\HTTPERR\\') AND " & _
                       "LastModified <= '" & ModifiedDate & "'")
            WScript.Echo
            WScript.Echo Space(4) & "Summary:"
            WScript.Echo Space(6) & "Found: " & COUNT
            WScript.Echo Space(6) & "Not Deleted: " & colItems.Count
            If COUNT = colItems.Count Then
                WScript.Echo Space(4) & "WARNING: No files deleted, please check permissions"
            End If
            WScript.Echo
        Else
            WScript.Echo Space(4) & colItems.Count & " items."
        End If
    End If

    Err.Clear
End Sub

Private Sub DeleteOldProfiles(server, daysOLD)
Dim objWMIService
Dim colItems
Dim objItem
Dim cmd
Dim COUNT

    On Error Resume Next

    If daysOLD = "" Then
        daysOLD = 360
    End If

    WScript.Echo Space(2) & "Executing DeleteOldProfiles older than " & daysOLD & " days"

    Set objWMIService = GetObject("winmgmts:\\" & server & "\root\cimv2")
    If Err.Number <> 0 Then
        WScript.Echo Space(4) & "Error: " & Err.Description
    Else
        Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_UserProfile WHERE SID LIKE 'S-1-5-21%' ")

        For Each objItem In colItems
            COUNT = colItems.Count
            'Delete only \PA* accounts
            If InStr(1, str, "\PA", vbTextCompare) <> 0 Then
                If DateDiff("d", WMIDateStringToDate(objItem.LastUseTime), Now) >= daysOLD Then
                    cmd = "RD /S /Q " & """" & objItem.LocalPath & """"
                    DeleteOldProfiles = RunRemoteProcess(server, cmd)
                Else
                    WScript.Echo Space(4) & "Skipping " & objItem.LocalPath
                    WScript.Echo Space(6) & DateDiff("d", WMIDateStringToDate(objItem.LastUseTime), Now) & " days old (within threshold)"
                End If
            Else
                WScript.Echo Space(4) & "Skipping " & objItem.LocalPath
                WScript.Echo Space(6) & "Non privileged account"
            End If

            'Check
            Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_UserProfile WHERE SID LIKE 'S-1-5-21%' ")
            WScript.Echo
            WScript.Echo Space(4) & "Summary:"
            WScript.Echo Space(6) & "Found: " & COUNT
            WScript.Echo Space(6) & "Not Deleted: " & colItems.Count
            If COUNT = colItems.Count Then
                WScript.Echo Space(4) & "WARNING: No profiles deleted, please check permissions"
            End If
            WScript.Echo
        Next
    End If

    Err.Clear
End Sub

Private Function WMIDateStringToDate(dtmDate)
    WMIDateStringToDate = CDate(Mid(dtmDate, 5, 2) & "/" & _
                          Mid(dtmDate, 7, 2) & "/" & Left(dtmDate, 4) _
                          & " " & Mid (dtmDate, 9, 2) & ":" & Mid(dtmDate, 11, 2) & ":" & Mid(dtmDate,13, 2))
End Function

Private Function RunRemoteProcess(server, cmd)
Dim objWMI
Dim objProcess
Dim errReturn
Dim intProcessID
Dim start_time
Dim colProcesses
Dim OK

    On Error Resume Next
    OK = False

    cmd = "cmd /c " & cmd
    WScript.Echo Space(4) & "Executing " & cmd

    Set objWMI = GetObject("winmgmts:\\" & server & "\root\cimv2")
    If Err.Number <> 0 Then
        WScript.Echo Space(6) & "Error: " & Err.Description
    Else
        Set objProcess = objWMI.Get("Win32_Process")
        errReturn = objProcess.Create(cmd, null, null, intProcessID)

        Select Case errReturn
            Case 0 'Successful Completion
                OK = True
                start_time = Now
                WScript.Echo Space(6) & "Running..."
                Do While True
                    WScript.Sleep 2000
                    Set colProcesses = objWMI.ExecQuery("Select * from Win32_Process Where ProcessID = " & intProcessID)
                    If colProcesses.Count = 0 Then Exit Do
                    If DateDiff("n", start_time, Now) > TIMEOUT_MINUTES Then
                        colProcesses.Terminate()
                        WScript.Echo Space(6) & "Warning: Timed-out"
                        OK = False
                        Exit Do
                    End If
                Loop
            Case 2 'Access Denied
                WScript.Echo Space(6) & "Error: Access Denied"
            Case 3 'Insufficient Privilege
                WScript.Echo Space(6) & "Error: Insufficient Privilege"
            Case 8 'Unknown Failure
                WScript.Echo Space(6) & "Error: Unknown Failure"
            Case 9 'Path not found
                WScript.Echo Space(6) & "Error: Path not found"
            Case 21 'Invalid Parameter
                WScript.Echo Space(6) & "Error: Invalid Parameter"
            Case Else
                WScript.Echo Space(6) & "Error code " & errReturn & " not found"
        End Select

    End If

    'WScript.Echo Space(6) & "Done."
    Err.Clear
    RunRemoteProcess = OK
End Function

Private Function GetDriveWithHighestFreeSpace(server)
Dim MAX
Dim DRIVE
Dim objWMI
Dim colItems
Dim objItem

    On Error Resume Next

    MAX = 1
    DRIVE = ""

    Set objWMI = GetObject("winmgmts:\\" & server & "\root\cimv2")

    If Err.Number = 0 Then
        Set colItems = objWMI.ExecQuery("Select * from Win32_LogicalDisk WHERE DriveType=3")
        For Each objItem in colItems
            If INT(MAX) < INT((objItem.FreeSpace / objItem.Size) * 1000)/10 Then
                DRIVE = objItem.Name
                MAX = INT((objItem.FreeSpace / objItem.Size) * 1000)/10
            End If
        Next
    End If

    Err.Clear

    If DRIVE = "C:" Then
        WScript.Echo Space(2) & "ERROR: Only 1 drive found"
        DRIVE = ""
    End If

    GetDriveWithHighestFreeSpace = Replace(DRIVE, ":", "")
End Function

Private Function AddZero(d)
    AddZero = Right("0" & d, 2)
End Function

'------------------
Private Sub DoExit
    If (Err.Number <> 0) Then
        WScript.Echo "Error: " & Err.Number & ": " & Err.Description
        DiagnosedFlag = True
    End If

    WScript.Echo

    WScript.Echo "Total servers processed: " & TotalServersProcessed
    WScript.Echo "Total servers in file  : " & TotalServersInFile
    If StrComp(TotalServersProcessed, TotalServersInFile) <> 0 Then
        WScript.Echo "Warning: Servers processed is not equal to total servers."
        DiagnosedFlag = True
    End If

    WScript.Echo
    WScript.Echo "Elapsed time: " & DateDiff("s", startTime, Now) & " seconds."
    WScript.Echo

    WScript.Echo """"
    If (DiagnosedFlag) Then
        WScript.Echo "RBA diagnose"
    Else
        WScript.Echo "RBA success"
    End If

    WScript.Quit cReturnCode
End Sub

Private Sub CheckParameter
    ParamSERVERLIST = WScript.Arguments.Named("SERVERLIST")
    ParamOLDIIS     = WScript.Arguments.Named("IIS")
    ParamOLDOTHER   = WScript.Arguments.Named("OTHERS")

    If Not FileExists(ParamSERVERLIST) Then
        ParamSERVERLIST = CreateObject("Scripting.FileSystemObject").GetParentFolderName(Wscript.ScriptFullName)
        If Not FileExists(ParamSERVERLIST) Then
            ParamSERVERLIST = ""
        End If
    End If

    If ParamOLDIIS = "" Then
        ParamOLDIIS = 30
    End If

    If ParamOLDOTHER = "" Then
        ParamOLDOTHER = 360
    End If

    If WScript.Arguments.Count = 0 Or ParamSERVERLIST = "" Then
        WScript.Echo "PURPOSE: "
        WScript.Echo "INPUTS: [/SERVERLIST:path_to_list] [/IIS:iis_age] [/OTHERS:others_age][/HELP]"
        WScript.Echo "PARAMETERS:"
        WScript.Echo "  path_to_list - optional, file containing list of servers to process"
        WScript.Echo "               - default, look at current folder for serverlist.txt"
        WScript.Echo "  iis_age      - optional, file age of IIS logs in days to check"
        WScript.Echo "               - default, 30 days"
        WScript.Echo "  others_age   - optional, file age of all logs in days to check"
        WScript.Echo "               - default, 360 days"
        WScript.Echo "  /HELP        - display this help"
        WScript.Echo
        WScript.Echo "USAGE:"
        WScript.Echo "  < script >"
        WScript.Echo "    - Run against serverlist.txt"
        WScript.Echo "  < script > /IIS:40"
        WScript.Echo "    - Change default 30 days for IIS logs"
        WScript.Echo "  < script > /SERVERLIST:""c:\temp\serverlist.txt"""
        WScript.Echo "    - Run against c:\temp\serverlist.txt"
        WScript.Echo "  < script >"
        WScript.Echo "    - Display this help screen"
        WScript.Echo
        WScript.Echo "REQUIRED:"
        WScript.Echo "  Increase HP SA script output to 999 KB"
        WScript.Echo " """
        WScript.Echo "RBA diagnose"
        WScript.Quit cReturnCode
    End If

    GetServerListFromFile
End Sub