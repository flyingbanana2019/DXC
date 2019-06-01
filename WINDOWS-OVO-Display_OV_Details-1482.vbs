Option Explicit
On Error Resume Next

' NAME: WINDOWS-OVO-Display_OV_Details-1482.vbs
' PURPOSE: Uptime, ovc, ovc -status, opcagt -status, opcagt -version, opctemplate,
'          perfstat, ovcert -check, ovcodautil -showds, dir C:\osit\etc, C:\osit\bin\acf\acf.cmd
'
' INPUTS: [/LOCAL | /MULTI] [/OVOPATH:"ovopath"]
' PARAMETERS:
'   LOCAL | /LOCAL - run against the local server only
'   MULTI | /MULTI - get list of servers from ServerList.txt
'                  - you may need to supply login credentials in HP SA Runtime Options
'   ovopath        - full path of folder where opcagt.bat is located
'   no param       - display usage help
'
' USAGE:
'   < script > /LOCAL
'     - Get details from local server
'   < script > /MULTI /OVOPATH:"C:\Program Files\HP OpenView\bin"
'     - Get details from the servers in the ServerList.txt file
'   < script >
'     - Display usage help
'
' NOTE:
'   HP OVO must be installed in the remote server(s)
'   Admin share C$ must be enabled on the remote server(s)
'   Increase HP SA script output to 999 KB
'
' OUTPUT:
'   Console
'     -------------------------------------
'     Hostname
'     -------------------------------------
'     OK
'     -------------------------------------
'
'   OVDETAILS-YYYY-MM-dd-hhmmss.csv
'   Date  HostName  Uptime  Status  ovc -version  opcagt  opcagt -version opctemplate perfstat  ovcert -check ovcodautil -showds   (List of file and folders in the directory) C:\osit\etc>dir  C:\osit\bin\acf\acf.cmd
'
' AUTHOR(s): Harris Co
' DATE WRITTEN: 23 Sep 2016
' MODIFICATION HISTORY: 23 Sep 2016 - initial release
'

Const cRetSuccess = 0
Const FORREADING = 1
Const FORWRITING = 2

Dim strComputer:  strComputer = CreateObject("WScript.Shell").Environment("Process")("COMPUTERNAME")
Dim ROOTFOLDER:   ROOTFOLDER  = WScript.CreateObject("WScript.Shell").CurrentDirectory
Dim REPORTFILE:   REPORTFILE  = ROOTFOLDER & "\" & "OVDETAILS-" & Year(Now) & "-" & Right("0" & Month(Now), 2) & "-" & Right("0" & Day(Now), 2) & "-" & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2) & ".csv"
Dim SERVERFILE:   SERVERFILE  = ROOTFOLDER & "\ServerList.txt"

'containers
Dim ServerList:   Set ServerList = CreateObject("Scripting.Dictionary")

Dim ParamMULTI
Dim ParamPATH

Dim ReportFileOK:       ReportFileOK = False

Dim objReportFile
Dim ReportFileError:    ReportFileError = False
Dim headerFlag:         headerFlag = False
Dim HostName
Dim TotalServersInFile
Dim TotalServersProcessed
Dim FatalError:         FatalError = False

Dim startTime

startTime = Now

CheckParameter

OpenReportFile

GetDetails

'DeleteOldLogFiles

DoExit


Private Sub OpenReportFile()
  On Error Resume Next

  ReportFileOK = True

  Set objReportFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(REPORTFILE, FORWRITING, True)

  If Err.Number <> 0 Then
    WScript.Echo "Error: " & REPORTFILE & " " & Err.Description
    ReportFileOK = False
    Err.Clear
  End If
End Sub

Private Function Quotes(str)
  Quotes = """" & str & """"
End Function

Private Sub ShowResults(server, Uptime, ovcstatus, ovcversion, opcagt, opcagtversion, opctemplate, perfstat, ovcertcheck, ovcodautilshowds, dirosit, acft)
  On Error Resume Next
  Err.Clear

  If ReportFileOK Then
    If Not headerFlag Then
      objReportFile.WriteLine "Date,HostName,Uptime,Status,ovc -version,opcagt,opcagt -version,opctemplate,perfstat,ovcert -check,ovcodautil -showds,(List of file and folders in the directory) C:\osit\etc,C:\osit\bin\acf"
    End If

    objReportFile.WriteLine Now & "," & server & "," & Quotes(Uptime) & "," & Quotes(ovcstatus) & "," & Quotes(ovcversion) & "," & Quotes(opcagt) & "," & Quotes(opcagtversion) & "," & _
                            Quotes(opctemplate) & "," & Quotes(perfstat) & "," & Quotes(ovcertcheck) & "," & Quotes(ovcodautilshowds) & "," & Quotes(dirosit) & "," & Quotes(acft)

    headerFlag = True
  End If

  If server <> HostName Then
    'new
    HostName = server
    WScript.Echo "-------------------------------------"
    WScript.Echo server
    WScript.Echo "-------------------------------------"
  End If

  WScript.Echo Space(2) & Uptime

  TotalServersProcessed = TotalServersProcessed + 1

  If Err.Number <> 0 Then
    WScript.Echo "WriteReport Error: " & Err.Number & " " & Err.Description
    ReportFileError = True
    Err.Clear
  End If
End Sub

Private Function RunRemote(server, ovoPath, command, parameter, id)
Dim tempFile, cmd
Dim objWMIService, colProcesses, intProcessID

Dim out

  If FatalError Then
    RunRemote = ""
    Exit Function
  End If

  tempFile = "ovdetails_" & id & "_delete_me.txt"

  WScript.Echo Space(2) & "Executing " & command & " " & parameter

  'Check existence of executables
  If FolderExists("\\" & server & "\C$") Then
    If FileExists("\\" & server & "\" & Replace(ovoPath, ":\", "$\") & "\" & command) Then

      Set objWMIService = GetObject("winmgmts:\\" & server & "\root\cimv2:Win32_Process")

      If FileExists("\\" & server & "\c$\" & tempFile) Then
        cmd = "cmd /c del c:\" & tempFile & " /f /q"
        objWMIService.Create cmd, Null, Null, intProcessID
      End If

      cmd = "cmd /c """ & ovoPath & "\" & command & """ " & parameter & " > " & "C:\" & tempFile & " 2>&1  && Exit"
      objWMIService.Create cmd, Null, Null, intProcessID

      Do While True
        WScript.Sleep 2000
        Set colProcesses = GetObject("winmgmts:\\" & server & "\root\cimv2").ExecQuery("Select * from Win32_Process Where ProcessID = " & intProcessID)
        If colProcesses.Count = 0 Then Exit Do
      Loop

      If Not FileExists("\\" & server & "\c$\" & tempFile) Then
        WScript.Sleep 10000   'give time for process to create file
      ENd If

      If FileExists("\\" & server & "\c$\" & tempFile) Then
        Dim tempfilesize

        tempfilesize = CreateObject("Scripting.FileSystemObject").GetFile("\\" & server & "\c$\" & tempFile).Size
        WScript.Echo Space(4) & "Output Size: " & tempfilesize

        If (tempfilesize > 0) Then
          Dim objFile

          Set objFile = CreateObject("Scripting.FileSystemObject").OpenTextFile("\\" & server & "\c$\" & tempFile, FORREADING)
          out = objFile.ReadAll

          While Right(out, 2) = vbCrLf
            out = Left(out, Len(out) - 2)
          Wend

          While Right(out, 1) = vbCr
            out = Left(out, Len(out) - 1)
          Wend

          While Right(out, 1) = vbLf
            out = Left(out, Len(out) - 1)
          Wend

          While Left(out, 2) = vbCrLf
            out = Right(out, Len(out) - 2)
          Wend

          While Left(out, 1) = vbCr
            out = Right(out, Len(out) - 1)
          Wend

          While Left(out, 1) = vbLf
            out = Right(out, Len(out) - 1)
          Wend

          While InStr(1, out, "  ") > 0
            out = Replace(out, "  ", " ")
          Wend

          out = Trim(Replace(out, """", "'"))

          objFile.Close
          Set objFile = Nothing

        Else
          WScript.Echo Space(4) & "\\" & server & "\c$\" & tempFile & " is empty"
          out = "\\" & server & "\c$\" & tempFile & " is empty"
        End If

      Else
        out = "Something went wrong, output file not created. Please check permission of file \\" & server & "\c$\" & tempFile
        WScript.Echo Space(4) & out
        If Err.Number <> 0 Then
          out = out & vbCrLf & "ERROR: " & Err.Number & ": " & Err.Description
          Err.Clear
        End If
      End If

      If FileExists("\\" & server & "\c$\" & tempFile) Then
        cmd = "cmd /c del c:\" & tempFile & " /f /q"
        objWMIService.Create cmd, Null, Null, intProcessID
      End If

    Else
      out = "\\" & server & "\" & Replace(ovoPath, ":\", "$\") & "\" & command & " does not exists. Please use /PATH."
      WScript.Echo Space(4) & out
    End If

  Else
    out = "\\" & server & "\C$ is not shared. Please enable."
    WScript.Echo Space(4) & out
    FatalError = True
  End If

  Set objWMIService = Nothing
  Set colProcesses  = Nothing

  RunRemote = out
End Function

Private Sub GetDetails()
Const HeaderSize = 300

Dim server
Dim subTime, count
Dim ovoPath
Dim exitCode, errCode

Dim Uptime, ovcstatus, ovcversion, opcagt, opcagtversion, opctemplate, perfstat, ovcertcheck, ovcodautilshowds, dirosit, acft


  Err.Clear
  On Error Resume Next

  count = 1

  For Each server In ServerList.Keys
    WScript.Echo "Processing (" & count & ") " & server
    subTime = Now

    If ParamPATH = "" Then
      ovoPath = GetOpenViewPath(server)
      While Right(ovoPath, 1) = "\"
        ovoPath = Mid(ovoPath, 1, Len(ovoPath) - 1)
      Wend

      If FileExists("\\" & server & "\" & Replace(ovoPath, ":\", "$\") & "\bin\opcagt.bat") Then
        ovoPath = ovoPath & "\bin"

      Else
        ovoPath = ovoPath & "\bin\win64"
      End If

    Else
      ovoPath = ParamPATH
      While Right(ovoPath, 1) = "\"
        ovoPath = Mid(ovoPath, 1, Len(ovoPath) - 1)
      Wend
    End If

    WScript.Echo Space(2) & "OVO Path: " & ovoPath

    Uptime = "": ovcstatus = "": ovcversion = "": opcagt = "": opcagtversion = "": opctemplate = "": perfstat = "": ovcertcheck = "": ovcodautilshowds = "": dirosit = "": acft = ""

    Uptime = GetServerUptime(server)
    ovcstatus = RunRemote(server, ovoPath, "ovc.exe", "-status", 1)
    ovcversion = RunRemote(server, ovoPath, "ovc.exe", "-version", 1)
    opcagt = RunRemote(server, ovoPath, "opcagt.bat", "", 2)
    opcagtversion = RunRemote(server, ovoPath, "opcagt.bat", "-version", 3)
    opctemplate = RunRemote(server, ovoPath, "opctemplate.bat", "", 4)
    ovcertcheck = RunRemote(server, ovoPath, "ovcert.exe", "-check", 5)
    ovcodautilshowds = RunRemote(server, ovoPath, "ovcodautil.exe", "-showds", 6)
    dirosit = Execute("C:\", "cmd /c ", "dir \\" & server & "\c$\osit\etc", 120, exitCode, errCode, False)
    acft = RunRemote(server, "C:\osit\bin\acf", "acf.cmd", "-t", 7)

    ovoPath = GetOpenViewPath(server)
    While Right(ovoPath, 1) = "\"
      ovoPath = Mid(ovoPath, 1, Len(ovoPath) - 1)
    Wend

    If FileExists("\\" & server & "\" & Replace(ovoPath, ":\", "$\") & "\bin\perfstat.exe") Then
      ovoPath = ovoPath & "\bin"

    Else
      ovoPath = ovoPath & "\bin\win64"
    End If

    perfstat = RunRemote(server, ovoPath, "perfstat.exe", "", 8)

    ShowResults server, Uptime, ovcstatus, ovcversion, opcagt, opcagtversion, opctemplate, perfstat, ovcertcheck, ovcodautilshowds, dirosit, acft

    WScript.Echo Space(4) & DateDiff("s", subTime, Now) & " seconds."
    count = count + 1
  Next

  WScript.Echo "-------------------------------------"
  WScript.Echo " "
End Sub

Private Function GetServerUptime(server)
Dim colItem, objItems
Dim intPerfTimeStamp, intPerfTimeFreq, intCounter
Dim iUptimeInSec
Dim ConvSec, ConvMin, ConvHour, ConvDays

  On Error Resume Next

  Set colItem = GetObject("winmgmts:\\.\root\cimv2").ExecQuery("SELECT * FROM Win32_PerfRawData_PerfOS_System")
  For Each objItems In colItem
   intPerfTimeStamp = objItems.Timestamp_Object
   intPerfTimeFreq = objItems.Frequency_Object
   intCounter = objItems.SystemUpTime
  Next

  iUptimeInSec = (intPerfTimeStamp - intCounter)/intPerfTimeFreq

  ConvSec = iUptimeInSec Mod 60
  ConvMin = (iUptimeInSec Mod 3600) \ 60
  ConvHour =  (iUptimeInSec Mod (3600 * 24)) \ 3600
  ConvDays =  iUptimeInSec \ (3600 * 24)

  If Err.Number = 0 Then
    GetServerUptime = ConvDays & " day(s), " & ConvHour & " hour(s), " _
               & ConvMin & " minute(s), " & ConvSec & " second(s)"

  Else
    GetServerUptime = "ERROR: " & Err.Number & " " & Err.Description
  End If
End Function

Private Function GetOpenViewPath(server)
Dim result

  On Error Resume Next
  Err.Clear

  ReadRegistryKeyValue server, "SOFTWARE\Hewlett-Packard\HP OpenView", "InstallDir", result
  If (result = "") Then
    result = GetEnv(server, "OvInstallDir")
    If (result = "") Then
      result = "C:\Program Files\HP OpenView"
    End If
  End If

  GetOpenViewPath = result
End Function

Private Function GetEnv(server, name)
Dim colItems, objItem
Dim temp

  On Error Resume Next

  temp = ""
  Set colItems = GetObject("winmgmts:\\" & server & "\root\cimv2").ExecQuery("SELECT * FROM Win32_Environment WHERE Name = '" & name & "'")
  For Each objItem in colItems
    temp = objItem.VariableValue
  Next

  Set colItems = Nothing

  Err.Clear
  GetEnv = temp
End Function

Private Sub ReadRegistryKeyValue(server, path, name, ByRef value)
Const HKEY_LOCAL_MACHINE = &H80000002
Dim oReg

  On Error Resume Next
  Set oReg = GetObject("winmgmts:\\" & server & "\root\default:StdRegProv")
  oReg.GetStringValue HKEY_LOCAL_MACHINE, path, name, value

  If (Err.Number <> 0 Or IsNull(value)) Then
    value = ""
  End If

  Set oReg = Nothing

  Err.Clear
End Sub

Private Sub DeleteOldLogFiles()
Const AGEINDAYS = 30

Dim colItem, objItems

  On Error Resume Next

  Set colItem = CreateObject("Scripting.FileSystemObject").GetFolder(ROOTFOLDER).Files

  For Each objItems in colItem
    If InStr(1, objItems.Name, REPORTFILE, vbTextCompare) > 0 Then
      If objItems.DateLastModified < (Date() - AGEINDAYS) Then
        objItems.Delete(True)
      End If
    End If
  Next

  Set colItem  = Nothing
  Set objItems = Nothing

  Err.Clear
End Sub

Private Function GetServerListFromFile(multi)
Const ADVARCHAR  = 200
Const MAXCHARACTERS = 255

Dim obj, item
Dim temp

  On Error Resume Next

  Set temp = CreateObject("ADOR.RecordSet")
  temp.Fields.Append "Item", ADVARCHAR, MAXCHARACTERS
  temp.Open

  TotalServersInFile = 1

  If Not multi Then
    'add current server only for local
    temp.AddNew
    temp("Item") = UCase(strComputer)
    temp.Update
  End If

  If multi Then
    If Not FileExists(SERVERFILE) Then
      WScript.Echo SERVERFILE & " does not exists.  Default=local"
    Else
      Set obj = CreateObject("Scripting.FileSystemObject").OpenTextFile(SERVERFILE, FORREADING)
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
  End If

  If Err.Number <> 0 Then
    WScript.Echo SERVERFILE & " cannot be read"
  Else
    If multi Then
      TotalServersInFile = temp.RecordCount
      WScript.Echo SERVERFILE & " contains " & TotalServersInFile & " servers"
      If TotalServersInFile = 0 Then
        WScript.Echo "Please list down server names in " & SERVERFILE
        WScript.Echo
        DoExit
      End If
    End If
  End If

  Sort temp, ServerList

  temp.Close
  Set temp = Nothing

  Err.Clear
End Function

Private Sub Sort(source, ByRef list)
Dim item

  If source.RecordCount = 0 Then Exit Sub

  source.Sort = "Item"
  source.MoveFirst

  Do Until source.EOF
    item = source("Item")
    If Not list.Exists(item) Then
      list.Add item, "No"
    End If
    source.MoveNext
  Loop
End Sub

Private Function FileExists(file)
  FileExists = CreateObject("Scripting.FileSystemObject").FileExists(file)
End Function

Private Function FolderExists(folder)
  FolderExists = CreateObject("Scripting.FileSystemObject").FolderExists(folder)
End Function

Private Sub DoExit()
  If (Err.Number <> 0) Then
    WScript.Echo "Error: " & Err.Number & ": " & Err.Description
  End If

  If ReportFileOK Then
    objReportFile.Close
    If Not ReportFileError Then
      WScript.Echo "Report written to " & REPORTFILE
    Else
      WScript.Echo "Please review errors regarding " & REPORTFILE
    End If
  Else
    WScript.Echo "Report was not written to " & REPORTFILE & " due to errors, please check."
  End If

  Set ServerList    = Nothing
  Set objReportFile = Nothing

  WScript.Echo

  WScript.Echo "Total servers processed: " & TotalServersProcessed
  WScript.Echo "Total servers in file  : " & TotalServersInFile
  If StrComp(TotalServersProcessed, TotalServersInFile) <> 0 Then
    WScript.Echo "Warning: Servers processed is not equal to total servers."
  End If

  WScript.Echo
  WScript.Echo "Elapsed time: " & DateDiff("s", startTime, Now) & " seconds."
  WScript.Echo

  WScript.Quit cRetSuccess
End Sub

Private Sub CheckParameter()
  On Error Resume Next

  ParamPATH  = WScript.Arguments.Named("OVOPATH")
  ParamMULTI = WScript.Arguments.Named.Exists("MULTI")

  If (WScript.Arguments.Count = 0) Then
    WScript.Echo "PURPOSE: Uptime, ovc, ovc -status, opcagt -status, opcagt -version, opctemplate,"
    WScript.Echo "         perfstat, ovcert -check, ovcodautil -showds, dir C:\osit\etc, C:\osit\bin\acf\acf.cmd"
    WScript.Echo
    WScript.Echo "INPUTS: [/LOCAL | /MULTI] [/OVOPATH:""ovopath""]"
    WScript.Echo "PARAMETERS:"
    WScript.Echo "  LOCAL | /LOCAL - run against the local server only"
    WScript.Echo "  MULTI | /MULTI - get list of servers from ServerList.txt"
    WScript.Echo "                 - you may need to supply login credentials in HP SA Runtime Options"
    WScript.Echo "  ovopath        - full path of folder where opcagt.bat is located"
    WScript.Echo "  no param       - display usage help"
    WScript.Echo
    WScript.Echo "USAGE:"
    WScript.Echo "  < script > /LOCAL"
    WScript.Echo "    - Get details from local server"
    WScript.Echo "  < script > /MULTI /OVOPATH:""C:\Program Files\HP OpenView\bin"""
    WScript.Echo "    - Get details from the servers in the ServerList.txt file"
    WScript.Echo "  < script >"
    WScript.Echo "    - Display usage help"
    WScript.Echo
    WScript.Echo "NOTE:"
    WScript.Echo "  HP OVO must be installed in the remote server(s)"
    WScript.Echo "  Admin share C$ must be enabled on the remote server(s)"
    WScript.Echo "  Increase HP SA script output to 999 KB"
    WScript.Echo

    WScript.Quit cRetSuccess
  End If

  GetServerListFromFile(ParamMULTI)
End Sub

' Return:
'     StdOut of command - success
'     timeout - error
'     exitCode - result of cmd
'     errCode  - result of function
Private Function Execute(currDir, command, param, timeout, exitCode, errCode, terminateFlag)
Dim runStart
Dim timedOut
Dim oExec
Dim WshShell

  exitCode = 0
  errCode  = 0

  runStart = Now()

  Err.Clear
  On Error Resume Next

  Set WshShell = WScript.CreateObject("WScript.Shell")
  WshShell.CurrentDirectory = currDir

  Set oExec = WshShell.Exec(command & " " & param)

  timedOut = False
  If (Err.Number = cRetSuccess) Then
    Do While (oExec.Status = cRetSuccess)
      WScript.Sleep 100
      If (DateDiff("s", runStart, Now) > timeout) Then
        timedOut = True
        Exit Do
      End If
    Loop

    If (terminateFlag) Then
      oExec.Terminate()
    End If

    If (timedOut) Then
      Execute = oExec.StdOut.ReadAll()
      If (Execute = "") Then
        Execute = oExec.StdErr.ReadAll()
        If (Execute = "") Then
          errCode = 999
        End If
      End If
    Else
      Execute = oExec.StdOut.ReadAll()
      If (Execute = "") Then
        Execute = oExec.StdErr.ReadAll()
      End If
    End If
  Else
    Execute = Err.Description
    errCode = Err.Number
  End If

  exitCode = oExec.ExitCode

  'oExec above also errors with 404 if file is not found,
  'need to clear error since we already trapped it above.
  Err.Clear

  Set WshShell = Nothing
  Set oExec = Nothing
End Function
                                                    
