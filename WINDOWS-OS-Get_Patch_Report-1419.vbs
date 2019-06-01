Option Explicit

' NAME: WINDOWS-OS-Get_Patch_Report-1419.vbs
' PURPOSE: Extract the following information:
'          Server Name, OS, Service Pack, LastBootUpTime, Model, Specified KB Patch Report
' INPUTS: KBListToCheck
'
' WHERE:
'   KBListToCheck - Partial KB numbers to check
'
' USAGE:
'   < script > KB12,KB234,KB5678
'
' REPORT:
'   Server Name, OS, Service Pack, LastBootUpTime, Model, Specified KB Patch Report
'
' NOTE:
'   Increase HP SA script output to 9999 KB
'
' AUTHOR(s): J, Kamalakannan (ITO-DO) <kamalakannan.j@hpe.com>
' DATE WRITTEN: 02 Mar 2015
' MODIFICATION HISTORY: 02 Mar 2015 - initial release
'

On Error Resume Next

Dim objWMI
Dim colItems
Dim objItem

Dim var_report
Dim hotfixes
Dim hotfix
Dim wmi_query
Dim i

If WScript.Arguments.Count = 0 Then
  WScript.Echo "KB input not found"
  WScript.Echo "Syntax:   < script >   KB12,KB234,KB5678"
  WScript.Quit(404)
End If

var_report = CreateObject("Wscript.Network").ComputerName & ","

Set objWMI = GetObject("WinMgmts:\\.\root\cimv2")
If Err.Number <> 0 Then
  WScript.Echo Err.Number & ": " & Err.Description
  WScript.Quit(1)
End If

Set colItems = objWMI.ExecQuery("SELECT Caption,CSDVersion,LastBootUpTime From WIN32_OperatingSystem")
For Each objItem In colItems
  var_report = var_report & Trim(Replace(objItem.Caption,","," ")) & ","
  var_report = var_report & objItem.CSDVersion & ","
  var_report = var_report & DateSerial(Left(objItem.LastBootUpTime,4), _
               Mid(objItem.LastBootUpTime,5,2), _
               Mid(objItem.LastBootUpTime,7,2)) + _
               TimeSerial(Mid(objItem.LastBootUpTime,9,2), _
               Mid(objItem.LastBootUpTime,11,2), _
               Mid(objItem.LastBootUpTime,13,2)) & ","
Next

Set colItems = objWMI.ExecQuery("SELECT Model From WIN32_ComputerSystem")
For Each objItem In colItems
  var_report = var_report & objItem.model & ","
Next

hotfixes = Split(WScript.Arguments.Item(0), ",")
i = 0

For Each hotfix In hotfixes
  hotfix = Trim(hotfix)
  If i = UBound(hotfixes) Then
    wmi_query = wmi_query & "Hotfixid like '" & hotfix & "%'"
  Else
    wmi_query = wmi_query & "Hotfixid like '" & hotfix & "%' or "
  End If
  i = i + 1
Next

Set colItems = objWMI.ExecQuery("SELECT HotFixID FROM win32_quickfixengineering WHERE " & wmi_query)
For Each objItem In colItems
  var_report = var_report & objItem.hotfixid & ","
Next

While Right(var_report, 1) = ","
  var_report = Left(var_report, Len(var_report)-1)
Wend

Set colItems = Nothing
Set objWMI = Nothing

WScript.Echo var_report

WScript.Quit(0)
    
