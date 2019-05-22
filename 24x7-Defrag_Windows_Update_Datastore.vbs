Option Explicit

Dim oShell
Dim command
Dim oExec
Set oShell = CreateObject("Wscript.Shell")
Dim shellOutput
command = "esentutl /d C:\Windows\SoftwareDistribution\DataStore\DataStore.edb"


Set oExec = oShell.Exec(command)

Do While oExec.Status = 0
                Wscript.Sleep 100
Loop

Select Case oExec.Status
                Case 1
                                shellOutput = oExec.StdOut.ReadAll
                Case 2
                                shellOutput = oExec.StdErr.ReadAll
End Select

WScript.Echo shellOutput

Dim objWMIService, objItem, colItems, strComputer
On Error Resume Next
strComputer = "."
Set objWMIService = GetObject _
("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery _
("Select * from Win32_LogicalDisk")

For Each objItem in colItems
Wscript.Echo "Computer: " & objItem.SystemName & VbCr & _ 
" ==================================" & VbCr & _ 
" Drive Letter: " & objItem.Name & vbCr & _ 
" Description: " & objItem.Description & vbCr & _ 
" Volume Name: " & objItem.VolumeName & vbCr & _ 
" Size: " & Int(objItem.Size /1073741824) & " GB" & vbCr & _ 
" Free Space: " & Int(objItem.FreeSpace /1073741824) & _
" GB" & vbCr & _ 
"" 
Next

WSCript.Quit
