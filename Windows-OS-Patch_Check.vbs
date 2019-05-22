On error resume next
'Option Explicit
Dim objFSO, objTextFile, RootDir, FileOutput, strComputer, objWMIService, Count

Const ForReading = 1

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile("C:\serverpatch.txt", ForReading)

Dim strHotFix
strHotFix = inputbox("Enter the patch KB Article")

set FileOutput = objFSO.CreateTextFile("C:\temp\Pat_Install_Status_KB.htm",true,false)

fileOutput.WriteLine("<HTML><HEAD><TITLE>Installed Patches Information</TITLE></HEAD><BODY>")
fileOutput.WriteLine("<TABLE border=""1"">")
fileOutput.WriteLine("<tr style=""background-color:#a0a0ff;font:8pt Arial;font-weight:bold;"" align=""left"">")
fileOutput.WriteLine("<TD>Computer Name</TD><TD>Description</TD><TD>Hotfix ID</TD><TD>Install Status</TD><TD>Installed On</TD><TD>Installed By</TD></TR>")

Do Until objTextFile.AtEndOfStream
    strComputer = objTextFile.Readline
    Count = 0
'Query command
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Dim QFEs    
Set QFEs = objWMIService.ExecQuery ("Select * from win32_QuickFixEngineering")

Dim strOutput
Dim QFE
    For Each QFE in QFEs
        if QFE.HotFixID = strHotFix then
            strOutput = "<TR style=""background-color:#e0f0f0;font:8pt Arial;"">"
            strOutput = strOutput + "<TD>" & strComputer & "</TD>" &_
                        "<TD>" & QFE.Description & "</TD>" &_
                        "<TD>" & QFE.HotFixID & "</TD>"  &_
                        "<TD>" & "Installed" & "</TD>"  &_
                        "<TD>" & QFE.InstalledOn & "</TD>" &_
                        "<TD>" & QFE.InstalledBy & "</TD>"
            Count= Count + 1    
            fileOutput.WriteLine(strOutPut)
            fileOutput.WriteLine("</TR>")
        end if
        
'        strOutPut=""
    Next
    if Count = 0 then
            strOutput = "<TR style=""background-color:#e0f0f0;font:8pt Arial;"" align=""center"">"
            strOutput = strOutput + "<TD>" & strComputer & "</TD>" &_
                        "<TD>" & "-" & "</TD>" &_
                        "<TD>" & strHotfix & "</TD>"  &_
                        "<TD>" & "Not Installed" & "</TD>"  &_
                        "<TD>" & "-" & "</TD>" &_
                        "<TD>" & "-" & "</TD>"
            fileOutput.WriteLine(strOutPut)
            fileOutput.WriteLine("</TR>")
    end if
Loop

fileOutput.WriteLine("</TABLE></BODY></HTML>")
wscript.echo "Result saved in Pat_Install_Status_KB.html file at specified location in line no 15 of this script"
WScript.Quit
