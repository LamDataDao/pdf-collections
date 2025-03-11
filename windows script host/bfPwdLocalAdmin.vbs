'https://social.technet.microsoft.com/Forums/ie/en-US/9810d2b8-2274-49cc-a45b-f93bd5de59b5/determining-the-local-administrator-password-vbscript?forum=ITCG
'--------

Option Explicit
'On Error Resume Next

Dim strSourceFile, strOutputFile

strSourceFile = "c:\scripts\local administrator check\list.txt"
strOutputFile = "c:\scripts\local administrator check\log.txt"

Call ReadSourceFile(strSourceFile, strOutputFile)

Public Function ReadSourceFile(strSourceFile, strOutputFile)
 Dim objFSO, objSourceFile
 Dim strComputer, strPassword
 
 strPassword = "123456"
 
 Const ForReading = 1, ForWriting = 2, ForAppending = 8
 
 Set objFSO = CreateObject("Scripting.FileSystemObject")
 Set objSourceFile = objFSO.OpenTextFile(strSourceFile, ForReading)

 Do Until objSourceFile.AtEndOfStream
 strComputer = UCase(objSourceFile.ReadLine)
 If VerifyConnectivity(strComputer, strOutputFile) Then
 If Not CheckAdminPassword(strComputer, strOutputFile, strPassword) Then
 strPassword = "1234567"
 End If 
 If Not CheckAdminPassword (strComputer,strOutputFile, strPassword) Then
 strPassword = ""
 End If
 If Not CheckAdminPassword (strComputer, strOutputFile, strPassword) Then
 strEventMessage = Date & "," & Time & "," & strComputer & ",Unable to determine what Admin password is set to."
 Call RecordEvent(strEventMessage, strOutputFile)
 End If
 End If
 Loop

 Set objSourceFile = Nothing
 Set objFSO = Nothing
 Set strPassword = Nothing
End Function

Public Function VerifyConnectivity(strComputer, strOutputFile)
 Dim objShell, objShellExecute
 Dim strCommand, strOutput, strEventMessage
 
 VerifyConnectivity = False
 strCommand = "%comspec% /c ping -l 0 -n 2 -w 1500 " & strComputer & ""
 
 Set objShell = CreateObject("Wscript.Shell")
 Set objShellExecute = objShell.Exec(strCommand)
 
 Do Until objShellExecute.StdOut.AtEndOfStream
 strOutput = objShellExecute.StdOut.ReadAll
 
 If InStr(strOutput, "Reply") > 0 Then
 VerifyConnectivity = True
 Else
 VerifyConnectivity = False
 strEventMessage = Date & "," & Time & "," & strComputer & ",Unable to connect to system"
 Call RecordEvent(strEventMessage, strOutputFile)
 End If
 Loop
 
 Set objShellExecute = Nothing
 Set objShell = Nothing
End Function

Public Function CheckAdminPassword(strComputer, strOutputFile, strPassword)
 Dim objSWbemLocator, objCheckAdminPassword, objSWbemServices
 Dim strEventMessage
 Dim colCheckAdminPassword
 Dim bolCheckAdminPassword
 
 Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
 Set objSWbemServices = objSWbemLocator.ConnectServer _
    (strComputer, "root\cimv2", strComputer &"\administrator", strPassword)
 objSWbemServices.Security_.ImpersonationLevel = 3
 Set colCheckAdminPassword = objSWbemServices.ExecQuery("Select * from Win32_OperatingSystem")
 
 bolAdminPassword = False
 
 For Each objCheckAdminPassword In colCheckAdminPassword
 If InStr(LCase(objCheckAdminPassword.Caption), "windows") > 0 Then
 bolCheckAdminPassword = True
 Else
 bolCheckAdminPassword = False
 End If
 Next
 
 CheckAdminPassword = bolAdminPassword
 If CheckAdminPassword = True Then
 strEventMessage = Date & "," & Time & "," & strComputer & ",Local Admin password is set to" & strPassword & ""
 Call RecordEvent(strEventMessage, strOutputFile)
 End If
End Function

Public Sub RecordEvent(strEventMessage, strOutputFile)
 Dim objFSO, objOutputFile

 Const ForReading = 1, ForWriting = 2, ForAppending = 8
 
 Set objFSO = CreateObject("Scripting.FileSystemObject")
 Set objOutputFile = objFSO.OpenTextFile(strOutputFile, ForAppending, True)
 
 objOutputFile.WriteLine(strEventMessage)
 WScript.Echo(strEventMessage)
 
 objOutputFile.Close
 
 Set objOutputFile = Nothing
 Set objFSO = Nothing
End Sub
