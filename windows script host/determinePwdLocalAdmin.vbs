'https://social.technet.microsoft.com/Forums/ie/en-US/9810d2b8-2274-49cc-a45b-f93bd5de59b5/determining-the-local-administrator-password-vbscript?forum=ITCG
'--------

Option Explicit
'On Error Resume Next

Dim objNet, objFSO, objLogStream
Dim strComputer, strLogFile

Set objNet = Wscript.CreateObject("Wscript.Network")
strComputer = objNet.ComputerName

Set objFSO = CreateObject("Scripting.FileSystemObject")
strLogFile = ".\" & strComputer & ".txt" 

If objFSO.FileExists(strLogFile) Then
 Call AccountProperties(strComputer, objFSO, strLogFile)
Else
 Set objLogStream = objFSO.CreateTextFile(strLogFile)
End If

objLogStream.Close

Call AccountProperties(strComputer, objFSO, strLogFile)

Public Function AccountProperties (strComputer, objFSO, strLogFile)
 
 Dim objUser
 Dim strFlag, strProperty

 Const ADS_UF_ACCOUNTDISABLE = &H0002 
 Const ADS_UF_PASSWD_NOTREQD = &H0020 
 
 Set objUser = GetObject("WinNT://" & strComputer & "/administrator")
 strFlag = objUser.Get("UserFlags")

 If strFlag AND ADS_UF_ACCOUNTDISABLE Then
 strProperty = 1
 Call Logging(strComputer, strProperty, objFSO, strLogFile)
 End If

 If strFlag AND ADS_UF_PASSWD_NOTREQD Then
 strProperty = 2
 Call Logging(strComputer, strProperty, objFSO, strLogFile)
 End If
 
 Set objUser = Nothing
 Set strFlag = Nothing
 Call CommandPrompt(strComputer, strLogFile)
End Function

Public Function CommandPrompt(strComputer, strLogFile)

 Dim objShell

 set objShell = Wscript.CreateObject("WScript.Shell")
 objShell.Run "RunAs /noprofile /user:administrator ""cmd.exe /c echo " & strComputer & ",Password is set to helpdesky2k >> " & strLogFile & """"
 WScript.Sleep 150
 objShell.SendKeys "123456{ENTER}"

 WScript.Sleep 500

 objShell.Run "RunAs /noprofile /user:administrator ""cmd.exe /c echo " & strComputer & ",Password is blank >> " & strLogFile & """"
 WScript.Sleep 150
 objShell.SendKeys "{ENTER}"
 
 Set objShell = Nothing
 Set strComputer = Nothing
 Call ScriptExit
End Function

Public Function Logging(strComputer, strProperty, objFSO, strLogFile)

Dim objLogStream

Const ForAppending = 8

Set objLogStream = objFSO.OpenTextFile(strLogFile, ForAppending, True)

 If strProperty = 1 Then
 objLogStream.WriteLine "" & strComputer & "Account is disabled"
 Else 
 If strProperty = 2 Then
 objLogStream.WriteLine "" & strComputer & "Password is not required"
 End If
 End If
 
 objLogStream.Close
 
 Set objFSO = Nothing
 Set objLogStream = Nothing
 Set strLogFile = Nothing
 Set strComputer = Nothing
 Set strProperty = Nothing
 Call ScriptExit
End Function

Public Function ScriptExit
 Wscript.Quit
End Function
