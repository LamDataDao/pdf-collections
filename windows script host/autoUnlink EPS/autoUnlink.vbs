Option Explicit
Dim fFullName, fSO
fFullName = ".\list.txt"
Set fSO = CreateObject("Scripting.FileSystemObject")

Function escapedKstrokes( byVal str)
  Dim newStr, i, currChr
  newStr = ""
  For i = 1 to Len(str)
  	currChr = Mid(str, i, 1)
    Select Case currChr
      Case "+","^","%","~","{","}": newStr = newStr & "{" & currChr & "}"
      Case Else: newStr = newStr & currChr
    End Select
  Next
  escapedKstrokes = newStr
End Function

Function MySendKeys(strApp, strKeys)
    MySendKeys = False
	If oShell.AppActivate(strApp) Then
		Wscript.Sleep 50
        	oShell.SendKeys escapedKstrokes(strKeys)
		Wscript.Sleep 50
		oShell.SendKeys "{ENTER}"
        MySendKeys = True
    End If
End Function

Function PIDsviaWin32Process()
	Dim objWMIService, processes, process, n
	Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")

	Set processes = objWMIService.ExecQuery("Select * From Win32_Process where Name <> 'svchost.exe'")
	n = 0
	For Each process in processes
		If InStr(process.name,"EPS")>0 Then
			PIDsviaWin32Process(n) = process.ProcessID
			n = n + 1
		End If
	Next
End Function

Sub main()
	'Khai bao
	Dim objWMIService, processes, process, processoShell, Success, oShell
	Set oShell = WScript.CreateObject("WScript.Shell")
	Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
	'Tim process ID chinh xac cua EPS
	Set processes = objWMIService.ExecQuery("Select * From Win32_Process where Name <> 'svchost.exe'")
	For Each process in processes
		If InStr(process.name,"EPS") > 0 Then
			Success = oShell.AppActivate(process.ProcessID)
			If Success Then Exit For
		End If
	Next
	'Check file list ton tai
	If fSO.FileExists(fFullName) Then
		Dim oTextFile
		If Not Success Then
			Wscript.Echo "Kindly open EPS then open Unlinking Tool and run me again"
			Exit Sub
		End If
		Set oTextFile = fSO.OpenTextFile(fFullName, 1)
		Do Until oTextFile.AtEndOfStream 
			oShell.AppActivate(process.ProcessID)
			oShell.SendKeys "{TAB}{TAB}{TAB}"
		    oShell.SendKeys oTextFile.ReadLine
			Wscript.Sleep 50
			oShell.SendKeys "{ENTER}"
			Wscript.Sleep 500
			oShell.SendKeys " %A"
			Wscript.Sleep 500
			oShell.SendKeys "%U"
		Loop
		WScript.Echo "Done"
	Else
		WScript.Echo "Input the serial_numbers you wish to unlink into the txt file with name 'list.txt'"
	End If
End Sub

main