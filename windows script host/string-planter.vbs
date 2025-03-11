'-------------------------------
Option Explicit
Dim oShell, oName, fName, fso
Set oShell = CreateObject("Wscript.Shell")
Dim strA()
oName = "Zscaler Client Connector"
fName = "ATree.txt"
Set fso = CreateObject("Scripting.FileSystemObject")

Sub main()
	Dim aLine, iA, f, f1, t, newfName
	If fso.FileExists(fName) Then 
		newfName = Left(fName,Len(fName)-4) & "New.txt"
		fso.CopyFile fName,  newfName
		Set f1 = fso.OpenTextFile(newfName, 1)
		Set f = fso.CreateTextFile(fName & ".log",2, True)
		'
	J@b1l%5g!~@G
J@b1lK=g%sY;
		Set f = fso.OpenTextFile(fName, 2, True)
		iA = 0
		Do Until f1.AtEndOfStream
			t = Time(): iA = iA + 1
			aLine = f1.ReadLine
			If not mySendKeys(oName, aLine) Then Wscript.Echo newStr: Exit Sub
			t = Time() - t
			f.WriteLine(iA & " " & aLine & " " & t)
		Loop	
		f.Close	
	End If
End Sub

Function escapedKstrokes( byVal str)
  Dim newStr, i
  newStr = ""
  For i = 1 to Len(str)
    Select Case Mid(str,i,1)
      Case "+","^","%","~","{","}": newStr = newStr & "{" & Mid(str,i,1) & "}"
      Case Else: newStr = newStr & Mid(str,i,1)
    End Select
  Next
  escapedKstrokes = newStr
End Function

Function MySendKeys(strApp, strKeys)
    MySendKeys = False
	If oShell.AppActivate(strApp) Then
		Wscript.Sleep 500
        	oShell.SendKeys escapedKstrokes(strKeys)
		Wscript.Sleep 500
		oShell.SendKeys "{ENTER}+{TAB}"
        MySendKeys = True
    End If
End Function

main()
WScript.Echo "OK"