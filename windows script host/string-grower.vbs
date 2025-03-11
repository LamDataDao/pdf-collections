'-------------------------------
Option Explicit
Dim oShell, oName, fName, letters, fso
Set oShell = CreateObject("Wscript.Shell")
Dim strA()
'letters = "!#$%&*,./0123456789:;<=>?@ABCDEGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdeghijklmnopqrstuvwxyz~"
letters = "@0123456789ABCDEGHIJKLMNOPQRSTUVWXYZabcdeghijklmnopqrstuvwxyz"
oName = "Zscaler Client Connector"
fName = "PwdText"
Set fso = CreateObject("Scripting.FileSystemObject")


Function loadStr()
	Dim oldStr, iA, f1
	If fso.FileExists(fName) Then 
		fso.CopyFile fName, fName & "tmp"
		Set f1 = fso.OpenTextFile(fName & "tmp", 1)
		Do Until f1.AtEndOfStream
			oldStr = f1.ReadLine
		Loop
		WScript.Echo oldStr
		loadStr = oldStr
		For iA = 1 to Len(oldStr)
			strA(iA - 1) = InStr(letters, mid(oldStr, iA, 1)) - 1
		Next
	End If
End Function

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

Sub main()
	Dim lenStr, iA, lLimit, newStr, tOut, f
	lenStr = 32:	lLimit = Len(letters) - 1
	ReDim strA(lenStr)
	lenStr = Len(loadStr())
	Set f = fso.OpenTextFile(fName, 2, True)
	tOut = Now
	Do
		For iA = 0 to lenStr
			If strA(lenStr - iA) < lLimit Then
				strA(lenStr - iA) = strA(lenStr - iA) + 1
				Exit For
			Else
				strA(lenStr - iA) = 0
			End If
		Next
		newStr = ""
		For iA = 0 to lenStr
			newStr = newStr & CStr(Mid(letters, strA(iA) + 1, 1))
		Next
		f.WriteLine(newStr)
		If not mySendKeys(oName, newStr) Then Wscript.Echo newStr: Exit Sub
		'If not mySendKeys("Untitled - Notepad", newStr) then :
	Loop Until (tOut + 0.001) = Now
	f.Close
End Sub

main()
WScript.Echo "OK"