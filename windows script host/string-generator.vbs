'-------------------------------
Option Explicit
Function RndB(byVal fromInt, byVal toInt)
  Randomize()
  Dim r
  r = (toInt - fromInt + 1) * Rnd + fromInt
  RndB = Int(r)
End Function

Function RndString(byVal strLen)
  Dim newStr, letters, i
  letters = "!#$%&*,./0123456789:;<=>?@ABCDEGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdeghijklmnopqrstuvwxyz~"
  For i = 1 to strLen
    newStr = newStr & mid(letters,RndB(1, Len(letters)),1)
  Next
  RndString = newStr
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

Function expMulti(byVal int, byVal log)
  Dim i, res
  res = 1
  For i = 1 to log
    res = res * int
  Next
  expMulti = res
End Function

Sub main()
  Dim f, sPrefix, sTmp, fName, oShell, i, sDic, strLen, time
  fName = "stringoutput.txt"

  Set f = CreateObject("Scripting.FileSystemObject").OpenTextFile(fName, 2, True)
  Set sDic = CreateObject("Scripting.Dictionary")
  Set oShell = CreateObject("Wscript.Shell")

  sTmp = ""
  sPrefix = "J@b1l"
  strLen = 7
  time = Minute(Now)
  While Hour(Now) <> 0
    i = 0
    Do
      sTmp = sPrefix & RndString(strLen)
      If Not sDic.Exists(sTmp) Then
        i = i + 1
        f.WriteLine(sTmp)
        sDic.Add sTmp, strLen
		If Minute(Now) > time Then Exit Sub
      End If
    Loop Until i >= expMulti(80, strLen)
    strLen = strLen + 1
    f.WriteLine(strLen & " new length of strings ")
    If Minute(Now) > time Then Exit Sub
  Wend
  
  f.Close
End Sub

main()
WScript.Echo "OK"