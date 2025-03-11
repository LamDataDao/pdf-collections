Set objFS = CreateObject("Scripting.FileSystemObject")
strFile = "c:\test\file.txt"
strTemp = "c:\test\temp.txt"
Set objFile = objFS.OpenTextFile(strFile)
Set objOutFile = objFS.CreateTextFile(strTemp,True)    
Do Until objFile.AtEndOfStream
    strLine = objFile.ReadLine
    ' do something with strLine 
    objOutFile.Write(strLine & "kndfffffff")
Loop
objOutFile.Close
objFile.Close
objFS.DeleteFile(strFile)
objFS.MoveFile strTemp,strFile 