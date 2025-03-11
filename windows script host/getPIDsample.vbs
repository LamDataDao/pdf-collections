Sub getPIDviaTaskList()
	Dim Tasks
	Tasks = WScript.CreateObject("WScript.Shell").Exec("tasklist /v /fo csv").StdOut.ReadAll()
	For Each task In Split(Tasks,vbCrLf)
	'task = Split(Trim(task),",")
    		objTextFile.WriteLine(task)
    	Next
	'Task title "Image Name","PID","Session Name","Session#","Mem Usage","Status","User Name","CPU Time","Window Title"
End Sub

Sub getPIDviaWin32Process()
	Dim objWMIService, processes
	Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")

	Set processes = objWMIService.ExecQuery("Select * From Win32_Process where Name <> 'svchost.exe'")
	For Each process in processes
	    	objTextFile.WriteLine(process.name & " " & process.ProcessID & " " & process.CommandLine & " " & process.Handle)
	    	'If process.name = "Calculator.exe" then process.terminate
	Next
End Sub

Dim fSO, task
Set fSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = fSO.OpenTextFile(".\pids.txt", 2, True)
getPIDviaWin32Process
objTextFile.Close
Wscript.Echo "Done"
