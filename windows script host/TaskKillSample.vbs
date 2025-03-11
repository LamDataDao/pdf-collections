'-------------------------------
'-----Task Kill ----------------
'-------------------------------
Option Explicit
Const strComputer = "." 
Dim WshShell : Set WshShell = CreateObject("WScript.Shell")
Dim objWMIService, colProcessList, objProcess
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colProcessList = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'C:\Program Files (x86)\Zscaler\ZSAService\ZSATunnel.exe'")
For Each objProcess in colProcessList 
  WshShell.Exec "PSKill " & objProcess.ProcessId 
Next
