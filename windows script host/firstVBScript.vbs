Option Explicit
WScript.Echo "What's up!"
On error resume next
Set objShell = CreateObject("Shell.Application")
Set objNS = objShell.namespace(&h2f)
set colItems = objNS.items
For each objItem in colItems
wScript.Echo objItem.Name
Next