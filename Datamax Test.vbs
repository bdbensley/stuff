'code to test the datamax printers, finds serial ports and prints to them.
Dim fso, f
dim id
On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_SerialPort")
Set fso = CreateObject("Scripting.FileSystemObject")
For Each objItem in colItems
	'finds port name - will be COMx  no colon
    id = objItem.DeviceID
	Set f = fso.OpenTextFile(id,2,True,False)
	' Write data to the port
	f.Write (Chr(02) + "T" + chr(13))
	f.Write (chr(02) + "L" + chr(13))
	f.Write ("D11131100001000050" + id + chr(13))
	f.Write ("Q0001" + chr(13))
	f.Write ("E" + chr(13))
	f.Close
	wscript.echo "Device ID: " & id
next

