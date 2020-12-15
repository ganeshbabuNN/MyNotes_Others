strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
  & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colSettings = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")

For Each objComputer in colSettings
  Wscript.Echo "System Manufacturer: " & objComputer.Manufacturer
  Wscript.Echo "System Model: " & objComputer.Model
Next

Set colBIOS = objWMIService.ExecQuery("Select * from Win32_BIOS")

For each objBIOS in colBIOS
  Wscript.Echo "Serial Number: " & objBIOS.SerialNumber
  Wscript.Echo
Next