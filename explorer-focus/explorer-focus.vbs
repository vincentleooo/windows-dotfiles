' File: explorer-focus.vbs
' Purpose: Run PowerShell script silently

Dim shell,command

Set objShell = CreateObject("Wscript.Shell")

' Full path to your PowerShell script
ps1 = "D:\repos\file-explorer-2.ps1"

' Build command (bypass policy + hidden)
cmd = "powershell.exe -WindowStyle hidden -NoLogo -NoProfile -ExecutionPolicy Bypass -File """ & ps1 & """"

' Run hidden (0 = hidden, False = don’t wait to finish)
objShell.Run cmd, 0, False
