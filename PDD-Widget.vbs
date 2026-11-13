' PDD Stock Widget - silent launcher
' Invoked by the Desktop shortcut: hides the console, shows only the widget.
Option Explicit
Dim shell, fso, scriptDir, psFile, cmd
Set shell = CreateObject("WScript.Shell")
Set fso   = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
psFile    = scriptDir & "\PDD-Widget.ps1"
cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -Sta -WindowStyle Hidden -File """ & psFile & """"
shell.Run cmd, 0, False
