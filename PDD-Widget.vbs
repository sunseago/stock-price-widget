' PDD Stock Widget - silent launcher
' Invoked by the Desktop shortcut: hides the console, shows only the widget.
' Kept pure ASCII on purpose: VBScript has no reliable UTF-8 support.
Option Explicit
Dim shell, fso, scriptDir, psFile, psExe, cmd
Set shell = CreateObject("WScript.Shell")
Set fso   = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
psFile    = scriptDir & "\PDD-Widget.ps1"
' Absolute path for PowerShell: do not rely on PATH (some machines / AV rewrite it)
psExe = shell.ExpandEnvironmentStrings("%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe")
If Not fso.FileExists(psExe) Then psExe = "powershell.exe"
cmd = """" & psExe & """ -NoProfile -ExecutionPolicy Bypass -Sta -WindowStyle Hidden -File """ & psFile & """""
shell.Run cmd, 0, False
