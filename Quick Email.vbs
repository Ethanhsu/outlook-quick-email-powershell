' Quick Email launcher - double-click this file to run
' No admin rights, no prompts, no SmartScreen issues
On Error Resume Next

Dim shell, shortcut, fso
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Get script directory (where this .vbs is located)
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)

' Path to the PowerShell script
psScript = scriptDir & "\quick-email.ps1"

' Create shortcut on Desktop if it doesn't exist
desktopPath = shell.SpecialFolders("Desktop")
shortcutPath = desktopPath & "\Quick Email.lnk"

If Not fso.FileExists(shortcutPath) Then
    Set shortcut = shell.CreateShortcut(shortcutPath)
    shortcut.TargetPath = "powershell.exe"
    shortcut.Arguments = "-ExecutionPolicy Bypass -WindowStyle Hidden -File """ & psScript & """"
    shortcut.WorkingDirectory = scriptDir
    shortcut.Description = "Quick Email - Create preset emails fast"
    shortcut.IconLocation = "outlook.exe,0"
    shortcut.Save
    WScript.Echo "Shortcut created on Desktop: " & shortcutPath
End If

' Run the PowerShell script
shell.Run "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & psScript & """", 1, False

If Err.Number <> 0 Then
    WScript.Echo "Error: " & Err.Description
End If