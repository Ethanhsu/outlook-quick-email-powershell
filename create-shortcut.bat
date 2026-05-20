@echo off
:: Create Quick Email shortcut on Desktop
powershell -ExecutionPolicy Bypass -Command ^
  "$ws = New-Object -ComObject WScript.Shell; ^
  $s = $ws.CreateShortcut(\"$env:USERPROFILE\Desktop\Quick Email.lnk\"); ^
  $s.TargetPath = 'powershell.exe'; ^
  $s.Arguments = '-ExecutionPolicy Bypass -File \"%~dp0quick-email.ps1\"'; ^
  $s.WorkingDirectory = '%~dp0'; ^
  $s.Description = 'Quick Email - Create preset emails fast'; ^
  $s.Save(); ^
  Write-Host 'Shortcut created on Desktop!'"