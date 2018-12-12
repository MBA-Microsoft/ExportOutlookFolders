@ECHO OFF
set userEmail=email.email@email.com
SET ThisScriptsDirectory=%~pd0
SET PowerShellScriptPath=%ThisScriptsDirectory%GetOutlookFolders.ps1
Powershell -NoProfile -ExecutionPolicy Bypass -Command "& '%PowerShellScriptPath%' '%userEmail%'"
