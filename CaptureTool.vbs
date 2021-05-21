Option Explicit

' ======================================================================
'
' PowerShell Script Launcher
'
' Copyright (C) 2021 tag. All rights reserved.
'
' ======================================================================

' Objects
Dim objFso : Set objFso = CreateObject("Scripting.FileSystemObject")
Dim objShellApp : Set objShellApp = CreateObject("Shell.Application")
Dim objShell : Set objShell = WScript.CreateObject("WScript.Shell")

' File Base Name
Dim strScriptName
strScriptName = objFso.GetBaseName(WScript.ScriptFullName)

' Execute PowerShell Script
objShell.Run "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass "".\" & strScriptName & ".ps1""", 0, True