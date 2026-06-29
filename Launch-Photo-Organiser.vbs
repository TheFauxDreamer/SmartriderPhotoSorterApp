' ============================================================
'  Photo Organiser launcher (no black window)
'  Double-click this file to open the Photo Organiser app.
'  Keep it in the SAME folder as Organise-Photos-GUI.ps1
' ============================================================

Dim shell, fso, here, ps1
Set shell = CreateObject("WScript.Shell")
Set fso   = CreateObject("Scripting.FileSystemObject")

here = fso.GetParentFolderName(WScript.ScriptFullName)
ps1  = here & "\Organise-Photos-GUI.ps1"

If Not fso.FileExists(ps1) Then
    MsgBox "Could not find Organise-Photos-GUI.ps1 in this folder." & vbCrLf & vbCrLf & here, vbCritical, "Photo Organiser"
    WScript.Quit
End If

' Launch PowerShell hidden so no console window appears - the app's own window shows instead.
shell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & ps1 & """", 0, False
