' Title: Attach Results Using Batch Script.BAS
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 2301

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered trade­mark
' of Siemens Digital Industries Software Inc. or its subsidiaries in the United
' States and in other countries.

' public declarations
Dim App As femap.model

Sub Main
    ' `Main` declarations
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Attach Results Using Batch Script")

    Dim rc As Long

    Dim filepath As String
    filepath = GetFilePath(,"Batch|*.bat|Powershell|*.ps1|Any|*.*", "Pick batch script...", 0)

    Set readOld = CreateObject("Scripting.FileSystemObject")
    FileOpen

    App.feViewRegenerate(0)

End Sub