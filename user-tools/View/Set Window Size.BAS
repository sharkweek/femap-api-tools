' Title: Set Window Size.BAS
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 2301

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered trade­mark
' of Siemens Digital Industries Software Inc. or its subsidiaries in the United
' States and in other countries.

' Usage Notes
' -----------
' This script is intended for manually setting the view window size so that
' any screenshots or generated GIF files can be consistent regardless of
' monitor resolution.

' public declarations
Dim App As femap.model

Sub Main
    ' `Main` declarations
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Set Window Size")

    Dim rc As Long

    Begin Dialog UserDialog 210,119,"Window Dimensions",.DialogFunc ' %GRID:10,7,1,1
        Text 10,7,90,14,"x:",.x_text,1
        Text 10,28,90,14,"y: ",.y_text,1
        Text 10,49,90,14,"Width:",.width_text,1
        Text 10,70,90,14,"Height",.height_text,1
        TextBox 110,7,90,14,.x
        TextBox 110,28,90,14,.y
        TextBox 110,49,90,14,.width
        TextBox 110,70,90,14,.height
        OKButton 10,91,90,21
        CancelButton 110,91,90,21
    End Dialog
    Dim dlg As UserDialog

    ' prompt user
    If Dialog(dlg) = 0 Then
        Exit Sub  ' exit if user cancels
    End If

    App.feWindowSetSize(0, Val(dlg.x), Val(dlg.y), Val(dlg.width), Val(dlg.Height))

    App.feViewRegenerate(0)

End Sub


Private Function DialogFunc(DlgItem$, Action%, SuppValue?) As Boolean
    Select Case Action%
        Case 1 ' Dialog box initialization
            DlgText "x", CStr(0)
            DlgText "y", CStr(0)
            ' initial dimensions are for 1080P in 3x2 dimensions
            DlgText "width", CStr(1440)
            DlgText "height", CStr(1080)
    End Select
End Function