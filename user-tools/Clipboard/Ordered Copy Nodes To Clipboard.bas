' Title: Ordered Copy Nodes To Clipboard.bas
' Author: Andy Perez
' Date: April 2019
' Femap API Version: 12.0

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Ordered Copy Nodes To Clipboard")
    App.feAppMessage(FCM_NORMAL, "Copy node IDs to clipboard ordered by selected criteria.")

    Dim rc As Long

    ' ordered by:
        ' node id
        ' along vector
        ' along curve
        ' distance from point

End Sub