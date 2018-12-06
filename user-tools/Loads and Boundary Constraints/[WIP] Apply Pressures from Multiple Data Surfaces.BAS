' Title: Apply Pressures from Multiple Data Surfaces.BAS
' Author: Andy Perez
' Date: November 2018
' Femap API Version: 12.0

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Apply Pressures from Multiple Data Surfaces")
    App.feAppMessage(FCM_NORMAL, "Maps pressure loads from multiple data surfaces to a set of ")
    App.feAppMessage(FCM_NORMAL, "selected elements. Loads will be created as multiple load sets.")

    Dim rc As Long

    

End Sub