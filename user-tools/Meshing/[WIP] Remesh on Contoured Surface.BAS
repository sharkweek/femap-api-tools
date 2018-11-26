' Title: Remesh on Contoured Surface.BAS
' Author: Andy Perez
' Date: November 2018
' Femap API Version: 12.0

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Remesh on Contoured Surface")
    App.feAppMessage(FCM_NORMAL, "Remesh shell elements on a contoured surface. The new mesh will derive its properties from the original:")
    App.feAppMessage(FCM_ERROR, "surface and are automatically be reassociated with the surface.")

    Dim rc As Long

    Dim eSet As femap.Set
    Set eSet = App.feSet

    Dim nSet As femap.Set
    Set nSet = App.feSet

    Dim e As femap.Elem
    Set e = App.feElem

    Dim n As femap.Node
    Set n = App.feNode

    Dim tracker As femap.TrackData
    Set tracker = App.feTrackData

    App.feRemesh()

End Sub

