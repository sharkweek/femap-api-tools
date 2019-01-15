' Title: [WIP] Break Curve at Plane.bas
' Author: Andy Perez
' Date: December 2018
' Femap API Version: 12.0

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "[WIP] Break Curve at Plane")
    App.feAppMessage(FCM_NORMAL, "Breaks a curve at the point it intersects with a specified plane.")

    Dim rc As Long

    App.feProjectOntoPlane(FT_POINT, entitySet, planeLoc, planeNormal)
    App

    get coord bounding box in plane csys
    create surface at plane location mirroring bounding box
    
    App.feCurveBreak(c.ID, break_loc)

End Sub