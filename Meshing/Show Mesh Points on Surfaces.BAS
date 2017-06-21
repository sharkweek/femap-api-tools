Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Show Mesh Points on Surfaces")

    Dim rc As Long

    Dim surfSet As femap.Set
    Set surfSet = App.feSet

    Dim surf As femap.Surface
    Set surf = App.feSurface

    Dim meshPointSet As femap.Set
    Set meshPointSet = App.feSet

    ' Prompt to select surfaces
    rc = surfSet.Select(FT_SURFACE, False, "Select surfaces")
    If rc = FE_CANCEL Then
        Exit Sub
    ElseIf rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "Selected surfaces do not exist.")
        Exit Sub
    End If

    ' Add mesh points to set
    Do While surf.NextInSet(surfSet.ID)
        meshPointSet.AddArray(149, surf.vMeshPoint)
    Loop

    ' Show mesh points'
    App.feViewShow(3, meshPointSet.ID)

End Sub
