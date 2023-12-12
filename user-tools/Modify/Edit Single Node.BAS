' Title: Edit Single Node.BAS
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 12.0

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered trade­mark
' of Siemens Digital Industries Software Inc. or its subsidiaries in the United
' States and in other countries.

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Edit Single Node")

    Dim rc As Long

    Dim n As femap.Node
    Set n = App.feNode

    Dim new_coord As Variant

    Do While n.SelectID("Select node...") = FE_OK
        App.feCoordPick("Pick new location for node...", new_coord)
        n.Get(n.ID)
        n.xyz = new_coord
        n.Put(n.ID)
        App.feViewRegenerate(0)
    Loop

End Sub