' Title: Copy Text Annotation.BAS
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

    App.feAppMessage(FCM_COMMAND, "Copy Text Annotation")
    App.feAppMessage(FCM_NORMAL, "Copy a text annotation from a base location.")

    Dim rc As Long

    Dim txt As femap.Text
    Set txt = App.feText()

    Dim txt_id As Long
    Dim old_xyz As Variant
    Dim new_xyz As Variant
    Dim diff_xyz(3) As Double


    txt.SelectID("Pick text...")
    txt_id = txt.ID
    App.feCoordPick("Pick location to copy from...", old_xyz)
    rc = App.feCoordPick("Pick location to copy to...", new_xyz)

    Do While rc = FE_OK
        diff_xyz(0) = new_xyz(0) - old_xyz(0)
        diff_xyz(1) = new_xyz(1) - old_xyz(1)
        diff_xyz(2) = new_xyz(2) - old_xyz(2)

        txt.Get(txt_id)

        txt.TextPosition(0) = txt.TextPosition(0) + diff_xyz(0)
        txt.TextPosition(1) = txt.TextPosition(1) + diff_xyz(1)
        txt.TextPosition(2) = txt.TextPosition(2) + diff_xyz(2)
        txt.PointerPosition(0) = txt.PointerPosition(0) + diff_xyz(0)
        txt.PointerPosition(1) = txt.PointerPosition(1) + diff_xyz(1)
        txt.PointerPosition(2) = txt.PointerPosition(2) + diff_xyz(2)

        txt.Put(txt.NextEmptyID())

        rc = App.feCoordPick("Pick location to copy to...", new_xyz)
        App.feViewRegenerate(0)
    Loop
End Sub