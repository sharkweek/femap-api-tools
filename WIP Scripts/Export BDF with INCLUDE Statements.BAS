'Macro Export by Group
'A. Pavageau 05/10
'v2 04/2012 - Tel: 5721

Sub Main
    Dim App As femap.model
    Set App = feFemap

    Dim doublePrec As Long
    Dim bdfName As String

    Dim grp As femap.Group

    Dim runDeck As femap.AnalysisMgr

    Dim grpLoadSet As femap.Set
    Dim grpSPCSet As femap.Set
    Dim grpElmSet As femap.Set
    Dim grpNodeSet As femap.Set
    Dim grpMatSet As femap.Set
    Dim grpPropSet As femap.Set
    Dim grpCsysSet As femap.Set
    Dim grSet As femap.Set
    Dim grRefSet As femap.Set

    Set runDeck = App.feAnalysisMgr

    Set grpLoadSet = App.feSet
    Set grpSPCSet = App.feSet
    Set grpElmSet = App.feSet
    Set grpNodeSet = App.feSet
    Set grpMatSet = App.feSet
    Set grpPropSet = App.feSet
    Set grpCsysSet = App.feSet

    'retrieve folder path
    Dim objShell
    Dim objFolder
    Dim objFolderItem

    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.BrowseForFolder(0, "Select folder to write file(s)",&H1)
    If objFolder Is Nothing Then
        Exit Sub
    Else
        Set objFolderItem = objFolder.Self
        ChDir(objFolderItem.Path)
    End If

    Set objFolderItem = Nothing
    Set objFolder = Nothing
    Set objShell = Nothing

    ' Select analysis manager to export
    If runDeck.SelectID("Select Analysis Set to Export") = FE_CANCEL Then
        End Sub
    End If

    ' Get bulk data group specified by Analysis Set
    grp.Get(runDeck.NasBulkGroupID)  'may not need this

    grpLoadSet.AddGroup(FT_LOAD_DEFINITION, runDeck.NasBulkGroupID)
    grpSPCSet.AddGroup(FT_BC_DEFINITION, runDeck.NasBulkGroupID)
    grpElmSet.AddGroup(FT_ELEM, runDeck.NasBulkGroupID)
    grpNodeSet.AddGroup(FT_NODE, runDeck.NasBulkGroupID)
    grpMatSet.AddGroup(FT_MATL, runDeck.NasBulkGroupID)
    grpPropSet.AddGroup(FT_PROP, runDeck.NasBulkGroupID)
    grpCsysSet.AddGroup(FT_CSYS, runDeck.NasBulkGroupID)

    Do While gr.NextInSet(grSet.ID)
        Set grRefSet = gr.ReferencedGroups
        If grRefSet.Count <> 0 Then
            Exit Do
        End If

        runDeck.NasBulkGroupID = gr.ID
        runDeck.Put(runDeck.ID)

        'deal with group names...
        bdfName = gr.title

        If bdfName = "" Then
            bdfName = "Group" & CStr(gr.ID) & ".bdf"
        End If

        If InStr(bdfName,".") = 0 Then
            bdfName = bdfName & ".bdf"
        End If

        While InStr(bdfName,"/") <> 0
            bdfName = Mid(bdfName,InStr(bdfName,"/")+1)
        Wend

        App.feFileWriteNastran(0, bdfName)

        Call modOutputFiles(bdfName,doublePrec)
        App.feAppMessage(0, "File \'" & bdfName & "\' created")

    Loop

    App.feAppMessage(0, "File \'" & bdfName & "\' created")

End Sub

'sub which deletes useless lines automatically writtenby FEMAP at the start and end of files
Private Sub modOutputFiles(file As String, doubleP As Long)
    Dim curline As String
    Dim s1 As String
    Dim l1 As Long

    If doubleP = False Then
        s1 = "CORD2S         2" : l1 = 16
    Else
        s1 = "CORD2S*                2" : l1 = 24
    End If

    Open file For Input As #1
    Open "temp" For Output As #2

    'look for definition of CSys 2
    Do
        Line Input #1, curline
    Loop Until Left(curline,l1) = s1
    Line Input #1, curline

    If doubleP = True Then
        Line Input #1, curline
        Line Input #1, curline
    End If

    Line Input #1, curline
    While Left(curline,7) <> "ENDDATA"
        Print #2, curline
        Line Input #1, curline
    Wend

    Close #1
    Close #2

    Kill file
    Name "temp" As file
End Sub
