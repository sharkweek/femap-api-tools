' Title: Ordered Copy Nodes To Clipboard.bas
' Author: Andy Perez
' Date: May 2019
' Femap API Version: 12.0


Sub Main
'   Declarations
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Ordered Copy Nodes As List")
    App.feAppMessage(FCM_NORMAL, "Copy node IDs to clipboard in order by selected criteria.")

    Dim rc As Long

    Dim nSort As femap.SortSet
    Set nSort = App.feSort

    Dim nSet As femap.Set
    Set nSet = App.feSet

    Dim n As femap.Node
    Set n = App.feNode

    Dim nList As String

    Dim length As Double
    Dim uBase As Variant
    Dim uDir As Variant

    Dim vDir(3) As Double

    Dim dot As Double

    Dim crv As femap.Curve
    Set crv = App.feCurve
    Dim crvNodes As femap.Set
    Set crvNodes = App.feSet
    Dim s As Double

'   Prompt user to select nodes
    rc = nSort.Select(FT_NODE, True, "Select nodes...")
    If rc = FE_CANCEL Then
        Exit Sub
    ElseIf rc = FE_FAIL Then
        rc = App.feAppMessage(FCM_ERROR, "Selected nodes do not exist.")
    End If

'   Dialog Box
    Begin Dialog UserDialog 235,245,"Ordered Copy Nodes as List" ' %GRID:10,7,1,1
        GroupBox 10,7,215,182,"Order by:",.GroupBox1
        OptionGroup .copy_type
            OptionButton 30,28,140,14,"Node &ID",.node_id
            OptionButton 30,49,160,14,"&Selection Order",.selection_order
            OptionButton 30,70,110,14,"Along &Vector",.along_vector
            OptionButton 30,91,110,14,"Along &Curve",.along_curve
            OptionButton 30,112,185,14,"Distance from Coo&rdinate",.dist_coord
            OptionButton 30,133,160,14,"Distance from &Axis",.dist_axis
            OptionButton 30,154,150,14,"A&bsolute distance",.dist_plane
        OKButton 10,217,100,21
        CancelButton 120,217,100,21
        CheckBox 20,196,190,14,"Print to Messages Window",.print_msg
        text 51,170,90,14,"from plane",.dist_plane_text
    End Dialog
    Dim dlg As UserDialog

    ' prompt user
    If Dialog(dlg) = 0 Then
        Exit Sub  ' exit if user hits 'Cancel'
    End If

'   Sorting
    Select Case dlg.copy_type
        Case 0  ' node id
            nSort.GetSet(nSet.ID, True)
            nSort.Clear
            Do While n.NextInSet(nSet.ID)
                nSort.Add(n.ID)
            Loop

        Case 1  ' selection order
            ' selection order is the default

        Case 2  ' along vector
            If App.feVectorPick("Select vector...", True, length, uBase, uDir) = FE_CANCEL Then
                Exit Sub
            End If

            ' add to sort object and sort
            nSort.Reset()
            Do While nSort.Next()
                n.Get(nSort.Current())
                For i = 0 To 2
                    vDir(i) = n.xyz(i) - uBase(i)
                Next

                App.feVectorDotProduct(vDir, uDir, dot)

                nSort.SetReal(nSort.Current(), dot, 0, 0)
            Loop

            ' sort by curve coordinate
            nSort.SortRemoveDuplicates(True)

        Case 3  ' along curve
            ' pick curve
            rc = crv.SelectID("Select curve...")
            If rc = FE_CANCEL Then
                Exit Sub
            ElseIf rc = FE_NOT_EXIST Then
                App.feAppMessage(FCM_ERROR, "Selected curve does not exist.")
                Exit Sub
            End If

            ' get xyz-coords from nodes
            Do While n.NextInSet(nSet.ID)
                n.Get(n.ID)
                crv.XYZToParam(n.xyz, s)
                nSort.AddReal(n.ID, s, 0, 0)
            Loop

            ' sort by curve coordinate
            nSort.SortRemoveDuplicates(True)

        Case 4  ' dist from coordinate
            ' prompt user to pick coordinate
            rc = App.feCoordPick("Select coordinate...", uBase)
            If rc = FE_CANCEL Then
                Exit Sub
            End If

            nSort.GetSet(nSet.ID, True)
            nSort.Clear()
            nSort.AddAroundPoint(FT_NODE, nSet.ID, uBase)
            nSort.SortRemoveDuplicates(True)

        Case 5  ' dist from axis
            ' pick axis vector
            If App.feVectorPick("Select axis vector...", True, length, uBase, uDir) = FE_CANCEL Then
                Exit Sub
            End If

            ' add to Sort object and sort
            nSort.GetSet(nSet.ID, True)
            nSort.Clear()
            nSort.AddAroundVector(FT_NODE, nSet.ID, uBase, uDir)
            nSort.SortRemoveDuplicates(True)

        Case 6  ' dist from plane
            ' pick plane normal
            If App.fePlanePick("Select plane...", uBase, uDir, vDir) = FE_CANCEL Then
                Exit Sub
            End If

            ' add to Sort object and sort
            nSort.GetSet(nSet.ID, True)
            nSort.Clear()
            nSort.AddAroundPlane(FT_NODE, nSet.ID, uBase, uDir, True)
            nSort.SortRemoveDuplicates(True)

    End Select

    ' add ordered nodes to string
    nList = ""
    nSort.Reset()
    If dlg.print_msg = False Then
        Do While nSort.Next()
            nList = nList & CStr(nSort.Current()) & vbCrLf
        Loop
    Else  ' print copied IDs to messages window
        Do While nSort.Next()
            nList = nList & CStr(nSort.Current()) & vbCrLf
            App.feAppMessage(FCM_NORMAL, CStr(nSort.Current()))
        Loop
    End If

    ' copy to clipboard
    Clipboard(nList)

End Sub
