' Title: Ordered Copy Points As List.BAS
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 2306

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered trade­mark
' of Siemens Digital Industries Software Inc. or its subsidiaries in the United
' States and in other countries.

Sub Main
'   Declarations
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Ordered Copy Points As List")
    App.feAppMessage(FCM_NORMAL, "Copy point IDs to clipboard in order by selected criteria.")

    Dim rc As Long

    Dim pSort As femap.SortSet
    Set pSort = App.feSort

    Dim pSet As femap.Set
    Set pSet = App.feSet

    Dim p As femap.Point
    Set p = App.fePoint

    Dim pList As String

    Dim length As Double
    Dim uBase As Variant
    Dim uDir As Variant

    Dim vDir(3) As Double

    Dim dot As Double

    Dim crv As femap.Curve
    Set crv = App.feCurve
    Dim crvPoints As femap.Set
    Set crvPoints = App.feSet
    Dim s As Double

'   Prompt user to select points
    rc = pSort.Select(FT_POINT, True, "Select points...")
    If rc = FE_CANCEL Then
        Exit Sub
    ElseIf rc = FE_FAIL Then
        rc = App.feAppMessage(FCM_ERROR, "Selected points do not exist.")
    End If

'   Dialog Box
    Begin Dialog UserDialog 235,245,"Ordered Copy Points as List" ' %GRID:10,7,1,1
        GroupBox 10,7,215,182,"Order by:",.GroupBox1
        OptionGroup .copy_type
            OptionButton 30,28,140,14,"Point &ID",.point_id
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
            pSort.GetSet(pSet.ID, True)
            pSort.Clear
            Do While p.NextIpSet(pSet.ID)
                pSort.Add(n.ID)
            Loop

        Case 1  ' selection order
            ' selection order is the default

        Case 2  ' along vector
            If App.feVectorPick("Select vector...", True, length, uBase, uDir) = FE_CANCEL Then
                Exit Sub
            End If

            ' add to sort object and sort
            pSort.Reset()
            Do While pSort.Next()
                p.Get(pSort.Current())
                For i = 0 To 2
                    vDir(i) = p.xyz(i) - uBase(i)
                Next

                App.feVectorDotProduct(vDir, uDir, dot)

                pSort.SetReal(pSort.Current(), dot, 0, 0)
            Loop

            ' sort by curve coordinate
            pSort.SortRemoveDuplicates(True)

        Case 3  ' along curve
            ' pick curve
            rc = crv.SelectID("Select curve...")
            If rc = FE_CANCEL Then
                Exit Sub
            ElseIf rc = FE_NOT_EXIST Then
                App.feAppMessage(FCM_ERROR, "Selected curve does not exist.")
                Exit Sub
            End If

            ' get xyz-coords from points
            Do While p.NextInSet(crvPoints.ID)
                p.Get(p.ID)
                crv.XYZToParam(p.xyz, s)
                pSort.AddReal(p.ID, s, 0, 0)
            Loop

            ' sort by curve coordinate
            pSort.SortRemoveDuplicates(True)

        Case 4  ' dist from coordinate
            ' prompt user to pick coordinate
            rc = App.feCoordPick("Select coordinate...", uBase)
            If rc = FE_CANCEL Then
                Exit Sub
            End If

            pSort.GetSet(pSet.ID, True)
            pSort.Clear()
            pSort.AddAroundPoint(FT_NODE, pSet.ID, uBase)
            pSort.SortRemoveDuplicates(True)

        Case 5  ' dist from axis
            ' pick axis vector
            If App.feVectorPick("Select axis vector...", True, length, uBase, uDir) = FE_CANCEL Then
                Exit Sub
            End If

            ' add to Sort object and sort
            pSort.GetSet(pSet.ID, True)
            pSort.Clear()
            pSort.AddAroundVector(FT_NODE, pSet.ID, uBase, uDir)
            pSort.SortRemoveDuplicates(True)

        Case 6  ' dist from plane
            ' pick plane normal
            If App.fePlanePick("Select plane...", uBase, uDir, vDir) = FE_CANCEL Then
                Exit Sub
            End If

            ' add to Sort object and sort
            pSort.GetSet(pSet.ID, True)
            pSort.Clear()
            pSort.AddAroundPlane(FT_NODE, pSet.ID, uBase, uDir, True)
            pSort.SortRemoveDuplicates(True)

    End Select

    ' add ordered points to string
    pList = ""
    pSort.Reset()
    If dlg.print_msg False Then
        Do While pSort.Next()
            pList = pList & CStr(pSort.Current()) & vbCrLf
        Loop
    Else  ' print copied IDs to messages window
        Do While pSort.Next()
            pList = pList & CStr(pSort.Current()) & vbCrLf
            App.feAppMessage(FCM_NORMAL, CStr(pSort.Current()))
        Loop
    End If

    ' copy to clipboard
    Clipboard(pList)

End Sub
