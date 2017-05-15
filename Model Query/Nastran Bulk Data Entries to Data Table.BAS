Sub Main
    Dim App As femap.model
    Set App = feFemap()

    Begin Dialog UserDialog 650,224,.dlgfunc ' %GRID:10,7,1,1
        OKButton 190,182,90,21,.okBox
        CancelButton 420,182,90,21,.cnBox
        text 40,7,330,28,"Select the type entities you would like to get the Bulk Data entries of",.Text1,2
        CheckBox 210,63,180,35,"Properties from Elements",.propElem
        CheckBox 50,56,90,14,"Properties",.prBox,1
        CheckBox 50,112,90,14,"Nodes",.ndBox
        CheckBox 50,140,90,14,"Elements",.elBox,1
        CheckBox 50,84,90,14,"Materials",.mtBox
        CheckBox 210,112,180,35,"Materials from Elements",.matElem
        GroupBox 440,42,170,126,"Sort By",.GroupBox1
        OptionGroup .Group1
            OptionButton 470,70,120,14,"ID",.OptionButton1
            OptionButton 470,105,130,14,"Selection Order",.OptionButton2
        CheckBox 470,133,110,14,"Ascending",.Acs
    End Dialog
    Dim dlg As UserDialog
    If Dialog(dlg) = 0 Then
        End
    End If

    Dim table As femap.DataTable
    Set table = App.feDataTable

    Dim mat As femap.Matl
    Set mat = App.feMatl

    Dim pSet As femap.SortSet
    Set pSet = App.feSort

    Dim eSet As femap.SortSet
    Set eSet = App.feSort

    Dim nSet As femap.SortSet
    Set nSet = App.feSort

    Dim fSet As femap.Set
    Set fSet = App.feSet

    Dim el As femap.Elem
    Set el = App.feElem

    Dim pr As femap.Prop
    Set pr = App.feProp

    Dim nd As femap.Node
    Set nd = App.feNode

    Dim sSet As femap.SortSet
    Set sSet = App.feSort

       Dim lUp As femap.Layup
    Set lUp = App.feLayup

    Dim gSet As femap.Set
    Set gSet = App.feSet

    Dim mSet As femap.SortSet
    Set mSet = App.feSort

    Dim func As femap.Function
    Set func = App.feFunction

    Dim sort As Long
    Dim tf As Double
    Dim t1 As Double
    Dim t2 As Double
    Dim t3 As Double
    Dim t4 As Double
    Dim cnt As Long
    Dim eRows As Variant
    Dim sDOF As String
    Dim list As Variant
    Dim ndList As Variant
    Dim dof As Variant
    Dim fc As Variant
    Dim wt As Variant
    Dim curCol As Long
    Dim elID As Long
    Dim topID As Long
    Dim glob As Boolean
    Dim pShell As Boolean
    Dim    pPlane As Boolean
    Dim typeID As Long
    Dim prID As Long
    Dim leng As Long
    Dim row As Long
    Dim mtType As Long
    Dim mtID As Long
    Dim j As Long
    Dim i As Long
    Dim k As Long
    Dim start As Long
    Dim propCase As Variant
    Dim rowID As Variant
    Dim bigMat As Variant
    Dim col1 As Variant
    Dim col2 As Variant
    Dim col3 As Variant
    Dim col4 As Variant
    Dim col5 As Variant
    Dim col6 As Variant
    Dim col7 As Variant
    Dim col8 As Variant
    Dim col9 As Variant

    If dlg.prBox Then
        pSet.Select(FT_PROP,True,"Please Select Property to Inspect")
        If dlg.group1 = 0 Then
            While pSet.Next
                pSet.SetInt(pSet.Current,pSet.Current,0,0)
            Wend
            pSet.Sort(False)
            pSet.Reset
        End If
        If dlg.Acs Then
            sort = pSet.Count
            While pSet.Next
                pSet.SetInt(pSet.Current,sort,0,0)
                sort = sort - 1
            Wend
            pSet.Sort(False)
            pSet.Reset
        End If
    End If


    If dlg.mtBox Then
        mSet.Select(FT_MATL,True,"Please Select Material to Inspect")
        If dlg.group1 = 0 Then
            While mSet.Next
                mSet.SetInt(mSet.Current,mSet.Current,0,0)
            Wend
            mSet.Sort(False)
            mSet.Reset
        End If
        If dlg.Acs Then
            sort = mSet.Count
            While mSet.Next
                mSet.SetInt(mSet.Current,sort,0,0)
                sort = sort - 1
            Wend
            mSet.Sort(False)
            mSet.Reset
        End If
    End If

    If dlg.ndBox Then
        nSet.Select(FT_NODE,True,"Select Nodes to get GRID for")
        If dlg.group1 = 0 Then
            While nSet.Next
                nSet.SetInt(nSet.Current,nSet.Current,0,0)
            Wend
            nSet.Sort(False)
            nSet.Reset
        End If
        If dlg.Acs Then
            sort = nSet.Count
            While nSet.Next
                nSet.SetInt(nSet.Current,sort,0,0)
                sort = sort - 1
            Wend
            nSet.Sort(False)
            nSet.Reset
        End If
    End If

    If dlg.elBox Then
        eSet.Select(FT_ELEM,True,"Please Select an element to inspect")
        If dlg.group1 = 0 Then
            While eSet.Next
                eSet.SetInt(eSet.Current,eSet.Current,0,0)
            Wend
            eSet.Sort(False)
            eSet.Reset
        End If
        If dlg.Acs Then
            sort = eSet.Count
            While eSet.Next
                eSet.SetInt(eSet.Current,sort,0,0)
                sort = sort - 1
            Wend
            eSet.Sort(False)
            eSet.Reset
        End If
    End If

    If dlg.propElem Then
        While eSet.Next
            el.Get(eSet.Current)
            pSet.Add(el.propID)
        Wend
        eSet.First
        eSet.Prev
    End If

   If dlg.matElem Then
        While eSet.Next
            el.Get(eSet.Current)
            pr.Get(el.propID)
            mSet.Add(pr.matlID)
        Wend
        eSet.First
        eSet.Prev
    End If

    ReDim bigMat(1 To 9, nSet.Count+ 4 * (pSet.Count + eSet.Count) + 1000)

    If dlg.prBox Or dlg.propElem Then
        While pSet.Next
            prID = pSet.Current
            pr.Get(prID)
            typeID = pr.type
            Select Case typeID
            Case FET_L_ROD
                Rod:
                bigMat(1,start) = "PROD"
                leng = 3
                propCase = Array(0,4,5,7)
                row = pValPrint(bigMat,start,leng,pr,False,propCase)
            Case FET_L_BAR
                bigMat(1,start) = "PBAR"
                leng = 16
                propCase = Array(0,1,2,4,7,Empty,8,9,10,11,12,13,14,15,5,6,3)
                row = pValPrint(bigMat,start,leng,pr,False,propCase)
                bigMat(9,start) = Empty
                If bigMat(2,start + 2) =  0 Then
                    bigMat(2,start + 2) = Empty
                End If
                If bigMat(3,start + 2) =  0 Then
                    bigMat(3,start + 2) = Empty
                End If
            Case FET_L_TUBE
                bigMat(1,start) = "PTUBE"
                leng = 3
                propCase = Array(0,Empty,7,8)
                row = pValPrint(bigMat,start,leng,pr,False,propCase)
                bigMat(5,start) = ((pr.pval(0) - pr.pval(1))/2)
                If bigMat(7,start) = 0 Then
                    bigMat(7,start) = Empty
                End If
            Case FET_L_BEAM
                bigMat(1,start) = "PBEAM"
                leng = 13
                propCase = Array(0,1,2,3,4,7,8,9,10,11,12,13,14,15)
                row = pValPrint(bigMat,start,leng,pr,False,propCase)
                row = 4
                bigMat(2,start + 2) = "YESA"
                bigMat(3,start + 2) = 1.0
                bigMat(2,start + 3) = pr.pval(5) / pr.pval(0)
                If bigMat(2,start + 3) = 0 Then
                    bigMat(2,start + 3) = Empty
                End If
                bigMat(3,start + 3) = pr.pval(6) / pr.pval(0)
                If bigMat(3,start + 3) = 0 Then
                    bigMat(3,start + 3) = Empty
                End If
            Case FET_L_SPRING, FET_L_SPRING_TO_GROUND
                If pr.flag(3) Then
                    bigMat(1,start) = "PBUSH"
                    leng = 28
                    propCase = Array(Empty,36,37,38,39,40,41,Empty,Empty,42,43,44,45,46,47,Empty,Empty,10,11,12,13,14,15,Empty,Empty,4,6,5,7)
                    row = pValPrint(bigMat,start,leng,pr,True,propCase)
                    If rowEmpty(start,3,bigMat) Then
                        bigMat(3,start) = Empty
                    Else
                        bigMat(3,start) = "K"
                    End If
                    If rowEmpty(start + 1,3,bigMat) Then
                        bigMat(3,start + 1) = Empty
                    Else
                        bigMat(3,start + 1) = "B"
                    End If
                    If rowEmpty(start + 2,3,bigMat) Then
                        bigMat(3,start + 2) = Empty
                    Else
                        bigMat(3,start + 2) = "GE"
                    End If
                    If rowEmpty(start + 3,3,bigMat) Then
                        bigMat(3,start + 3) = Empty
                    Else
                        bigMat(3,start + 3) = "RCV"
                    End If
                Else
                    If pr.pval(0) = 0 And pr.pval(1) <> 0 Then
                        bigMat(1,start) = "PVISC"
                        bigMat(2,start) = prID
                        If pr.flag(1) Then
                            bigMat(3,start) = pr.pval(1)
                        Else
                            bigMat(4,start) = pr.pval(1)
                        End If
                        row = 1
                    Else
                        GoTo Rod
                    End If
                End If
            Case FET_L_DOF_SPRING, FET_L_DOF_SPRING_TO_GROUND
                If pr.pval(0) = 0 And pr.pval(1) <> 0 Then
                    bigMat(1,start) = "PDAMP"
                    leng  = 8
                    propCase = Array(1,Empty,Empty,Empty,Empty,Empty,Empty,Empty,23)
                    row = pValPrint(bigMat,start,leng,pr,True,propCase)
                    If Not rowEmpty(start + 1,3,bigMat) Then
                        bigMat(1,start + 1) = "PDAMPT"
                        bigMat(2, start + 1) = pr.ID
                    End If
                Else
                    bigMat(1,start) = "PELAS"
                    leng = 10
                    propCase = Array(0,1,2,Empty,Empty,Empty,Empty,Empty,20,21,22)
                    row = pValPrint(bigMat,start,leng,pr,True,propCase)
                    If bigMat(5,start) = 0 Then
                        bigMat(5,start) = Empty
                    End If
                    If Not rowEmpty(start + 1,3,bigMat) Then
                        bigMat(1,start + 1) = "PELAST"
                        bigMat(2, start + 1) = pr.ID
                    End If
                End If
            Case FET_L_CURVED_BEAM
                bigMat(1,start) = "PBEND"
                leng = 19
                propCase = Array(0,1,2,4,17,Empty,8,9,10,11,12,13,14,15,5,6,7,Empty,Empty,Empty)
                row = pValPrint(bigMat,start,leng,pr,False,propCase)
                bigMat(2,start + 2) = pr.pval(5) / pr.pval(0)
                bigMat(3,start + 2) = pr.pval(6) / pr.pval(0)
            Case FET_L_CURVED_TUBE
                bigMat(1,start) = "PBEND"
                leng = 19
                propCase = Array(Empty, Empty, Empty, 20, 17,Empty,23,24,7,21,22,25,Empty,Empty,26,27,28,Empty,29,30)
                row = pValPrint(bigMat,start,leng,pr,False,propCase)
                bigMat(4,start) = pr.flagI(0)
                bigMat(6,start) = (pr.pval(0) - pr.pval(1))/2
                bigMat(5,start) = (pr.pval(1) + bigMat(6,start))/2
                rowEmpty(start + 1,2,bigMat)
                rowEmpty(start + 2,2,bigMat)
                If bigMat(7,start) =  0 Then
                    bigMat(7,start) = Empty
                End If
                If bigMat(9,start) =  0 Then
                    bigMat(9,start) = Empty
                End If
            Case FET_L_GAP
                bigMat(1,start) = "PGAP"
                leng = 9
                propCase = Array(0,6,1,2,3,4,5,11,12,13)
                row = pValPrint(bigMat,start,leng,pr,True,propCase)
                rowEmpty(start + 1,2,bigMat)
            Case FET_L_SHEAR
                bigMat(1,start) = "PSHEAR"
                leng = 3
                propCase = Array(0,7,8,9)
                row = pValPrint(bigMat,start,leng,pr,False,propCase)
            Case FET_L_MEMBRANE To FET_P_PLATE
                bigMat(1,start) = "PSHELL"
                leng = 8
                propCase = Array(0,1,10,2,11,7,9,8,3)
                row = pValPrint(bigMat,start,leng,pr,False,propCase)
                Select Case typeID
                Case FET_L_MEMBRANE, FET_P_MEMBRANE
                    bigMat(3,start) = pr.matlID
                    bigMat(5,start) = Empty
                    bigMat(7,start) = Empty
                    bigMat(4,start + 1) = Empty
                Case FET_L_BENDING, FET_P_BENDING
                    bigMat(3,start) = Empty
                    bigMat(5,start) = pr.matlID
                    bigMat(7,start) = Empty
                    bigMat(4,start + 1) = Empty
                Case FET_L_PLATE, FET_P_PLATE
                    bigMat(3,start) = pr.matlID
                    bigMat(5,start) = pr.matlID
                    bigMat(7,start) = pr.matlID
                    bigMat(4,start + 1) = Empty
                End Select
                If bigMat(8,start) = 0 Then
                    bigMat(8,start) = Empty
                End If
                If bigMat(6,start) = 0 Then
                    bigMat(6,start) = Empty
                End If
                rowEmpty(start + 1,2,bigMat)
            Case FET_L_PLANE_STRAIN, FET_P_PLANE_STRAIN
                mat.Get(pr.matlID)
                If mat.type = FMT_HYPERELASTIC Then
                    bigMat(1,start) = "PLPLANE"
                    leng = 2
                    propCase = Array(Empty,Empty,0)
                    row = pValPrint(bigMat,start,leng,pr,False,propCase)
                    bigMat(4,start) = pr.refCS
                    If bigMat(4,start) =  0 Then
                        bigMat(4,start) = Empty
                    End If
                    If bigMat(5,start) =  0 Then
                        bigMat(5,start) = Empty
                    End If
                Else
                    isPlane(prID,pShell,pPlane)
                    If pShell Then
                        bigMat(1,start) = "PSHELL"
                        leng = 8
                        propCase = Array(0,1,10,2,11,7,9,8,3)
                        row = pValPrint(bigMat,start,leng,pr,False,propCase)
                        bigMat(3,start) = pr.matlID
                        bigMat(5,start) = -1
                        bigMat(7,start) = Empty
                        bigMat(4,start + 1) = Empty
                        If bigMat(8,start) = 0 Then
                        bigMat(8,start) = Empty
                        End If
                        If bigMat(6,start) = 0 Then
                            bigMat(6,start) = Empty
                        End If
                        rowEmpty(start + 1,2,bigMat)
                        If pPlane Then
                            start = start + row
                            GoTo bothPlane
                        End If
                    Else
                        bothPlane:
                        bigMat(1,start) = "PPLANE"
                        leng = 2
                        propCase = Array(0,7,20)
                        row = pValPrint(bigMat,start,leng,pr,False,propCase)
                        If bigMat(6,start) =  0 Then
                            bigMat(6,start) = Empty
                        End If
                    End If
                End If
            Case FET_L_LAMINATE_PLATE, FET_P_LAMINATE_PLATE
                gSet.Clear
                gSet.AddAll(FT_LAYUP_GLOBALPLY)
                lUp.Get(pr.layupID)
                If(gSet.Count = lUp.NumberOfPlys) Then
                    leng = 6 + (8 * lUp.NumberOfPlys)
                    glob = True
                Else
                    leng = 6 + (4 * lUp.NumberOfPlys)
                    glob = False
                End If
                ReDim propCase(leng)
                propCase(1) = 1
                propCase(2) = 2
                propCase(4) = 3
                propCase(5) = 4
                row = pValPrint(bigMat,start,leng,pr,True,propCase)
                Select Case pr.flagI(0)
                Case 0
                    bigMat(6, start) = Empty
                Case 1
                    bigMat(6, start) = "HILL"
                Case 2
                    bigMat(6, start) = "HOFF"
                Case 3
                    bigMat(6, start) = "TSAI"
                Case 4
                    bigMat(6, start) = "STRN"
                End Select
                bigMat(9,start) = pr.flagI(1)
                bigMat(3,start) = pr.flagI(2)
                If glob Then
                    bigMat(1,start) = "PCOMPG"
                    For i = 0  To (lUp.NumberOfPlys - 1)
                        bigMat(2, i +1 + start) = lUp.globalply(i)
                        bigMat(3,i + 1 + start) = lUp.matlID(i)
                        bigMat(4,i + 1 + start) = lUp.thickness(i)
                        bigMat(5,i + 1 + start) = lUp.angle(i)
                        bigMat(6,i + 1 + start) = "YES"
                    Next
                Else
                    bigMat(1,start) = "PCOMP"
                    k = 0
                    For i = 0 To lUp.NumberOfPlys Step 2
                        If lUp.NumberOfPlys > 0 Then
                            bigMat(2,start +(i/2) + 1) = lUp.matlID(k)
                            bigMat(3,start +(i/2) + 1) = lUp.thickness(k)
                            bigMat(4,start +(i/2) + 1) = lUp.angle(k)
                            bigMat(5,start +(i/2) + 1) = "YES"
                            k = k + 1
                        End If
                        If (lUp.NumberOfPlys - k) > 0 Then
                            bigMat(6,start +(i/2) + 1) = lUp.matlID(k)
                            bigMat(7,start +(i/2) + 1) = lUp.thickness(k)
                            bigMat(8,start +(i/2) + 1) = lUp.angle(k)
                            bigMat(9,start +(i/2) + 1) = "YES"
                            k = k +1
                        End If
                    Next
                    If bigMat(3,start) =  0 Then
                        bigMat(3,start) = Empty
                    End If
                    If bigMat(5,start) =  0 Then
                        bigMat(5,start) = Empty
                    End If
                    If bigMat(7,start) =  0 Then
                        bigMat(7,start) = Empty
                    End If
                    If bigMat(8,start) =  0 Then
                        bigMat(8,start) = Empty
                    End If
                    If bigMat(9,start) =  0 Then
                        bigMat(9,start) = Empty
                    End If
                End If
                Case FET_L_AXISYM_SHELL, FET_P_AXISYM_SHELL
                    row = 1
                    bigMat(1,start) = "PLPLANE"
                    bigMat(2,start) = prID
                    bigMat(3,start) = pr.matlID
                    bigMat(4,start) = pr.refCS
                Case FET_L_SOLID, FET_P_SOLID,FET_L_AXISYM, FET_P_AXISYM
                    row = 1
                    mat.Get(pr.matlID)
                    If mat.type = FMT_HYPERELASTIC Then
                        bigMat(1,start) = "PLSOLID"
                        bigMat(2,start) = prID
                        bigMat(3,start) = pr.matlID
                        row = 1
                    Else
                        bigMat(1,start) = "PSOLID"
                        bigMat(2,start) = prID
                        bigMat(3,start) = pr.matlID
                        bigMat(4,start) = pr.refCS
                        bigMat(6,start) = pr.pval(0)
                        If mat.type = FMT_GENERAL Then
                            If mat.subtype = 506 Or mat.subtype = 606 Then
                                bigMat(8,start) = "PFLUID"
                            End If
                        End If
                        If bigMat(6,start) =  0 Then
                            bigMat(6,start) = Empty
                        End If
                    End If
            Case FET_L_SLIDE_LINE
                leng = 2
                propCase = Array(Empty, 1, 4)
                row = pValPrint(bigMat,start,leng,pr,True,propCase)
                If Not rowEmpty(start,3,bigMat) Then
                    bigMat(1,start) = "BFRIC"
                    start = start + row
                Else
                    bigMat(2,start) = Empty
                End If
                If pr.pval(10) <> 0 Then
                    bigMat(1,start) = "BWIDTH"
                    master = slidePropCount(prID)
                    bigMat(2,start) = master
                    bigMat(3,start) = pr.pval(10)
                    start = start + 1
                End If
                If pr.pval(11) <> 0 Then
                    bigMat(1,start) = "BWIDTH"
                    bigMat(2,start) = master + 1
                    bigMat(3,start) = pr.pval(11)
                    start = start + 1
                End If
                row = 1
                bigMat(1,start) = "BOUTPUT"
                bigMat(2,start) = master
                bigMat(3,start) = "ALL"
            Case FET_L_WELD
                If pr.pval(9) = 0 Then
                    row = 2
                    bigMat(1,start) = "PWELD"
                    bigMat(2,start) = prID
                    bigMat(3,start) = pr.matlID
                    bigMat(4,start) = pr.pval(0)
                    If pr.flag(0) Then
                        bigMat(7,start) = "ON"
                    Else
                        bigMat(7,start) = "OFF"
                    End If
                    If pr.flag(1) Then
                        bigMat(9,start) = "SPOT"
                    End If
                Else
                    bigMat(1,start) = "PFAST"
                    leng = 10
                    propCase = Array(0,Empty,Empty,3,4,5,6,7,8,1,2)
                    row = pValPrint(bigMat,start,leng,pr,True,propCase)
                    If pr.refCS=0 Then
                        bigMat(4,start) = -1
                    Else
                        bigMat(4,start) =pr.refCS
                    End If
                    bigMat(5,start) = pr.flagI(3)
                End If
            Case FET_L_LAMINATE_SOLID, FET_P_LAMINATE_SOLID
                bigMat(1,start) = "PCOMPS"
                lUp.Get(pr.layupID)
                leng = 6 + (8 * lUp.NumberOfPlys)
                ReDim propCase(leng)
                propCase(1) = 6
                propCase(2) = 2
                propCase(3) = 5
                propCase(4) = 3
                propCase(5) = 4
                row = pValPrint(bigMat,start,leng,pr,True,propCase)
                Select Case bigMat(4,start)
                Case 0
                    bigMat(4,start) = Empty
                Case 2
                    bigMat(4,start) = 13
                Case 1
                    bigMat(4,start) = 12
                Case 3
                    bigMat(4,start) = 21
                Case 4
                    bigMat(4,start) = 23
                Case 5
                    bigMat(4,start) = 31
                Case 6
                    bigMat(4,start) = 32
                End Select
                bigMat(3,start) = pr.refCS
                For i = 0 To (lUp.NumberOfPlys - 1)
                    bigMat(2 ,i + 1 + start) = lUp.globalply(i)
                    bigMat(3 ,i + 1 + start) = lUp.matlID(i)
                    bigMat(4 ,i + 1 + start) = lUp.thickness(i)
                    bigMat(5 ,i + 1 + start) = lUp.angle(i)
                    mat.Get(lUp.matlID(i))
                    bigMat(6, start +i +1) = mat.ival(11)
                    Select Case bigMat(6, start +i +1)
                    Case 0
                        bigMat(6, start +i +1) = Empty
                    Case 1
                        bigMat(6, start +i +1) = "HILL"
                    Case 2
                        bigMat(6, start +i +1) = "HOFF"
                    Case 3
                        bigMat(6, start +i +1) = "TSAI"
                    Case 4
                        bigMat(6, start +i +1) = "STRN"
                    Case 5
                        bigMat(6, start +i +1) = "STRS"
                    Case 6
                        bigMat(6, start +i +1) = "TR"
                    End Select
                    bigMat(7, start +i +1) = mat.ival(12)
                    Select Case bigMat(7, start +i +1)
                    Case 0
                        bigMat(7, start +i +1) = Empty
                    Case 1
                        bigMat(7, start +i +1) = "SB"
                    Case 2
                        bigMat(7, start +i +1) = "NB"
                    End Select
                    bigMat(8, start +i +1) = "YES"
                Next
                If bigMat(5,start) = 0 Then
                    bigMat(5,start) = Empty
                End If
                If bigMat(6,start) = 0 Then
                    bigMat(6,start) = Empty
                End If
            Case Else
                GoTo pExit
            End Select
            pExit:
            start = start + row
        Wend
    End If

    If dlg.mtbox Or dlg.matElem Then
        While mSet.Next
            mtID = mSet.Current
            mat.Get(mtID)
            mtType = mat.type
            Select Case mtType
            Case    FMT_ISOTROPIC

                If isSolLam(mtID) And mat.ival(11) <> 0 Then
                    Select Case    mat.ival(11)
                    Case 1, 2, 3, 5
                        bigMat(1,start) = "MATFT"
                        bigMat(2,start) = mtID
                        bigMat(3,start + 1) = mat.mval(128)
                        bigMat(4,start + 1) = mat.mval(125)
                        bigMat(5,start + 1) = mat.mval(129)
                        bigMat(6,start + 1) = mat.mval(126)
                        bigMat(7,start + 1) = mat.mval(130)
                        bigMat(8,start + 1) = mat.mval(127)
                        bigMat(9,start + 1) = mat.mval(131)
                        bigMat(3,start + 2) = mat.mval(133)
                        bigMat(4,start + 2) = mat.mval(132)
                        Select Case mat.ival(11)
                        Case 1
                            bigMat(2,start + 1) = "HILL"
                        Case 2
                            bigMat(2,start + 1) = "HOFF"
                        Case 3
                            bigMat(2,start + 1) = "TSAI"
                            bigMat(5,start + 2) = mat.mval(143)
                            bigMat(6,start + 2) = mat.mval(145)
                            bigMat(7,start + 2) = mat.mval(144)
                        Case 5
                            bigMat(2,start + 1) = "STRS"
                        End Select
                        rowEmpty(start + 1,2,bigMat)
                        rowEmpty(start + 2,2,bigMat)
                        start = start + 3
                    Case 4
                        bigMat(1,start) = "MATFT"
                        bigMat(2,start) = mtID
                        bigMat(2,start + 1) = "STRN"
                        bigMat(3,start + 1) = mat.mval(137)
                        bigMat(4,start + 1) = mat.mval(134)
                        bigMat(5,start + 1) = mat.mval(138)
                        bigMat(6,start + 1) = mat.mval(135)
                        bigMat(7,start + 1) = mat.mval(139)
                        bigMat(8,start + 1) = mat.mval(136)
                        bigMat(9,start + 1) = mat.mval(140)
                        bigMat(3,start + 2) = mat.mval(142)
                        bigMat(4,start + 2) = mat.mval(141)
                        rowEmpty(start + 1,2,bigMat)
                        rowEmpty(start + 2,2,bigMat)
                        start = start + 3
                    Case 6
                        bigMat(2,start + 1) = "TS"
                        bigMat(3,start + 1) = mat.mval(133)
                        bigMat(4,start + 1) = mat.mval(132)
                        rowEmpty(start + 1,2,bigMat)
                        start = start + 2
                    End Select
                End If

                If mat.ival(2) <> 0 Then
                    bigMat(1,start) = "MATS1"
                    bigMat(2,start) = mtID
                    Select Case mat.ival(2)
                    Case 1
                        If mat.fval(5) <> 0 Then
                            bigMat(4,start) = "NLELAST"
                            bigMat(3,start) = mat.fval(5)
                        End If
                    Case 2
                        bigMat(4,start) = "PLASTIC"
                        If mat.fval(5) <> 0 Then
                            bigMat(3,start) = mat.fval(5)
                        End If
                        bigMat(6,start) = mat.ival(4) + 1
                        bigMat(7,start) = mat.ival(3) + 1
                        bigMat(8,start) = mat.mval(85)
                        If mat.ival(4) <> 0 And mat.ival(11) <> 1 Then
                            bigMat(9,start) = mat.mval(86)
                        End If
                    Case 3
                        bigMat(4,start) = "PLASTIC"
                        bigMat(5,start) = mat.mval(84)
                        bigMat(6,start) = mat.ival(4) + 1
                        bigMat(7,start) = mat.ival(3) + 1
                        bigMat(8,start) = mat.mval(85)
                        If mat.ival(4) <> 0 And mat.ival(11) <> 1 Then
                            bigMat(9,start) = mat.mval(86)
                        End If
                    End Select
                    start = start + 1
                    If mat.fval(5) <> 0 And (mat.ival(2) = 1 Or mat.ival(2) = 2) Then
                        func.Get(mat.fval(5))
                        'if func.type =
                        bigMat(1,start) = "TABLES1"
                        bigMat(2,start) = func.ID
                        func.GetFunctionList(cnt,x,y)
                        start = start + 1
                        For i = 0 To (cnt - 1)
                            bigMat(((i * 2) Mod 8) + 2,(i \ 4) + start) = x(i)
                            bigMat((((i * 2) + 1) Mod 8) + 2,((i) \ 4) + start) = y(i)
                        Next
                        bigMat((((i * 2)) Mod 8) + 2,((i) \ 4) + start) = "ENDT"
                        start = start + (i \ 4) + 1
                    End If
                End If

                If mat.ival(5) <> 0 Then
                    bigMat(1,start) = "CREEP"
                    bigMat(2,start) = mtID
                    bigMat(3,start) = mat.mval(91)
                    bigMat(4,start) = mat.mval(92)
                    Select Case mat.ival(5)
                    Case 1
                        bigMat(5,start) = "CRLAW"
                    Case 2
                        bigMat(5,start) = "TABLE"
                        For i = 6 To 8
                            If mat.fval(i) Then
                                bigMat(i,start) = mat.fval(i)
                            End If
                        Next
                    End Select
                    bigMat(9,start) = mat.mval(92)
                    start = start + 1

                    bigMat(2,start) = ((mat.ival(6) - 1) * (-189)) + ((mat.ival(7) + 1) * 100) + ((mat.ival(8) + 1) * 10) + (mat.ival(9) + 1)
                    For i = 3 To 9
                        bigMat(i,start) = mat.mval(i + 90)
                    Next
                    start = start + 1
                End If

                bigMat(1,start) = "MAT1"
                bigMat(2,start) = mtID
                bigMat(3,start) = mat.mval(0)
                If mat.mval(3) <> 0 Then
                    bigMat(4,start) = mat.mval(3)
                End If
                If mat.mval(6) <> 0 Then
                    bigMat(5,start) = mat.mval(6)
                End If
                If mat.mval(49) <> 0 Then
                    bigMat(6,start) = mat.mval(49)
                End If
                bigMat(7,start) = mat.mval(36)
                bigMat(8,start) = mat.mval(51)
                If mat.mval(50) <> 0 Then
                    bigMat(9,start) = mat.mval(50)
                End If
                start = start + 1
                bigMat(2,start) = mat.mval(52)
                bigMat(3,start) = mat.mval(54)
                bigMat(4,start) = mat.mval(56)
                If Not rowEmpty(start,2,bigMat) Then
                    start = start + 1
                End If

                bigMat(3,start) = mat.tval(0)
                bigMat(4,start) = mat.tval(3)
                bigMat(5,start) = mat.tval(6)
                bigMat(6,start) = mat.tval(49)
                bigMat(7,start) = mat.tval(36)
                bigMat(9,start) = mat.tval(50)
                bigMat(2,start + 1) = mat.tval(52)
                bigMat(3,start + 1) = mat.tval(54)
                bigMat(4,start + 1) = mat.tval(56)
                If Not rowEmpty(start,3,bigMat) Or Not rowEmpty(start + 1,2,bigMat) Then
                    bigMat(1,start) = "MATT1"
                    bigMat(2,start) = mtID
                    start = start + 2
                End If

                bigMat(1,start) = "MAT4"
                bigMat(2,start) = mtID
                bigMat(3,start) = mat.mval(42)
                bigMat(4,start) = mat.mval(48)
                bigMat(5,start) = mat.mval(49)
                bigMat(8,start) = mat.mval(100)
                bigMat(9,start) = mat.mval(101)
                rowEmpty(start,2,bigMat)
                start = start + 1
                bigMat(2,start) = mat.mval(102)
                bigMat(3,start) = mat.mval(103)
                bigMat(4,start) = mat.mval(104)
                If Not rowEmpty(start,2,bigMat) Then
                    start = start + 1
                End If

                bigMat(3,start) = mat.tval(42)
                bigMat(4,start) = mat.tval(48)
                bigMat(8,start) = mat.tval(53)
                If Not rowEmpty(start,3,bigMat) Then
                    bigMat(1,start) = "MATT4"
                    bigMat(2,start) = mtID
                    start = start + 1
                End If

                For i = 0 To 69
                    If mat.tval(i) <> 0 Then
                        fSet.Add(mat.tval(i))
                    End If
                Next
                While fSet.Next
                    func.Get(fSet.CurrentID)
                    If anType = 5 Then
                        bigMat(1,start) = "TABLEM1"
                    Else
                        bigMat(1,start) = "TABLEM2"
                    End If
                    bigMat(2,start) = func.ID
                    start = start + 1
                    func.GetFunctionList(cnt,x,y)
                    For i = 0 To (cnt - 1)
                        bigMat(((i * 2) Mod 8) + 2,(i \ 4) + start) = x(i)
                        bigMat((((i * 2) + 1) Mod 8) + 2,((i) \ 4) + start) = y(i)
                    Next
                    bigMat((((i * 2)) Mod 8) + 2,((i) \ 4) + start) = "ENDT"
                    start = start + (i \ 4) + 1
                Wend
                fSet.Clear

            Case FMT_ORTHOTROPIC_2D

                bigMat(1,start) = "MAT8"
                bigMat(2,start) = mtID
                bigMat(3,start) = mat.mval(0)
                bigMat(4,start) = mat.mval(1)
                bigMat(5,start) = mat.mval(6)
                bigMat(6,start) = mat.mval(3)
                bigMat(7,start) = mat.mval(4)
                bigMat(8,start) = mat.mval(5)
                bigMat(9,start) = mat.mval(49)
                start = start + 1
                bigMat(2,start) = mat.mval(36)
                bigMat(3,start) = mat.mval(37)
                bigMat(4,start) = mat.mval(51)
                bigMat(5,start) = mat.mval(2)
                bigMat(6,start) = mat.mval(54)
                bigMat(7,start) = mat.mval(53)
                bigMat(8,start) = mat.mval(55)
                bigMat(9,start) = mat.mval(56)
                rowEmpty(start,5,bigMat)
                start = start + 1
                bigMat(2,start) = mat.mval(50)
                bigMat(3,start) = mat.mval(57)
                bigMat(4,start) = -mat.bVal(0)
                rowEmpty(start,2,bigMat)
                start = start + 1

                bigMat(3,start) = mat.tval(0)
                bigMat(4,start) = mat.tval(1)
                bigMat(5,start) = mat.tval(6)
                bigMat(6,start) = mat.tval(3)
                bigMat(7,start) = mat.tval(4)
                bigMat(8,start) = mat.tval(5)
                bigMat(9,start) = mat.tval(49)
                bigMat(2,start + 1) = mat.tval(36)
                bigMat(3,start + 1) = mat.tval(39)
                bigMat(5,start + 1) = mat.tval(52)
                bigMat(6,start + 1) = mat.tval(53)
                bigMat(7,start + 1) = mat.tval(54)
                bigMat(8,start + 1) = mat.tval(55)
                bigMat(9,start + 1) = mat.tval(56)
                bigMat(2,start + 2) = mat.tval(50)
                If Not rowEmpty(start,3,bigMat) Or Not rowEmpty(start + 1,2,bigMat) Or Not rowEmpty(start + 2,2,bigMat) Then
                    bigMat(1,start) = "MATT8"
                    bigMat(2,start) = mtID
                    start = start + 3
                End If

                bigMat(1,start) = "MAT5"
                bigMat(2,start) = mtID
                bigMat(3,start) = mat.mval(42)
                bigMat(4,start) = mat.mval(43)
                bigMat(5,start) = mat.mval(44)
                bigMat(6,start) = mat.mval(45)
                bigMat(7,start) = mat.mval(46)
                bigMat(8,start) = mat.mval(47)
                bigMat(9,start) = mat.mval(48)
                rowEmpty(start,2,bigMat)
                start = start + 1
                bigMat(2,start) = mat.mval(49)
                bigMat(3,start) = mat.mval(100)
                If Not rowEmpty(start,2,bigMat) Then
                    start = start + 1
                End If

                bigMat(3,start) = mat.tval(42)
                bigMat(4,start) = mat.tval(43)
                bigMat(5,start) = mat.tval(44)
                bigMat(6,start) = mat.tval(45)
                bigMat(7,start) = mat.tval(46)
                bigMat(8,start) = mat.tval(47)
                bigMat(9,start) = mat.tval(48)
                bigMat(3,start + 1) = mat.tval(53)
                If Not rowEmpty(start,3,bigMat) Or Not rowEmpty(start + 1,2,bigMat) Then
                    bigMat(1,start) = "MATT5"
                    bigMat(2,start) = mtID
                    start = start + 2
                End If

                For i = 0 To 69
                    If mat.tval(i) <> 0 Then
                        fSet.Add(mat.tval(i))
                    End If
                Next
                While fSet.Next
                    func.Get(fSet.CurrentID)
                    If anType = 5 Then
                        bigMat(1,start) = "TABLEM1"
                    Else
                        bigMat(1,start) = "TABLEM2"
                    End If
                    bigMat(2,start) = func.ID
                    start = start + 1
                    func.GetFunctionList(cnt,x,y)
                    For i = 0 To (cnt - 1)
                        bigMat(((i * 2) Mod 8) + 2,(i \ 4) + start) = x(i)
                        bigMat((((i * 2) + 1) Mod 8) + 2,((i) \ 4) + start) = y(i)
                    Next
                    bigMat((((i * 2)) Mod 8) + 2,((i) \ 4) + start) = "ENDT"
                    start = start + (i \ 4) + 1
                Wend
                fSet.Clear

            Case FMT_ORTHOTROPIC_3D

                If isSolLam(mtID) And mat.ival(11) <> 0 Then
                    Select Case    mat.ival(11)
                    Case 1, 2, 3, 5
                        bigMat(1,start) = "MATFT"
                        bigMat(2,start) = mtID
                        bigMat(3,start + 1) = mat.mval(128)
                        bigMat(4,start + 1) = mat.mval(125)
                        bigMat(5,start + 1) = mat.mval(129)
                        bigMat(6,start + 1) = mat.mval(126)
                        bigMat(7,start + 1) = mat.mval(130)
                        bigMat(8,start + 1) = mat.mval(127)
                        bigMat(9,start + 1) = mat.mval(131)
                        bigMat(3,start + 2) = mat.mval(133)
                        bigMat(4,start + 2) = mat.mval(132)
                        Select Case mat.ival(11)
                        Case 1
                            bigMat(2,start + 1) = "HILL"
                        Case 2
                            bigMat(2,start + 1) = "HOFF"
                        Case 3
                            bigMat(2,start + 1) = "TSAI"
                            bigMat(5,start + 2) = mat.mval(143)
                            bigMat(6,start + 2) = mat.mval(145)
                            bigMat(7,start + 2) = mat.mval(144)
                        Case 5
                            bigMat(2,start + 1) = "STRS"
                        End Select
                        rowEmpty(start + 1,2,bigMat)
                        rowEmpty(start + 2,2,bigMat)
                        start = start + 3
                    Case 4
                        bigMat(1,start) = "MATFT"
                        bigMat(2,start) = mtID
                        bigMat(2,start + 1) = "STRN"
                        bigMat(3,start + 1) = mat.mval(137)
                        bigMat(4,start + 1) = mat.mval(134)
                        bigMat(5,start + 1) = mat.mval(138)
                        bigMat(6,start + 1) = mat.mval(135)
                        bigMat(7,start + 1) = mat.mval(139)
                        bigMat(8,start + 1) = mat.mval(136)
                        bigMat(9,start + 1) = mat.mval(140)
                        bigMat(3,start + 2) = mat.mval(142)
                        bigMat(4,start + 2) = mat.mval(141)
                        rowEmpty(start + 1,2,bigMat)
                        rowEmpty(start + 2,2,bigMat)
                        start = start + 3
                    Case 6
                        bigMat(2,start + 1) = "TS"
                        bigMat(3,start + 1) = mat.mval(133)
                        bigMat(4,start + 1) = mat.mval(132)
                        rowEmpty(start + 1,2,bigMat)
                        start = start + 2
                    End Select
                End If

                If isSolid(mtID) Then
                    bigMat(1,start) = "MAT11"
                    bigMat(2,start) = mtID
                    bigMat(3,start) = mat.mval(0)
                    bigMat(4,start) = mat.mval(1)
                    bigMat(5,start) = mat.mval(2)
                    bigMat(6,start) = mat.mval(6)
                    bigMat(7,start) = mat.mval(8)
                    bigMat(8,start) = mat.mval(7)
                    bigMat(9,start) = mat.mval(3)
                    bigMat(2,start + 1) = mat.mval(4)
                    bigMat(3,start + 1) = mat.mval(5)
                    bigMat(4,start + 1) = mat.mval(49)
                    bigMat(5,start + 1) = mat.mval(36)
                    bigMat(6,start + 1) = mat.mval(39)
                    bigMat(7,start + 1) = mat.mval(41)
                    bigMat(8,start + 1) = mat.mval(51)

                    bigMat(9,start + 1) = mat.mval(50)
                    start = start + 2

                    bigMat(3,start) = mat.tval(0)
                    bigMat(4,start) = mat.tval(1)
                    bigMat(5,start) = mat.tval(2)
                    bigMat(6,start) = mat.tval(6)
                    bigMat(7,start) = mat.tval(8)
                    bigMat(8,start) = mat.tval(7)
                    bigMat(9,start) = mat.tval(3)
                    bigMat(2,start + 1) = mat.tval(5)
                    bigMat(3,start + 1) = mat.tval(4)
                    bigMat(4,start + 1) = mat.tval(49)
                    bigMat(5,start + 1) = mat.tval(36)
                    bigMat(6,start + 1) = mat.tval(39)
                    bigMat(7,start + 1) = mat.tval(41)
                    bigMat(9,start + 1) = mat.tval(50)
                    If Not rowEmpty(start,3,bigMat) Or Not rowEmpty(start + 1,2,bigMat) Then
                        bigMat(1,start) = "MATT11"
                        bigMat(2,start) = mtID
                        start = start + 2
                    End If
                Else
                    bigMat(1,start) = "MAT3"
                    bigMat(2,start) = mtID
                    bigMat(3,start) = mat.mval(0)
                    bigMat(4,start) = mat.mval(1)
                    bigMat(5,start) = mat.mval(2)
                    bigMat(6,start) = mat.mval(6)
                    bigMat(7,start) = mat.mval(7)
                    bigMat(8,start) = mat.mval(8)
                    bigMat(9,start) = mat.mval(49)
                    bigMat(4,start + 1) = mat.mval(56)
                    bigMat(5,start + 1) = mat.mval(36)
                    bigMat(6,start + 1) = mat.mval(37)
                    bigMat(7,start + 1) = mat.mval(38)
                    bigMat(8,start + 1) = mat.mval(51)
                    bigMat(9,start + 1) = mat.mval(50)
                    start = start + 2

                    bigMat(3,start) = mat.tval(0)
                    bigMat(4,start) = mat.tval(1)
                    bigMat(5,start) = mat.tval(2)
                    bigMat(6,start) = mat.tval(6)
                    bigMat(7,start) = mat.tval(7)
                    bigMat(8,start) = mat.tval(8)
                    bigMat(9,start) = mat.tval(49)
                    bigMat(4,start + 1) = mat.tval(3)
                    bigMat(5,start + 1) = mat.tval(36)
                    bigMat(6,start + 1) = mat.tval(39)
                    bigMat(7,start + 1) = mat.tval(41)
                    bigMat(9,start + 1) = mat.tval(50)
                    If Not rowEmpty(start,3,bigMat) Or Not rowEmpty(start + 1,2,bigMat) Then
                        bigMat(1,start) = "MATT3"
                        bigMat(2,start) = mtID
                        start = start + 2
                    End If

                End If

                For i = 0 To 69
                    If mat.tval(i) <> 0 Then
                        fSet.Add(mat.tval(i))
                    End If
                Next

                bigMat(1,start) = "MAT5"
                bigMat(2,start) = mtID
                bigMat(3,start) = mat.mval(42)
                bigMat(4,start) = mat.mval(43)
                bigMat(5,start) = mat.mval(44)
                bigMat(6,start) = mat.mval(45)
                bigMat(7,start) = mat.mval(46)
                bigMat(8,start) = mat.mval(47)
                bigMat(9,start) = mat.mval(48)
                rowEmpty(start,2,bigMat)
                start = start + 1
                bigMat(2,start) = mat.mval(49)
                bigMat(3,start) = mat.mval(100)
                If Not rowEmpty(start,2,bigMat) Then
                    start = start + 1
                End If
                bigMat(3,start) = mat.tval(42)
                bigMat(4,start) = mat.tval(43)
                bigMat(5,start) = mat.tval(44)
                bigMat(6,start) = mat.tval(45)
                bigMat(7,start) = mat.tval(46)
                bigMat(8,start) = mat.tval(47)
                bigMat(9,start) = mat.tval(48)
                bigMat(3,start + 1) = mat.tval(53)
                If Not rowEmpty(start,3,bigMat) Or Not rowEmpty(start + 1,2,bigMat) Then
                    bigMat(1,start) = "MATT5"
                    bigMat(2,start) = mtID
                    start = start + 2
                End If

                While fSet.Next
                    func.Get(fSet.CurrentID)
                    If anType = 5 Then
                        bigMat(1,start) = "TABLEM1"
                    Else
                        bigMat(1,start) = "TABLEM2"
                    End If
                    bigMat(2,start) = func.ID
                    start = start + 1
                    func.GetFunctionList(cnt,x,y)
                    For i = 0 To (cnt - 1)
                        bigMat(((i * 2) Mod 8) + 2,(i \ 4) + start) = x(i)
                        bigMat((((i * 2) + 1) Mod 8) + 2,((i) \ 4) + start) = y(i)
                    Next
                    bigMat((((i * 2)) Mod 8) + 2,((i) \ 4) + start) = "ENDT"
                    start = start + (i \ 4) + 1
                Wend
                fSet.Clear

            Case FMT_ANISOTROPIC_2D

                If mat.ival(2) <> 0 Then
                    bigMat(1,start) = "MATS1"
                    bigMat(2,start) = mtID
                    Select Case mat.ival(2)
                    Case 1
                        If mat.fval(5) <> 0 Then
                            bigMat(4,start) = "NLELAST"
                            bigMat(3,start) = mat.fval(5)
                        End If
                    Case 2
                        bigMat(4,start) = "PLASTIC"
                        If mat.fval(5) <> 0 Then
                            bigMat(3,start) = mat.fval(5)
                        End If
                        bigMat(6,start) = mat.ival(4) + 1
                        bigMat(7,start) = mat.ival(3) + 1
                        bigMat(8,start) = mat.mval(85)
                        If mat.ival(4) <> 0 And mat.ival(11) <> 1 Then
                            bigMat(9,start) = mat.mval(86)
                        End If
                    Case 3
                        bigMat(4,start) = "PLASTIC"
                        bigMat(5,start) = mat.mval(84)
                        bigMat(6,start) = mat.ival(4) + 1
                        bigMat(7,start) = mat.ival(3) + 1
                        bigMat(8,start) = mat.mval(85)
                        If mat.ival(4) <> 0 And mat.ival(11) <> 1 Then
                            bigMat(9,start) = mat.mval(86)
                        End If
                    End Select
                    start = start + 1
                    If mat.fval(5) <> 0 And (mat.ival(2) = 1 Or mat.ival(2) = 2) Then
                        func.Get(mat.fval(5))
                        'if func.type =
                        bigMat(1,start) = "TABLES1"
                        bigMat(2,start) = func.ID
                        func.GetFunctionList(cnt,x,y)
                        start = start + 1
                        For i = 0 To (cnt - 1)
                            bigMat(((i * 2) Mod 8) + 2,(i \ 4) + start) = x(i)
                            bigMat((((i * 2) + 1) Mod 8) + 2,((i) \ 4) + start) = y(i)
                        Next
                        bigMat((((i * 2)) Mod 8) + 2,((i) \ 4) + start) = "ENDT"
                        start = start + (i \ 4) + 1
                    End If
                End If

                If mat.ival(5) <> 0 Then
                    bigMat(1,start) = "CREEP"
                    bigMat(2,start) = mtID
                    bigMat(3,start) = mat.mval(91)
                    bigMat(4,start) = mat.mval(92)
                    Select Case mat.ival(5)
                    Case 1
                        bigMat(5,start) = "CRLAW"
                    Case 2
                        bigMat(5,start) = "TABLE"
                        For i = 6 To 8
                            If mat.fval(i) Then
                                bigMat(i,start) = mat.fval(i)
                            End If
                        Next
                    End Select
                    bigMat(9,start) = mat.mval(92)
                    start = start + 1

                    bigMat(2,start) = ((mat.ival(6) - 1) * (-189)) + ((mat.ival(7) + 1) * 100) + ((mat.ival(8) + 1) * 10) + (mat.ival(9) + 1)
                    For i = 3 To 9
                        bigMat(i,start) = mat.mval(i + 90)
                    Next
                    start = start + 1
                End If

                bigMat(1,start) = "MAT2"
                bigMat(2,start) = mtID
                bigMat(3,start) = mat.mval(30)
                bigMat(4,start) = mat.mval(31)
                bigMat(5,start) = mat.mval(32)
                bigMat(6,start) = mat.mval(33)
                bigMat(7,start) = mat.mval(34)
                bigMat(8,start) = mat.mval(35)
                bigMat(9,start) = mat.mval(49)
                bigMat(2,start + 1) = mat.mval(36)
                bigMat(3,start + 1) = mat.mval(39)
                bigMat(4,start + 1) = mat.mval(37)
                bigMat(5,start + 1) = mat.mval(51)
                bigMat(6,start + 1) = mat.mval(50)
                bigMat(7,start + 1) = mat.mval(52)
                bigMat(8,start + 1) = mat.mval(54)
                bigMat(9,start + 1) = mat.mval(56)
                rowEmpty(start + 1, 6,bigMat)
                start = start + 2

                bigMat(3,start) = mat.tval(30)
                bigMat(4,start) = mat.tval(31)
                bigMat(5,start) = mat.tval(32)
                bigMat(6,start) = mat.tval(33)
                bigMat(7,start) = mat.tval(34)
                bigMat(8,start) = mat.tval(35)
                bigMat(9,start) = mat.tval(49)
                bigMat(2,start + 1) = mat.tval(36)
                bigMat(3,start + 1) = mat.tval(39)
                bigMat(4,start + 1) = mat.tval(37)
                bigMat(5,start + 1) = mat.tval(51)
                bigMat(6,start + 1) = mat.tval(50)
                bigMat(7,start + 1) = mat.tval(52)
                bigMat(8,start + 1) = mat.tval(54)
                bigMat(9,start + 1) = mat.tval(56)
                If Not rowEmpty(start,3,bigMat) Or Not rowEmpty(start + 1,2,bigMat) Then
                    bigMat(1,start) = "MATT2"
                    bigMat(2,start) = mtID
                    start = start + 2
                End If

                bigMat(1,start) = "MAT5"
                bigMat(2,start) = mtID
                bigMat(3,start) = mat.mval(42)
                bigMat(4,start) = mat.mval(43)
                bigMat(5,start) = mat.mval(44)
                bigMat(6,start) = mat.mval(45)
                bigMat(7,start) = mat.mval(46)
                bigMat(8,start) = mat.mval(47)
                bigMat(9,start) = mat.mval(48)
                rowEmpty(start,2,bigMat)
                start = start + 1
                bigMat(2,start) = mat.mval(49)
                bigMat(3,start) = mat.mval(100)
                If Not rowEmpty(start,2,bigMat) Then
                    start = start + 1
                End If
                bigMat(3,start) = mat.tval(42)
                bigMat(4,start) = mat.tval(43)
                bigMat(5,start) = mat.tval(44)
                bigMat(6,start) = mat.tval(45)
                bigMat(7,start) = mat.tval(46)
                bigMat(8,start) = mat.tval(47)
                bigMat(9,start) = mat.tval(48)
                bigMat(3,start + 1) = mat.tval(53)
                If Not rowEmpty(start,3,bigMat) Or Not rowEmpty(start + 1,2,bigMat) Then
                    bigMat(1,start) = "MATT5"
                    bigMat(2,start) = mtID
                    start = start + 2
                End If

            Case FMT_ANISOTROPIC_3D

                If isSolLam(mtID) And mat.ival(11) <> 0 Then
                    Select Case    mat.ival(11)
                    Case 1, 2, 3, 5
                        bigMat(1,start) = "MATFT"
                        bigMat(2,start) = mtID
                        bigMat(3,start + 1) = mat.mval(128)
                        bigMat(4,start + 1) = mat.mval(125)
                        bigMat(5,start + 1) = mat.mval(129)
                        bigMat(6,start + 1) = mat.mval(126)
                        bigMat(7,start + 1) = mat.mval(130)
                        bigMat(8,start + 1) = mat.mval(127)
                        bigMat(9,start + 1) = mat.mval(131)
                        bigMat(3,start + 2) = mat.mval(133)
                        bigMat(4,start + 2) = mat.mval(132)
                        Select Case mat.ival(11)
                        Case 1
                            bigMat(2,start + 1) = "HILL"
                        Case 2
                            bigMat(2,start + 1) = "HOFF"
                        Case 3
                            bigMat(2,start + 1) = "TSAI"
                            bigMat(5,start + 2) = mat.mval(143)
                            bigMat(6,start + 2) = mat.mval(145)
                            bigMat(7,start + 2) = mat.mval(144)
                        Case 5
                            bigMat(2,start + 1) = "STRS"
                        End Select
                        rowEmpty(start + 1,2,bigMat)
                        rowEmpty(start + 2,2,bigMat)
                        start = start + 3
                    Case 4
                        bigMat(1,start) = "MATFT"
                        bigMat(2,start) = mtID
                        bigMat(2,start + 1) = "STRN"
                        bigMat(3,start + 1) = mat.mval(137)
                        bigMat(4,start + 1) = mat.mval(134)
                        bigMat(5,start + 1) = mat.mval(138)
                        bigMat(6,start + 1) = mat.mval(135)
                        bigMat(7,start + 1) = mat.mval(139)
                        bigMat(8,start + 1) = mat.mval(136)
                        bigMat(9,start + 1) = mat.mval(140)
                        bigMat(3,start + 2) = mat.mval(142)
                        bigMat(4,start + 2) = mat.mval(141)
                        rowEmpty(start + 1,2,bigMat)
                        rowEmpty(start + 2,2,bigMat)
                        start = start + 3
                    Case 6
                        bigMat(2,start + 1) = "TS"
                        bigMat(3,start + 1) = mat.mval(133)
                        bigMat(4,start + 1) = mat.mval(132)
                        rowEmpty(start + 1,2,bigMat)
                        start = start + 2
                    End Select
                End If

                If mat.ival(2) <> 0 Then
                    bigMat(1,start) = "MATS1"
                    bigMat(2,start) = mtID
                    Select Case mat.ival(2)
                    Case 1
                        If mat.fval(5) <> 0 Then
                            bigMat(4,start) = "NLELAST"
                            bigMat(3,start) = mat.fval(5)
                        End If
                    Case 2
                        bigMat(4,start) = "PLASTIC"
                        If mat.fval(5) <> 0 Then
                            bigMat(3,start) = mat.fval(5)
                        End If
                        bigMat(6,start) = mat.ival(4) + 1
                        bigMat(7,start) = mat.ival(3) + 1
                        bigMat(8,start) = mat.mval(85)
                        If mat.ival(4) <> 0 And mat.ival(11) <> 1 Then
                            bigMat(9,start) = mat.mval(86)
                        End If
                    Case 3
                        bigMat(4,start) = "PLASTIC"
                        bigMat(5,start) = mat.mval(84)
                        bigMat(6,start) = mat.ival(4) + 1
                        bigMat(7,start) = mat.ival(3) + 1
                        bigMat(8,start) = mat.mval(85)
                        If mat.ival(4) <> 0 And mat.ival(11) <> 1 Then
                            bigMat(9,start) = mat.mval(86)
                        End If
                    End Select
                    start = start + 1
                    If mat.fval(5) <> 0 And (mat.ival(2) = 1 Or mat.ival(2) = 2) Then
                        func.Get(mat.fval(5))
                        'if func.type =
                        bigMat(1,start) = "TABLES1"
                        bigMat(2,start) = func.ID
                        func.GetFunctionList(cnt,x,y)
                        start = start + 1
                        For i = 0 To (cnt - 1)
                            bigMat(((i * 2) Mod 8) + 2,(i \ 4) + start) = x(i)
                            bigMat((((i * 2) + 1) Mod 8) + 2,((i) \ 4) + start) = y(i)
                        Next
                        bigMat((((i * 2)) Mod 8) + 2,((i) \ 4) + start) = "ENDT"
                        start = start + (i \ 4) + 1
                    End If
                End If

                If mat.ival(5) <> 0 Then
                    bigMat(1,start) = "CREEP"
                    bigMat(2,start) = mtID
                    bigMat(3,start) = mat.mval(91)
                    bigMat(4,start) = mat.mval(92)
                    Select Case mat.ival(5)
                    Case 1
                        bigMat(5,start) = "CRLAW"
                    Case 2
                        bigMat(5,start) = "TABLE"
                        For i = 6 To 8
                            If mat.fval(i) Then
                                bigMat(i,start) = mat.fval(i)
                            End If
                        Next
                    End Select
                    bigMat(9,start) = mat.mval(92)
                    start = start + 1

                    bigMat(2,start) = ((mat.ival(6) - 1) * (-189)) + ((mat.ival(7) + 1) * 100) + ((mat.ival(8) + 1) * 10) + (mat.ival(9) + 1)
                    For i = 3 To 9
                        bigMat(i,start) = mat.mval(i + 90)
                    Next
                    start = start + 1
                End If

                bigMat(1,start) = "MAT9"
                bigMat(2,start) = mtID
                bigMat(3,start) = mat.mval(9)
                bigMat(4,start) = mat.mval(10)
                bigMat(5,start) = mat.mval(11)
                bigMat(6,start) = mat.mval(12)
                bigMat(7,start) = mat.mval(13)
                bigMat(8,start) = mat.mval(14)
                bigMat(9,start) = mat.mval(15)
                bigMat(2,start + 1) = mat.mval(16)
                bigMat(3,start + 1) = mat.mval(17)
                bigMat(4,start + 1) = mat.mval(18)
                bigMat(5,start + 1) = mat.mval(19)
                bigMat(6,start + 1) = mat.mval(20)
                bigMat(7,start + 1) = mat.mval(21)
                bigMat(8,start + 1) = mat.mval(22)
                bigMat(9,start + 1) = mat.mval(23)
                bigMat(2,start + 2) = mat.mval(24)
                bigMat(3,start + 2) = mat.mval(25)
                bigMat(4,start + 2) = mat.mval(26)
                bigMat(5,start + 2) = mat.mval(27)
                bigMat(6,start + 2) = mat.mval(28)
                bigMat(7,start + 2) = mat.mval(29)
                bigMat(8,start + 2) = mat.mval(49)
                bigMat(9,start + 2) = mat.mval(36)
                bigMat(2,start + 3) = mat.mval(39)
                bigMat(3,start + 3) = mat.mval(41)
                bigMat(4,start + 3) = mat.mval(37)
                bigMat(5,start + 3) = mat.mval(40)
                bigMat(6,start + 3) = mat.mval(38)
                bigMat(7,start + 3) = mat.mval(51)
                bigMat(8,start + 3) = mat.mval(50)
                start = start + 4

                bigMat(3,start) = mat.tval(9)
                bigMat(4,start) = mat.tval(10)
                bigMat(5,start) = mat.tval(11)
                bigMat(6,start) = mat.tval(12)
                bigMat(7,start) = mat.tval(13)
                bigMat(8,start) = mat.tval(14)
                bigMat(9,start) = mat.tval(15)
                bigMat(2,start + 1) = mat.tval(16)
                bigMat(3,start + 1) = mat.tval(17)
                bigMat(4,start + 1) = mat.tval(18)
                bigMat(5,start + 1) = mat.tval(19)
                bigMat(6,start + 1) = mat.tval(20)
                bigMat(7,start + 1) = mat.tval(21)
                bigMat(8,start + 1) = mat.tval(22)
                bigMat(9,start + 1) = mat.tval(23)
                bigMat(2,start + 2) = mat.tval(24)
                bigMat(3,start + 2) = mat.tval(25)
                bigMat(4,start + 2) = mat.tval(26)
                bigMat(5,start + 2) = mat.tval(27)
                bigMat(6,start + 2) = mat.tval(28)
                bigMat(7,start + 2) = mat.tval(29)
                bigMat(8,start + 2) = mat.tval(49)
                bigMat(9,start + 2) = mat.tval(36)
                bigMat(2,start + 3) = mat.tval(39)
                bigMat(3,start + 3) = mat.tval(41)
                bigMat(4,start + 3) = mat.tval(37)
                bigMat(5,start + 3) = mat.tval(40)
                bigMat(6,start + 3) = mat.tval(38)
                bigMat(7,start + 3) = mat.tval(51)
                bigMat(8,start + 3) = mat.tval(50)
                If Not rowEmpty(start,3,bigMat) Or Not rowEmpty(start + 1,2,bigMat) Or Not rowEmpty(start + 2,2,bigMat) Or Not rowEmpty(start + 3,2,bigMat) Then
                    bigMat(1,start) = "MATT9"
                    bigMat(2,start) = mtID
                    start = start + 4
                End If

                bigMat(1,start) = "MAT5"
                bigMat(2,start) = mtID
                bigMat(3,start) = mat.mval(42)
                bigMat(4,start) = mat.mval(43)
                bigMat(5,start) = mat.mval(44)
                bigMat(6,start) = mat.mval(45)
                bigMat(7,start) = mat.mval(46)
                bigMat(8,start) = mat.mval(47)
                bigMat(9,start) = mat.mval(48)
                rowEmpty(start,2,bigMat)
                start = start + 1
                bigMat(2,start) = mat.mval(49)
                bigMat(3,start) = mat.mval(100)
                If Not rowEmpty(start,2,bigMat) Then
                    start = start + 1
                End If
                bigMat(3,start) = mat.tval(42)
                bigMat(4,start) = mat.tval(43)
                bigMat(5,start) = mat.tval(44)
                bigMat(6,start) = mat.tval(45)
                bigMat(7,start) = mat.tval(46)
                bigMat(8,start) = mat.tval(47)
                bigMat(9,start) = mat.tval(48)
                bigMat(3,start + 1) = mat.tval(53)
                If Not rowEmpty(start,3,bigMat) Or Not rowEmpty(start + 1,2,bigMat) Then
                    bigMat(1,start) = "MATT5"
                    bigMat(2,start) = mtID
                    start = start + 2
                End If

                While fSet.Next
                    func.Get(fSet.CurrentID)
                    If anType = 5 Then
                        bigMat(1,start) = "TABLEM1"
                    Else
                        bigMat(1,start) = "TABLEM2"
                    End If
                    bigMat(2,start) = func.ID
                    start = start + 1
                    func.GetFunctionList(cnt,x,y)
                    For i = 0 To (cnt - 1)
                        bigMat(((i * 2) Mod 8) + 2,(i \ 4) + start) = x(i)
                        bigMat((((i * 2) + 1) Mod 8) + 2,((i) \ 4) + start) = y(i)
                    Next
                    bigMat((((i * 2)) Mod 8) + 2,((i) \ 4) + start) = "ENDT"
                    start = start + (i \ 4) + 1
                Wend
                fSet.Clear

            Case FMT_HYPERELASTIC
                bigMat(1,start) = "MATHP"
                bigMat(2,start) = mtID
                bigMat(3,start) = mat.mval(64)
                bigMat(4,start) = mat.mval(59)
                bigMat(5,start) = mat.mval(79)
                bigMat(6,start) = mat.mval(49)
                bigMat(7,start) = mat.mval(36)
                bigMat(8,start) = mat.mval(51)
                bigMat(9,start) = mat.mval(50)
                bigMat(3,start + 1) = mat.ival(0)
                bigMat(4,start + 1) = mat.ival(1)
                bigMat(2,start + 2) = mat.mval(69)
                bigMat(3,start + 2) = mat.mval(65)
                bigMat(4,start + 2) = mat.mval(60)
                bigMat(5,start + 2) = mat.mval(80)
                bigMat(2,start + 3) = mat.mval(73)
                bigMat(3,start + 3) = mat.mval(70)
                bigMat(4,start + 3) = mat.mval(66)
                bigMat(5,start + 3) = mat.mval(61)
                bigMat(6,start + 3) = mat.mval(81)
                bigMat(2,start + 4) = mat.mval(76)
                bigMat(3,start + 4) = mat.mval(74)
                bigMat(4,start + 4) = mat.mval(71)
                bigMat(5,start + 4) = mat.mval(67)
                bigMat(6,start + 4) = mat.mval(62)
                bigMat(7,start + 4) = mat.mval(82)
                bigMat(2,start + 5) = mat.mval(78)
                bigMat(3,start + 5) = mat.mval(77)
                bigMat(4,start + 5) = mat.mval(75)
                bigMat(5,start + 5) = mat.mval(72)
                bigMat(6,start + 5) = mat.mval(68)
                bigMat(7,start + 5) = mat.mval(63)
                bigMat(8,start + 5) = mat.mval(83)
                bigMat(2,start + 6) = mat.ival(0)
                bigMat(3,start + 6) = mat.ival(1)
                bigMat(4,start + 6) = mat.ival(2)
                bigMat(5,start + 6) = mat.ival(3)
                bigMat(9,start + 6) = mat.ival(4)
                rowEmpty(start + 1,2,bigMat)
                rowEmpty(start + 6,2,bigMat)
                start = start + 7
            Case Else
                Select Case mat.subtype
                Case 501 To 504, 507
                    bigMat(1,start) = "MATHE"
                    bigMat(2,start) = mtID
                    If mat.mval(123) <> 0 Then
                        bigMat(5,start) = mat.mval(123)
                    End If
                    bigMat(6,start) = mat.mval(49)
                    bigMat(7,start) = mat.mval(36)
                    start = start
                    Select Case mat.subtype
                    Case 501
                        bigMat(3,start) = "MOONEY"
                        bigMat(2,start + 1) = mat.mval(64)
                        bigMat(3,start + 1) = mat.mval(59)
                        bigMat(2,start + 2) = mat.mval(69)
                        bigMat(3,start + 2) = mat.mval(65)
                        bigMat(4,start + 2) = mat.mval(60)
                        bigMat(2,start + 3) = mat.mval(73)
                        bigMat(3,start + 3) = mat.mval(70)
                        bigMat(4,start + 3) = mat.mval(66)
                        bigMat(5,start + 3) = mat.mval(61)
                        start = start + 4
                    Case 502
                        bigMat(3,start) = "FOAM"
                        Ogden:
                        bigMat(2,start + 1) = mat.mval(130)
                        bigMat(3,start + 1) = mat.mval(140)
                        bigMat(4,start + 1) = mat.mval(170)
                        bigMat(2,start + 2) = mat.mval(131)
                        bigMat(3,start + 2) = mat.mval(141)
                        bigMat(4,start + 2) = mat.mval(171)
                        bigMat(5,start + 2) = mat.mval(132)
                        bigMat(6,start + 2) = mat.mval(142)
                        bigMat(7,start + 2) = mat.mval(172)
                        bigMat(2,start + 3) = mat.mval(133)
                        bigMat(3,start + 3) = mat.mval(143)
                        bigMat(4,start + 3) = mat.mval(173)
                        bigMat(5,start + 3) = mat.mval(134)
                        bigMat(6,start + 3) = mat.mval(144)
                        bigMat(7,start + 3) = mat.mval(174)
                        bigMat(2,start + 4) = mat.mval(135)
                        bigMat(3,start + 4) = mat.mval(145)
                        bigMat(4,start + 4) = mat.mval(175)
                        bigMat(5,start + 4) = mat.mval(136)
                        bigMat(6,start + 4) = mat.mval(146)
                        bigMat(7,start + 4) = mat.mval(176)
                        bigMat(2,start + 5) = mat.mval(137)
                        bigMat(3,start + 5) = mat.mval(147)
                        bigMat(4,start + 5) = mat.mval(177)
                        bigMat(5,start + 5) = mat.mval(138)
                        bigMat(6,start + 5) = mat.mval(148)
                        bigMat(7,start + 5) = mat.mval(178)
                        start = start + 6
                    Case 503
                        bigMat(3,start) = "OGDEN"
                        GoTo Ogden
                    Case 504
                        bigMat(3,start) = "ABOYCE"
                        bigMat(2,start + 1) = mat.mval(98)
                        bigMat(3,start + 1) = mat.mval(99)
                        start = start + 2
                    Case 507
                        bigMat(3,start) = "SUSSBAT"
                        bigMat(2,start + 1) = mat.fval(0)
                        bigMat(3,start + 1) = -mat.bVal(4)
                        bigMat(4,start + 1) = mat.mval(102)
                        start = start + 2
                    End Select
                    bigMat(3,start) = mat.mval(103)
                    bigMat(4,start) = mat.mval(104)
                    bigMat(5,start) = mat.mval(105)
                    If Not rowEmpty(start,3,bigMat) Then
                        bigMat(1,start) = "MATHEM"
                        bigMat(2,start) = mtID
                        start = start + 1
                    End If

                    bigMat(3,start) = mat.ival(1)
                    bigMat(4,start) = mat.mval(106)
                    bigMat(5,start) = mat.mval(107)
                    bigMat(6,start) = mat.mval(108)
                    If Not rowEmpty(start,3,bigMat) Then
                        bigMat(1,start) = "MATHEV"
                        bigMat(2,start) = mtID
                        start = start + 1
                        For i = 0 To 9
                            bigMat(2,start) = mat.mval(110 + i)
                            bigMat(3,start) = mat.mval(120 + i)
                            bigMat(4,start) = mat.mval(160 + i)
                            bigMat(5,start) = mat.ival(10 + i)
                            If Not rowEmpty(start,2,bigMat) Then
                                start = start + 1
                            Else
                                Exit For
                            End If
                        Next
                    End If
                Case 510
                    bigMat(1,start) = "MATVE"
                    bigMat(2,start) = mtID
                    bigMat(3,start) = mat.fval(1)
                    bigMat(4,start) = mat.fval(2)
                    bigMat(5,start) = mat.mval(49)
                    bigMat(6,start) = mat.mval(36)
                    If mat.ival(1) <> 0 Then
                        bigMat(2,start + 1) = mat.ival(1)
                    End If
                    bigMat(3,start + 1) = mat.mval(103)
                    bigMat(4,start + 1) = mat.mval(104)
                    bigMat(5,start + 1) = mat.mval(51)
                    start = start + 2
                Case 505
                    bigMat(1,start) = "MATG"
                    bigMat(2,start) = mtID
                    bigMat(3,start) = mat.ival(0)
                    bigMat(4,stat) = 0
                    bigMat(5,start) = mat.fval(0)
                    bigMat(6,start) = mat.fval(1)
                    bigMat(7,start) = mat.fval(2)
                    bigMat(8,start) = mat.fval(3)
                    bigMat(9,start) = mat.fval(4)
                    For i = 2 To 7
                        If mat.fval(i + 3) <> 0 Then
                            bigMat(i,start) = mat.fval(i + 3)
                        End If
                    Next
                    bigMat(8,start + 1) = mat.mval(85)
                    bigMat(9,start + 1) = mat.mval(0)
                    bigMat(2,start + 2) = mat.mval(3)
                    rowEmpty(start,5,bigMat)
                    start = start + 3
                Case 506
                    bigMat(1,start) = "MAT10"
                    bigMat(2,start) = mtID
                    bigMat(3,start) = mat.mval(123)
                    bigMat(4,start) = mat.mval(49)
                    bigMat(5,start) = mat.mval(124)
                    bigMat(7,start) = mat.mval(170)
                    rowEmpty(start,3,bigMat)
                    bigMat(6,start) = mat.mval(50)
                    start = start + 1
                    bigMat(3,start) = mat.fval(39)
                    bigMat(4,start) = mat.fval(49)
                    bigMat(6,start) = mat.fval(40)
                    bigMat(7,start) = mat.fval(41)
                    If Not rowEmpty(start,3,bigMat) Then
                        start = start + 1
                    End If
                Case 508
                    bigMat(1,start) = "MATSMA"
                    bigMat(2,start) = mtID
                    bigMat(3,start) = mat.mval(0)
                    bigMat(4,start) = mat.mval(1)
                    bigMat(5,start) = mat.mval(6)
                    bigMat(6,start) = mat.mval(7)
                    bigMat(7,start) = mat.mval(60)
                    bigMat(8,start) = mat.mval(61)
                    bigMat(9,start) = mat.mval(49)
                    bigMat(2,start + 1) = mat.mval(36)
                    bigMat(3,start + 1) = mat.mval(37)
                    bigMat(4,start + 1) = mat.mval(51)
                    bigMat(5,start + 1) = mat.mval(52)
                    bigMat(6,start + 1) = mat.mval(53)
                    bigMat(7,start + 1) = mat.mval(54)
                    bigMat(8,start + 1) = mat.mval(55)
                    bigMat(2,start + 2) = mat.mval(62)
                    bigMat(3,start + 2) = mat.mval(63)
                    bigMat(4,start + 2) = mat.mval(64)
                    bigMat(5,start + 2) = mat.mval(65)
                    bigMat(6,start + 2) = mat.mval(66)
                    start = start + 3
                End Select
            End Select
        Wend
    End If

    If dlg.ndBox Then
        While nSet.Next
            bigMat(2,start) = nSet.Current
            nd.Get(bigMat(2,start))
            Select Case nd.type
            Case FND_NODE
                bigMat(1,start) = "GRID"
                   bigMat(7,start) = nd.outCSys
            Case FND_SPOINT
                bigMat(1,start) = "SPOINT"
                GoTo Done
            Case FND_EPOINT
                bigMat(1,start) = "EPOINT"
                GoTo Done
            Case FND_FLPOINT
                bigMat(1,start) = "GRID"
                bigMat(7,start)= -1
            End Select
            bigMat(3,start)= nd.defCSys
            bigMat(4,start) = nd.x
            bigMat(5,start) = nd.y
            bigMat(6,start) = nd.z
            For k = 0 To 5
                If nd.permBC(k) Then
                    bigMat(8,start) = bigMat(8,start) + Str$(k+1)
                End If
            Next
            If nd.superelementID = 0 Then
                bigMat(9,start) = Empty
            Else
                bigMat(9,start) = nd.superelementID
            End If
            Done:
            start = start + 1
        Wend
    End If

    If dlg.elBox Then
        While eSet.Next
             elID = eSet.Current
             el.Get(elID)
               typeID = el.type
               topID = el.topology
              Select Case typeID
              Case FET_L_ROD
                  bigMat(3,start) = el.propID
                  pRod:
                    bigMat(1,start) = "CROD"
                    bigMat(2,start) = elID
                    bigMat(4,start) = el.Node(0)
                    bigMat(5,start) = el.Node(1)
                    start = start + 1
                Case FET_L_BAR
                    bigMat(1,start) = "CBAR"
                    bigMat(2,start) = elID
                    bigMat(3,start) = el.propID
                    bigMat(4,start) = el.Node(0)
                    bigMat(5,start) = el.Node(1)
                    bigMat(6,start) = el.orient(0)
                    bigMat(7,start) = el.orient(1)
                    bigMat(8,start) = el.orient(2)
                    bigMat(2,start + 1) = el.formulation(0)
                    bigMat(3,start + 1) = el.formulation(1)
                bigMat(4,start + 1) = el.offset(0,0)
                bigMat(5,start + 1) = el.offset(0,1)
                bigMat(6,start + 1) = el.offset(0,2)
                bigMat(7,start + 1) = el.offset(1,0)
                bigMat(8,start + 1) = el.offset(1,1)
                bigMat(9,start + 1) = el.offset(1,2)
                rowEmpty(start + 1, 2,bigMat)
                start = start + 2
            Case FET_L_BEAM, FET_P_BEAM
                bigMat(1,start) = "CBEAM"
                    bigMat(2,start) = elID
                    bigMat(3,start) = el.propID
                    bigMat(4,start) = el.Node(0)
                    bigMat(5,start) = el.Node(1)
                    bigMat(6,start) = el.orient(0)
                    bigMat(7,start) = el.orient(1)
                    bigMat(8,start) = el.orient(2)
                    bigMat(9,start) = Empty
                    bigMat(2,start + 1) = el.formulation(0)
                    bigMat(3,start + 1) = el.formulation(1)
                bigMat(4,start + 1) = el.offset(0,0)
                bigMat(5,start + 1) = el.offset(0,1)
                bigMat(6,start + 1) = el.offset(0,2)
                bigMat(7,start + 1) = el.offset(1,0)
                bigMat(8,start + 1) = el.offset(1,1)
                bigMat(9,start + 1) = el.offset(1,2)
                bigMat(2,start + 2) = el.Warping(0)
                bigMat(3,start + 2) = el.Warping(1)
                rowEmpty(start + 1, 2,bigMat)
                rowEmpty(start + 2, 2,bigMat)
                start = start + 3
            Case FET_L_SPRING
                If el.formulation(0) Then
                    SpringGr:
                    bigMat(1,start) = "CBUSH"
                    bigMat(2,start) = elID
                        bigMat(3,start) = el.propID
                        bigMat(4,start) = el.Node(0)
                        bigMat(5,start) = el.Node(1)
                        bigMat(6,start) = el.orient(0)
                        bigMat(7,start) = el.orient(1)
                        bigMat(8,start) = el.orient(2)
                        bigMat(9,start) = el.MaterialCSys
                        bigMat(2,start + 1) = el.SpringLocation
                    bigMat(3,start + 1) = el.SpringCID
                    If el.SpringCID <> 0 Then
                        bigMat(4,start + 1) = el.offset(0,0)
                        bigMat(5,start + 1) = el.offset(0,1)
                        bigMat(6,start + 1) = el.offset(0,2)
                    End If
                start = start + 2
                Else
                    pr.Get(el.propID)
                    If pr.pval(0) = 0 And pr.pval(1) <> 0 Then
                        bigMat(1,start) = "CVISC"
                        bigMat(2,start) = elID
                        bigMat(3,start) = el.propID
                        bigMat(4,start) = el.Node(0)
                        bigMat(5,start) = el.Node(1)
                        start = start + 1
                    Else
                        bigMat(3,start) = pr.NextEmptyID
                        GoTo pRod
                    End If
                End If
            Case FET_L_DOF_SPRING, FET_L_DOF_SPRING_TO_GROUND
                pr.Get(el.propID)
                If el.formulation(0) Then
                    If pr.pval(0) = 0 And pr.pval(1) <> 0 Then
                        Damp:
                        bigMat(1,start) = "CDAMP1"
                    Else
                        bigMat(1,start) = "CELAS1"
                    End If
                    bigMat(2,start) = elID
                    bigMat(3,start) = pr.ID
                    bigMat(4,start) = el.Node(0)
                    bigMat(5,start) = pr.flagI(2)
                    bigMat(6,start) = el.Node(1)
                    bigMat(7,start) = pr.flagI(3)
                Else
                    If pr.pval(0) = 0 And pr.pval(1) <> 0 Then
                        If pr.pval(23) <> 0 Then
                            GoTo Damp
                        Else
                            bigMat(1,start) = "CDAMP2"
                            bigMat(2,start) = elID
                            bigMat(3,start) = pr.pval(1)
                            If bigMat(4,start) <> 0 Then
                                bigMat(5,start) = 1
                            End If
                            bigMat(6,start) = el.Node(1)
                            If bigMat(6,start) <> 0 Then
                                bigMat(7,start) = 1
                            End If
                        End If
                    Else
                        bigMat(1,start) = "CELAS2"
                        bigMat(2,start) = elID
                        bigMat(3,start) = pr.pval(0)
                        bigMat(4,start) = el.Node(0)
                        If bigMat(4,start) <> 0 Then
                            bigMat(5,start) = 1
                        End If
                        bigMat(6,start) = el.Node(1)
                        If bigMat(6,start) <> 0 Then
                            bigMat(7,start) = 1
                        End If
                        bigMat(8,start) = pr.pval(1)
                        bigMat(9,start) = pr.pval(2)
                    End If
                End If
                start = start + 1
            Case FET_L_CURVED_BEAM, FET_L_CURVED_TUBE
                bigMat(1,start) = "CBEND"
                    bigMat(2,start) = elID
                    bigMat(3,start) = el.propID
                    bigMat(4,start) = el.Node(0)
                    bigMat(5,start) = el.Node(1)
                    bigMat(6,start) = el.orient(0)
                    bigMat(7,start) = el.orient(1)
                    bigMat(8,start) = el.orient(2)
                   bigMat(9,start) = 3
                    start = start + 1
                Case FET_L_GAP
                bigMat(1,start) = "CGAP"
                    bigMat(2,start) = elID
                    bigMat(3,start) = el.propID
                    bigMat(4,start) = el.Node(0)
                    bigMat(5,start) = el.Node(1)
                    bigMat(6,start) = el.orient(0)
                    bigMat(7,start) = el.orient(1)
                    bigMat(8,start) = el.orient(2)
                    If el.MaterialCSys <> 0 Then
                    bigMat(9,start) = el.MaterialCSys
                    End If
                    start = start + 1
                Case FET_L_PLOT
                    bigMat(1,start) = "PLOTEL"
                    bigMat(2,start) = elID
                    bigMat(3,start) = el.Node(0)
                    bigMat(4,start) = el.Node(1)
                    start = start + 1
                Case FET_L_SHEAR
                    bigMat(1,start) = "CSHEAR"
                    bigMat(2,start) = elID
                    bigMat(3,start) = el.propID
                    bigMat(4,start) = el.Node(0)
                    bigMat(5,start) = el.Node(1)
                    bigMat(6,start) = el.Node(2)
                    bigMat(7,start) = el.Node(3)
                    start = start + 1
            Case FET_L_PLANE_STRAIN
                If el.formulation(0) = 1 Then
                    Select Case topID
                    Case FTO_TRIA3
                        bigMat(1,start) = "CPLSTN3"
                            bigMat(2,start) = elID
                            bigMat(3,start) = el.propID
                            bigMat(4,start) = el.Node(0)
                            bigMat(5,start) = el.Node(1)
                            bigMat(6,start) = el.Node(2)
                            If el.MaterialCSys <> 0  Then
                                bigMat(7,start) = el.MaterialCSys
                            End If
                            start = start + 1
                        Case FTO_TRIA6
                        bigMat(1,start) = "CPLSTN6"
                            bigMat(2,start) = elID
                            bigMat(3,start) = el.propID
                            bigMat(4,start) = el.Node(0)
                            bigMat(5,start) = el.Node(1)
                            bigMat(6,start) = el.Node(2)
                            bigMat(7,start) = el.Node(3)
                            bigMat(8,start) = el.Node(4)
                            bigMat(9,start) = el.Node(5)
                            If el.MaterialCSys <> 0  Then
                                bigMat(2,start + 1) = el.MaterialCSys
                            End If
                            start = start + 2
                        Case FTO_QUAD4
                        bigMat(1,start) = "CPLSTN4"
                            bigMat(2,start) = elID
                            bigMat(3,start) = el.propID
                            bigMat(4,start) = el.Node(0)
                            bigMat(5,start) = el.Node(1)
                            bigMat(6,start) = el.Node(2)
                            bigMat(7,start) = el.Node(3)
                            If el.MaterialCSys <> 0  Then
                                bigMat(8,start) = el.MaterialCSys
                            End If
                            start = start + 1
                        Case FTO_QUAD8
                        bigMat(1,start) = "CPLSTN3"
                            bigMat(2,start) = elID
                            bigMat(3,start) = el.propID
                            bigMat(4,start) = el.Node(0)
                            bigMat(5,start) = el.Node(1)
                            bigMat(6,start) = el.Node(2)
                            bigMat(7,start) = el.Node(3)
                            bigMat(8,start) = el.Node(4)
                            bigMat(9,start) = el.Node(5)
                            bigMat(2,start + 1) = el.Node(6)
                            bigMat(3,start + 1) = el.Node(7)
                            If el.MaterialCSys <> 0  Then
                                bigMat(4,start + 1) = el.MaterialCSys
                            End If
                            start = start + 2
                        End Select
                    ElseIf el.formulation(0) = 2 Then
                        Select Case topID
                        Case FTO_TRIA3
                        bigMat(1,start) = "CPLSTS3"
                            bigMat(2,start) = elID
                            bigMat(3,start) = el.propID
                            bigMat(4,start) = el.Node(0)
                            bigMat(5,start) = el.Node(1)
                            bigMat(6,start) = el.Node(2)
                            If el.MaterialCSys <> 0  Then
                                bigMat(8,start) = el.MaterialCSys
                            End If
                        start = start + 2
                    Case FTO_TRIA6
                        bigMat(1,start) = "CPLSTS6"
                            bigMat(2,start) = elID
                            bigMat(3,start) = el.propID
                            bigMat(4,start) = el.Node(0)
                            bigMat(5,start) = el.Node(1)
                            bigMat(6,start) = el.Node(2)
                            bigMat(7,start) = el.Node(3)
                            bigMat(8,start) = el.Node(4)
                            bigMat(9,start) = el.Node(5)
                            If el.MaterialCSys <> 0  Then
                                bigMat(4,start + 1) = el.MaterialCSys
                            End If
                            start = start + 2
                        Case FTO_QUAD4
                        bigMat(1,start) = "CPLSTS4"
                            bigMat(2,start) = elID
                            bigMat(3,start) = el.propID
                            bigMat(4,start) = el.Node(0)
                            bigMat(5,start) = el.Node(1)
                            bigMat(6,start) = el.Node(2)
                            bigMat(7,start) = el.Node(3)
                            If el.MaterialCSys <> 0  Then
                                bigMat(8,start) = el.MaterialCSys
                            End If
                            start = start + 2
                        Case FTO_QUAD8
                        bigMat(1,start) = "CPLSTN3"
                            bigMat(2,start) = elID
                            bigMat(3,start) = el.propID
                            bigMat(4,start) = el.Node(0)
                            bigMat(5,start) = el.Node(1)
                            bigMat(6,start) = el.Node(2)
                            bigMat(7,start) = el.Node(3)
                            bigMat(8,start) = el.Node(4)
                            bigMat(9,start) = el.Node(5)
                            bigMat(2,start + 1) = el.Node(6)
                            bigMat(3,start + 1) = el.Node(7)
                            If el.MaterialCSys <> 0  Then
                                bigMat(4,start + 1) = el.MaterialCSys
                            End If
                            start = start + 2
                        End Select
                    Else
                        GoTo Geom
                End If
            Case FET_L_TUBE
                bigMat(1,start) = "CTUBE"
                    bigMat(2,start) = elID
                    bigMat(3,start) = el.propID
                    bigMat(4,start) = el.Node(0)
                    bigMat(5,start) = el.Node(1)
                    start = start + 1
                Case FET_L_WELD
                    pr.Get(el.propID)
                    If pr.pval(9) = 0  Then
                        bigMat(1,start) = "CWELD"
                        bigMat(2,start) = elID
                        bigMat(3,start) = pr.ID
                    Select Case el.connectTYPE(0)
                    Case 0,1
                        bigMat(4,start) = el.orientID
                        bigMat(5,start) = "ELEMID"
                        bigMat(2,start + 1) = el.connectSEG(0)
                        bigMat(3,start + 1) = el.connectSEG(1)
                        start = start + 2
                    Case 2
                        bigMat(5,start) = "ALIGN"
                        bigMat(6,start) = el.connectSEG(0)
                        bigMat(7,start) = el.connectSEG(1)
                        start = start + 1
                    Case 3
                        bigMat(4,start) = el.orientID
                        bigMat(5,start) = "ELPAT"
                        bigMat(2,start + 1) = el.connectSEG(0)
                        bigMat(3,start + 1) = el.connectSEG(1)
                        start = start + 1
                    Case 4
                        bigMat(4,start) = el.orientID
                        bigMat(5,start) = "ELPAT"
                        bigMat(2,start + 1) = el.connectSEG(0)
                        bigMat(3,start + 1) = el.connectSEG(1)
                        start = start + 1
                    Case 5,6
                        bigMat(4,start) = el.orientID
                        bigMat(5,start) = "GRIDID"
                        bigMat(2,start + 1) = el.Node(4)
                        bigMat(3,start + 1) = el.Node(5)
                        bigMat(4,start + 1) = el.Node(6)
                        bigMat(5,start + 1) = el.Node(7)
                        bigMat(6,start + 1) = el.Node(8)
                        bigMat(7,start + 1) = el.Node(9)
                        bigMat(8,start + 1) = el.Node(10)
                        bigMat(9,start + 1) = el.Node(11)
                        bigMat(2,start + 2) = el.Node(12)
                        bigMat(3,start + 2) = el.Node(13)
                        bigMat(4,start + 2) = el.Node(14)
                        bigMat(5,start + 2) = el.Node(15)
                        bigMat(6,start + 2) = el.Node(16)
                        bigMat(7,start + 2) = el.Node(17)
                        bigMat(8,start + 2) = el.Node(18)
                        bigMat(9,start + 2) = el.Node(19)
                        start = start +2
                    End Select
                Else
                    bigMat(1,start) = "CFAST"
                    bigMat(2,start) = elID
                    bigMat(3,start) = el.propID
                    bigMat(5,start) = el.connectSEG(0)
                    bigMat(6,start) = el.connectSEG(1)
                    bigMat(7,start) = el.orientID
                    bigMat(8,start) = el.Node(0)
                    bigMat(9,start) = el.Node(1)
                    Select Case el.connectTYPE(0)
                    Case 7
                        bigMat(4,start) = "ELEM"
                    Case 8
                        bigMat(4,start) = "PROP"
                    End Select
                    bigMat(2,start + 1) = el.orient(0)
                    bigMat(3,start + 1) = el.orient(1)
                    bigMat(4,start + 1) = el.orient(2)
                    rowEmpty(start,1,bigMat)
                    rowEmpty(start + 1,2,bigMat)
                    start = start + 2
                    End If
                Case FET_L_SPRING_TO_GROUND
                    GoTo SpringGr
                Case FET_L_MASS
                    pr.Get(el.propID)
                    bigMat(1,start) = "CONM2"
                    bigMat(2,start) = elID
                    bigMat(3,start) = el.propID
                    bigMat(4,start) = el.MaterialCSys
                    bigMat(5,start) = pr.pval(7)
                    bigMat(6,start) = pr.pval(8)
                    bigMat(7,start) = pr.pval(9)
                    bigMat(8,start) = pr.pval(10)
                    bigMat(2, start + 1) = pr.pval(1)
                    bigMat(3, start + 1) = pr.pval(2)
                    bigMat(4, start + 1) = pr.pval(3)
                    bigMat(5, start + 1) = pr.pval(4)
                    bigMat(6, start + 1) = pr.pval(5)
                    bigMat(7, start + 1) = pr.pval(6)
                    rowEmpty(start + 1,2,bigMat)
                    start = start + 2
              Case FET_L_MASS_MATRIX
                pr.Get(el.propID)
                    bigMat(1,start) = "CONM1"
                    bigMat(2,start) = elID
                    bigMat(3,start) = el.propID
                    bigMat(4,start) = el.MaterialCSys
                    bigMat(5,start) = pr.pval(0)
                    bigMat(6,start) = pr.pval(1)
                    bigMat(7,start) = pr.pval(7)
                    bigMat(8,start) = pr.pval(2)
                    bigMat(9,start) = pr.pval(8)
                    bigMat(2,start + 1) = pr.pval(14)
                    bigMat(3,start + 1) = pr.pval(3)
                    bigMat(4,start + 1) = pr.pval(9)
                    bigMat(5,start + 1) = pr.pval(15)
                  bigMat(6,start + 1) = pr.pval(21)
                bigMat(7,start + 1) = pr.pval(4)
                bigMat(8,start + 1) = pr.pval(10)
                bigMat(9,start + 1) = pr.pval(16)
                bigMat(2,start + 2) = pr.pval(22)
                bigMat(3,start + 2) = pr.pval(28)
                bigMat(4,start + 2) = pr.pval(5)
                bigMat(5,start + 2) = pr.pval(11)
                bigMat(6,start + 2) = pr.pval(17)
                bigMat(7,start + 2) = pr.pval(23)
                bigMat(8,start + 2) = pr.pval(29)
                bigMat(9,start + 2) = pr.pval(35)
                start = start + 3
            Case FET_L_RIGID
                If el.formulation(1) = 0 Then
                    If el.RigidInterpolate Then
                        bigMat(1,start) = "RBE3"
                            bigMat(2,start) = elID
                            bigMat(4,start) = el.Node(0)
                            For i = 0 To 5
                                If el.release(0,i) Then
                                    bigMat(5,start) = bigMat(5,start) + Str$(i + 1)
                                End If
                            Next
                            RigidList(list,0,el,cnt)
                            curCol = 4
                            For i = curCol To (cnt + curCol - 1)
                            bigMat((i Mod 8) + 2, start + (i\8)) = list(i - curCol)
                            Next
                            start = start + 1 + (i\8)
                            rc = RigidList(list,1,el,count)
                            If rc = FE_OK Then
                                bigMat(2,start) = "UM"
                                curCol = 1
                                For i = curCol To (count + curCol - 1)
                                bigMat((i Mod 8) + 2, start + (i\8)) = list(i - curCol)
                                Next
                        start = start + 1 + (i\8)
                        End If
                        If el.RigidThermalExpansion <> 0 Then
                            bigMat(2,start) = "ALPHA"
                            bigMat(3,start) = el.RigidThermalExpansion
                            start = start + 1
                        End If
                    Else
                        bigMat(1,start) = "RBE2"
                        bigMat(2,start) = elID
                            bigMat(3,start) = el.Node(0)
                            rc = el.GetNodeList(0,cnt, list,fc,wt,dof)
                            For i = 0 To 5
                                If dof(i) = 0 Then
                                    bigMat(4,start) = bigMat(4,start) + Str$(i + 1)
                                End If
                            Next
                            curCol = 3
                        For i = curCol To (cnt + curCol - 1)
                            bigMat((i Mod 8) + 2, start + (i\8)) = list(i - curCol)
                            Next
                            start = start + 1 + (i\8)
                    End If
                Else
                    bigMat(1,start) = "RSPLINE"
                    bigMat(2,start) = elID
                    rc = RsplineList(ndList,el,cnt)
                    curCol = 1
                        For i = curCol To (cnt + curCol - 1)
                        bigMat((i Mod 8) + 2, start + (i\8)) = ndList(i - curCol)
                        Next
                        start = start + (i\8)
                End If
            Case FET_L_SLIDE_LINE
                master = slideCount(elID)
                slave = master + 1
                slidelist(el,list,cnt,0)
                bigMat(1,start) = "BLSEG"
                bigMat(2,start) = master
                curCol = 1
                For i = curCol To (cnt + curCol - 1)
                        bigMat((i Mod 8) + 2, start + (i\8)) = list(i - curCol)
                    Next
                    start = start + (i\8) + 1
                    slidelist(el,list,cnt,1)
                    bigMat(1,start) = "BLSEG"
                bigMat(2,start) = slave
                curCol = 1
                For i = curCol To (cnt + curCol - 1)
                        bigMat((i Mod 8) + 2, start + (i\8)) = list(i - curCol)
                    Next
                    start = start + (i\8) + 1
                    pr.Get(el.propID)
                    bigMat(1,start) = "BCONP"
                    bigMat(2,start) = elID
                    bigMat(3,start) = slave
                    bigMat(4,start) = master
                    If pr.pval(0) <> 0 Then
                        bigMat(6,start) = pr.pval(0)
                    bigMat(7,start) = el.propID
                End If
                    If pr.flag(0) Then
                        bigMat(8,start) = 2
                    Else
                    bigMat(8,start) = 1
                    End If
                    If el.MaterialCSys <> 0 Then
                        bigMat(9,start) = el.MaterialCSys
                    End If
                    start = start + 1
                Case FET_L_AXISYM_SHELL, FET_P_AXISYM_SHELL, FET_L_AXISYM, FET_P_AXISYM
            Select Case topID
                Case FTO_TRIA3
                        bigMat(1,start) = "CTRAX3"
                        bigMat(2,start) = elID
                        bigMat(3,start) = el.propID
                        bigMat(4,start) = el.Node(0)
                        bigMat(5,start) = el.Node(1)
                        bigMat(6,start) = el.Node(2)
                        bigMat(7,start) = el.orient(0)
                        start = start + 1
                    Case FTO_TRIA6
                        bigMat(1,start) = "CTRAX6"
                        bigMat(2,start) = elID
                        bigMat(3,start) = el.propID
                        bigMat(4,start) = el.Node(0)
                        bigMat(5,start) = el.Node(1)
                        bigMat(6,start) = el.Node(2)
                        bigMat(7,start) = el.Node(4)
                        bigMat(8,start) = el.Node(5)
                        bigMat(9,start) = el.Node(6)
                        bigMat(2,start + 1) = el.orient(0)
                        rowEmpty(start + 1,1,bigMat)
                        start = start +2
                    Case FTO_QUAD4
                    bigMat(1,start) = "CQUADX4"
                        bigMat(2,start) = elID
                        bigMat(3,start) = el.propID
                        bigMat(4,start) = el.Node(0)
                        bigMat(5,start) = el.Node(1)
                        bigMat(6,start) = el.Node(2)
                        bigMat(7,start) = el.Node(3)
                        bigMat(8,start) = el.orient(0)
                        start = start + 1
                    Case FTO_QUAD8
                        bigMat(1,start) = "CQUADX8"
                        bigMat(2,start) = elID
                        bigMat(3,start) = el.propID
                        bigMat(4,start) = el.Node(0)
                        bigMat(5,start) = el.Node(1)
                        bigMat(6,start) = el.Node(2)
                        bigMat(7,start) = el.Node(3)
                        bigMat(8,start) = el.Node(4)
                        bigMat(9,start) = el.Node(5)
                        bigMat(2,start + 1) = el.Node(6)
                        bigMat(3,start + 1) = el.Node(7)
                        bigMat(4,start + 1) = el.orient(0)
                        rowEmpty(start + 1,1,bigMat)
                        start = start + 2
                    End Select
                    rowEmpty(start - 1,1,bigMat)

            Case Else
                    Geom:
                    Select Case topID
                    Case FTO_TRIA3
                        bigMat(1,start) = "CTRIA3"
                        bigMat(2,start) = elID
                        bigMat(3,start) = el.propID
                        bigMat(4,start) = el.Node(0)
                        bigMat(5,start) = el.Node(1)
                        bigMat(6,start) = el.Node(2)
                        bigMat(7,start) = el.orient(0)
                        bigMat(8,start) = el.PlateOffset
                        rowEmpty(start,1,bigMat)
                        start = start + 1
                        el.thickness(tf,t1,t2,t3,t4)
                        If t1 <> t2 Or t1 <> t3 Or t2 <> t3 Then
                            bigMat(4,start) = t1
                            bigMat(5,start) = t2
                            bigMat(6,start) = t3
                        If Not rowEmpty(start,4,bigMat) Then
                            start = start + 1
                        End If
                    End If
                    Case FTO_TRIA6
                        bigMat(1,start) = "CTRIA6"
                        bigMat(2,start) = elID
                        bigMat(3,start) = el.propID
                        bigMat(4,start) = el.Node(0)
                        bigMat(5,start) = el.Node(1)
                        bigMat(6,start) = el.Node(2)
                        bigMat(7,start) = el.Node(4)
                        bigMat(8,start) = el.Node(5)
                        bigMat(9,start) = el.Node(6)
                        bigMat(2,start + 1) = el.orient(0)
                        bigMat(3,start + 1) = el.PlateOffset
                        el.thickness(tf,t1,t2,t3,t4)
                        If t1 <> t2 Or t1 <> t3 Or t2 <> t3 Then
                            bigMat(4,start + 1) = t1
                            bigMat(5,start + 1) = t2
                            bigMat(6,start + 1) = t3
                        End If
                        rowEmpty(start + 1,1,bigMat)
                        start = start +2
                    Case FTO_QUAD4
                    bigMat(1,start) = "CQUAD4"
                        bigMat(2,start) = elID
                        bigMat(3,start) = el.propID
                        bigMat(4,start) = el.Node(0)
                        bigMat(5,start) = el.Node(1)
                        bigMat(6,start) = el.Node(2)
                        bigMat(7,start) = el.Node(3)
                        bigMat(8,start) = el.orient(0)
                        bigMat(9,start) = el.PlateOffset
                        rowEmpty(start,1,bigMat)
                        start = start + 1
                    el.thickness(tf,t1,t2,t3,t4)
                    If t1 <> t2 Or t1 <> t3 Or t1 <> t3 Or t2 <> t3 Or t2 <> t4 Or t3 <> t4 Then
                            bigMat(4,start) = t1
                            bigMat(5,start) = t2
                            bigMat(6,start) = t3
                            bigMat(7,start) = t4
                        If Not rowEmpty(start,4,bigMat) Then
                            start = start + 1
                        End If
                    End If
                    Case FTO_QUAD8
                        bigMat(1,start) = "CQUAD8"
                        bigMat(2,start) = elID
                        bigMat(3,start) = el.propID
                        bigMat(4,start) = el.Node(0)
                        bigMat(5,start) = el.Node(1)
                        bigMat(6,start) = el.Node(2)
                        bigMat(7,start) = el.Node(3)
                        bigMat(8,start) = el.Node(4)
                        bigMat(9,start) = el.Node(5)
                        bigMat(2,start + 1) = el.Node(6)
                        bigMat(3,start + 1) = el.Node(7)
                        el.thickness(tf,t1,t2,t3,t4)
                        If t1 <> t2 Or t1 <> t3 Or t1 <> t3 Or t2 <> t3 Or t2 <> t4 Or t3 <> t4 Then
                            bigMat(4,start) = t1
                            bigMat(5,start) = t2
                            bigMat(6,start) = t3
                            bigMat(7,start) = t4
                        End If
                        bigMat(8,start + 1) = el.orient(0)
                        bigMat(9,start + 1) = el.PlateOffset
                    rowEmpty(start + 1,1,bigMat)
                        start = start + 2
                    Case FTO_TETRA4, FTO_TETRA10
                    bigMat(1,start) = "CTETRA"
                        bigMat(2,start) = elID
                        bigMat(3,start) = el.propID
                        bigMat(4,start) = el.Node(0)
                        bigMat(5,start) = el.Node(1)
                        bigMat(6,start) = el.Node(2)
                      bigMat(7,start) = el.Node(4)
                        bigMat(8,start) = el.Node(8)
                        bigMat(9,start) = el.Node(9)
                        bigMat(2,start + 1) = el.Node(10)
                        bigMat(3,start + 1) = el.Node(12)
                        bigMat(4,start + 1) = el.Node(13)
                        bigMat(5,start + 1) = el.Node(14)
                        start = start + 1
                        If topID = FTO_TETRA10 Then start = start + 1
                    Case FTO_PYRAMID5, FTO_PYRAMID13
                    bigMat(1,start) = "CPYRAM"
                        bigMat(2,start) = elID
                        bigMat(3,start) = el.propID
                        bigMat(4,start) = el.Node(0)
                        bigMat(5,start) = el.Node(1)
                        bigMat(6,start) = el.Node(2)
                        bigMat(7,start) = el.Node(3)
                        bigMat(8,start) = el.Node(4)
                        bigMat(9,start) = el.Node(8)
                        bigMat(2,start + 1) = el.Node(9)
                        bigMat(3,start + 1) = el.Node(10)
                        bigMat(4,start + 1) = el.Node(11)
                        bigMat(5,start + 1) = el.Node(12)
                        bigMat(6,start + 1) = el.Node(13)
                        bigMat(7,start + 1) = el.Node(14)
                        bigMat(8,start + 1) = el.Node(15)
                        start = start + 1
                        If topID = FTO_PYRAMID13 Then start = start + 1
                    Case FTO_WEDGE15, FTO_WEDGE6
                    bigMat(1,start) = "CPENTA"
                        bigMat(2,start) = elID
                        bigMat(3,start) = el.propID
                        bigMat(4,start) = el.Node(0)
                        bigMat(5,start) = el.Node(1)
                        bigMat(6,start) = el.Node(2)
                        bigMat(7,start) = el.Node(4)
                        bigMat(8,start) = el.Node(5)
                        bigMat(9,start) = el.Node(6)
                        bigMat(2,start + 1) = el.Node(8)
                        bigMat(3,start + 1) = el.Node(9)
                        bigMat(4,start + 1) = el.Node(10)
                        bigMat(5,start + 1) = el.Node(12)
                        bigMat(6,start + 1) = el.Node(13)
                        bigMat(7,start + 1) = el.Node(14)
                        bigMat(8,start + 1) = el.Node(16)
                        bigMat(9,start + 1) = el.Node(17)
                        bigMat(2,start + 2) = el.Node(18)
                        start = start + 1
                        If topID = FTO_WEDGE15 Then start = start + 2
                    Case FTO_BRICK8, FTO_BRICK20
                    bigMat(1,start) = "CHEXA"
                        bigMat(2,start) = elID
                        bigMat(3,start) = el.propID
                        bigMat(4,start) = el.Node(0)
                        bigMat(5,start) = el.Node(1)
                        bigMat(6,start) = el.Node(2)
                        bigMat(7,start) = el.Node(3)
                        bigMat(8,start) = el.Node(4)
                        bigMat(9,start) = el.Node(5)
                        bigMat(2,start + 1) = el.Node(6)
                        bigMat(3,start + 1) = el.Node(7)
                        bigMat(4,start + 1) = el.Node(8)
                        bigMat(5,start + 1) = el.Node(9)
                        bigMat(6,start + 1) = el.Node(10)
                        bigMat(7,start + 1) = el.Node(11)
                        bigMat(8,start + 1) = el.Node(12)
                        bigMat(9,start + 1) = el.Node(13)
                        bigMat(2,start + 2) = el.Node(14)
                        bigMat(3,start + 2) = el.Node(15)
                        bigMat(4,start + 2) = el.Node(16)
                        bigMat(5,start + 2) = el.Node(17)
                        bigMat(6,start + 2) = el.Node(18)
                        bigMat(7,start + 2) = el.Node(19)
                        start = start + 2
                        If topID = FTO_BRICK20 Then start = start + 1
                    Case Else
                    End Select
            End Select
        Wend
    End If

    If start < 1 Then
        End
    End If

    ReDim rowID(start)
    For i = 0 To start - 1
        rowID(i) = i + 1
    Next

    k = 0
    ReDim eRows(start)
    For i =(start - 1) To 0 Step -1
        If IsEmpty(bigMat(2,i)) And IsEmpty(bigMat(3,i)) And IsEmpty(bigMat(4,i)) And IsEmpty(bigMat(5,i)) And IsEmpty(bigMat(6,i)) And IsEmpty(bigMat(7,i)) And IsEmpty(bigMat(8,i)) And IsEmpty(bigMat(9,i)) Then
            eRows(k) = i
            k = k + 1
        End If
    Next

    MatrixCol(col1,1,bigMat,start)
    MatrixCol(col2,2,bigMat,start)
    MatrixCol(col3,3,bigMat,start)
    MatrixCol(col4,4,bigMat,start)
    MatrixCol(col5,5,bigMat,start)
    MatrixCol(col6,6,bigMat,start)
    MatrixCol(col7,7,bigMat,start)
    MatrixCol(col8,8,bigMat,start)
    MatrixCol(col9,9,bigMat,start)

    App.feAppManagePanes("Data Table", 1)
    table.Visible = True
    table.Lock(False)
    table.Clear()

    DataPrint("1",col1,table,rowID,start)
    DataPrint("2", col2,table,rowID,start)
    DataPrint("3", col3,table,rowID,start)
    DataPrint("4", col4,table,rowID,start)
    DataPrint("5", col5,table,rowID,start)
    DataPrint("6", col6,table,rowID,start)
    DataPrint("7", col7,table,rowID,start)
    DataPrint("8", col8,table,rowID,start)
    DataPrint("9", col9,table,rowID,start)


    table.DeleteRows((k),eRows)

    App.feAppMessage(FCM_NORMAL,Str$(rc))
    table.SetColumnWidth (-1, 80)
    table.SetColumnVisible(2,False)
    table.Visible = True
End Sub
Sub DataPrint(pName As String, value As Variant, table As DataTable, prID As Variant, rows As Long)
    Dim colID As Long
    rc = table.AddColumn(False,False,FT_PROP,0,pName,FCT_STRING,rows,prID,value,colID)
End Sub
Sub MatrixCol(col As Variant,colID As Long, matrix As Variant, leng As Long)
    ReDim col(leng - 1)
    For i = 0 To leng - 1
        col(i) = matrix(colID, i)
    Next
End Sub
Function pValPrint(bigMat As Variant, start As Long, leng As Long, pr As Prop, noMat As Boolean, propCase As Variant) As Long
        k = 0
        If noMat Then
            row = ((leng + 1) \ 8) + 1
        Else
            row = ((leng +2) \ 8) + 1
            bigMat(3,start) = pr.matlID
        End If
        bigMat(2,start) = pr.ID
        k = 0
        For j = start To (row + start)
            For    i = 2 To 9
                If (i = 2 Or i = 3) And j = start Then
                    If i = 3 And noMat Then
                        GoTo Skip
                    End If
                Else
                    Skip:
                    If IsEmpty(propCase(k)) Then
                        bigMat(i,j) = Empty
                    Else
                        bigMat(i,j) = pr.pval(propCase(k))
                    End If
                    k = k + 1
                    If k > leng Then
                        GoTo OverDim
                    End If
                End If
            Next
        Next
        OverDim:
        pValPrint = row
End Function
Function isPlane(prID As Long,pShell As Boolean, pPlane As Boolean) As Boolean
    Dim App As femap.model
    Set App = feFemap()

    Dim el As femap.Elem
    Set el = App.feElem

    Dim eSet As femap.Set
    Set eSet = App.feSet

    pPlane = False
    pShell = False
    eSet.AddRule(prID,FGD_ELEM_BYPROP)

    While eSet.Next
        el.Get(eSet.CurrentID)
        If el.formulation(0) = 0 Then
            pShell  = True
        Else
            pPlane = True
        End If
    Wend
End Function
Function slidePropCount(prID As Long) As Long
    Dim App As femap.model
    Set App = feFemap()

    Dim slSet As femap.Set
    Set slSet = App.feSet

    Dim el As femap.Elem
    Set el = App.feElem

    Dim k As Long

    slSet.AddRule(ft_slide_line,FGD_ELEM_BYTYPE)
    k = 0

    While slSet.Next
        el.Get(slSet.CurrentID)
            If el.propID = prID Then
                Exit While
            End If
            k = k + 1
    Wend
    slidePropCount = (k * 2) + 1
End Function
Function RigidList(list As Variant, listID As Long, el As Elem, count As Long) As Long
    list = 2009
    ReDim list(6000)
    Dim cnt As Long
    Dim nd As Variant
    Dim wt As Variant
    Dim fc As Variant
    Dim dof As Variant
    Dim k As Long
    Dim vWt As Double
    Dim vDOF As String
    Dim sDOF As String


    rc = el.GetNodeList(listID,cnt,nd,fc,wt,dof)
    If listID = 0 Then
        vWt = wt(0)
        list(0) = vWt
        For i = 0 To 5
            If dof(i) Then
                vDOF = vDOF + Str$(i + 1)
            End If
        Next
        list(1) = vDOF

        k = 2
        For i = 0 To (cnt - 1)
            For j = 0 To 5
                If dof(j + (i * 6)) Then
                    sDOF = sDOF + Str$(j + 1)
                End If
            Next
            If vWt <> wt(i) Or sDOF <> vDOF Then
                vWt = wt(i)
                list(k) = vWt
                k = k + 1
                vDOF = sDOF
                list(k) = vDOF
                k = k + 1
            End If
            list(k) = nd(i)
            k = k +1
            sDOF = Empty
        Next
        ElseIf listID = 1 Then
            If rc = FE_OK Then
                k = 0
                For i = 0 To (cnt - 1)
                    For j = 0 To 5
                        If dof(j + (i * 6)) Then
                            sDOF = sDOF + Str$(j + 1)
                        End If
                    Next
                    list(k) = sDOF
                    k = k + 1
                    list(k) = nd(i)
                    k = k + 1
                    sDOF = Empty
                Next
            End If
    End If
    count = k
    RigidList = rc
End Function
Function RsplineList(list As Variant,el As Elem, count As Long) As Long
    list = 2009
    ReDim list(6000)
    Dim cnt As Long
    Dim nd As Variant
    Dim wt As Variant
    Dim fc As Variant
    Dim dof As Variant
    Dim k As Long
    Dim sDOF As String
    rc = el.GetNodeList(listID,cnt,nd,fc,wt,dof)
    list(0) = wt(0)
    list(1) = nd(0)
    k = 2
    For i = 1 To (cnt - 2)
        list(k) = nd(i)
        k = k +1
        For j = 0 To 5
            If dof((i * 6) + j) Then
                sDOF = sDOF + Str$(j + 1)
            End If
        Next
        list(k) = sDOF
        k = k +1
        sDOF = Empty
    Next
    list(k) = nd(cnt - 1)
    RsplineList = rc
    count = k + 1
End Function
Function slideCount(slID As Long) As Long
    Dim App As femap.model
    Set App = feFemap()

    Dim slSet As femap.Set
    Set slSet = App.feSet

    Dim el As femap.Elem
    Set el = App.feElem

    Dim k As Long

    slSet.AddAll(FT_ELEM)
    k = 0

    While slSet.Next
        el.Get(slSet.CurrentID)
        If el.type = FET_L_SLIDE_LINE Then
            If el.ID = slID Then
                Exit While
            End If
            k = k + 1
        End If
    Wend
    slideCount = (k * 2) + 1
End Function
Sub slidelist(el As Elem, list As Variant, count As Long, listID As Long)
    list = 2009
    ReDim list(6000)
    Dim cnt As Long
    Dim nd As Variant
    Dim wt As Variant
    Dim fc As Variant
    Dim dof As Variant
    Dim k As Long

    el.GetNodeList(listID,count,nd,fc,wt,dof)

    For i = 0 To count - 1
        list(i) = nd(i)
    Next
End Sub
Rem See DialogFunc help topic for more information.
Private Function dlgfunc(DlgItem$, Action%, SuppValue?) As Boolean
    Select Case Action%
    Case 1 ' Dialog box initialization
    Case 2 ' Value changing or button pressed
        If DlgValue("propElem") = 1 Then
            DlgValue("prBox",0)
            DlgValue("elBox",1)
            DlgEnable("elBox",False)
            DlgEnable("prBox", False)
        ElseIf DlgValue("propElem") = 0 Then
            DlgEnable("prBox", True)
        End If
        If DlgValue("matElem") = 1 Then
            DlgValue("mtBox",0)
            DlgValue("elBox",1)
            DlgEnable("elBox",False)
            DlgEnable("mtBox", False)
        ElseIf DlgValue("matElem") = 0 Then
            DlgEnable("mtBox", True)
        End If
        If DlgValue("matElem") = 0 And DlgValue("propElem") = 0 Then
            DlgEnable("elBox",True)
        End If
        Rem .dlgfunc = True ' Prevent button press from closing the dialog box
    Case 3 ' TextBox or ComboBox text changed
    Case 4 ' Focus changed
    Case 5 ' Idle
        Rem Wait .1 : .dlgfunc = True ' Continue getting idle actions
    Case 6 ' Function key
    End Select
End Function
Function rowEmpty(rowID As Long, first As Long, bigMat As Variant) As Boolean
    rowEmpty = True
    For i = first To 9
        If bigMat(i,rowID) <> 0 And bigMat(i,rowID) <> Empty And bigMat(i,rowID) <> "" Then
            rowEmpty = False
        Else
            bigMat(i,rowID) = Empty
        End If
    Next
End Function
Function isSolLam(mtID As Long) As Boolean
    Dim App As femap.model
    Set App = feFemap

    Dim el As femap.Elem
    Set el = App.feElem

    Dim eSet As femap.Set
    Set eSet = App.feSet

    eSet.AddRule(mtID,FGD_ELEM_BYMATL)

    isSolLam = False

    While eSet.Next
        el.Get(eSet.CurrentID)
        If el.type = FET_P_LAMINATE_SOLID Or el.type = FET_L_LAMINATE_SOLID Then
            isSolLam = True
        End If
    Wend

End Function
Function isSolid(mtID As Long) As Boolean

    Dim App As femap.model
    Set App = feFemap

    Dim el As femap.Elem
    Set el = App.feElem

    Dim eSet As femap.Set
    Set eSet = App.feSet

    eSet.AddRule(mtID,FGD_ELEM_BYMATL)

    isSolid = False

    While eSet.Next
        el.Get(eSet.CurrentID)
        Select Case el.type
        Case FET_L_SOLID, FET_P_SOLID, FET_L_LAMINATE_SOLID, FET_P_LAMINATE_SOLID
            isSolid= True
            Exit While
        End Select
    Wend

End Function
