'Macro 2D Thickness Plot
'From an original idea From E. Crespin 02/2009
'A. Pavageau 04/2012

Const wPC = 3   '% of width (from left)
Const hPC = 10   '% of height (from top)
Const vSpac = 4   'Vertical spacing

Sub Main
    Dim App As femap.model
    Set App = feFemap()
    App.feAppMessage(FCM_COMMAND,"Create 2D Thickness Plot")

    Dim d1 As Double
    Dim l1 As Long, l2 As Long

    Dim pr2DSet As femap.Set, pr As femap.Prop, sSet As femap.SortSet
    Dim txt As femap.text
    Dim v As femap.View

    Set pr2DSet = App.feSet
    Set pr = App.feProp
    Set sSet = App.feSort
    Set txt = App.feText
    Set v = App.feView


    pr2DSet.AddRule(FET_L_PLATE,FGD_PROP_BYTYPE)
    pr2DSet.AddRule(FET_P_PLATE,FGD_PROP_BYTYPE)
    pr2DSet.AddRule(FET_L_LAMINATE_PLATE,FGD_PROP_BYTYPE)
    pr2DSet.AddRule(FET_P_LAMINATE_PLATE,FGD_PROP_BYTYPE)

    'set view => color by prop
    App.feAppGetActiveView(l1)
    v.Get(l1)
    v.ColorMode(FVI_ELEMENT) = 3
    v.Put(v.ID)

    'sort
    While pr.NextInSet(pr2DSet.ID)
        sSet.AddReal(pr.ID,pr.pval(0),pr.ID,0)
    Wend
    sSet.Sort(True)

    'time to plot!
    With txt
           .ModelPosition = False : .AllViews = True : .layer = 1
        .FontNumber = v.Label(FVI_LABEL)

        'text box
        .DrawBorder = False : .color = v.color(FVI_VIEW_LEGEND)
        .text = "2D Thickness:"
        .TextPosition(0) = wPC+3 : .TextPosition(1) = hPC
        .Put(.NextEmptyID)
    End With

    sSet.Reset : d1 = 0 : l1 = 1
    While sSet.Next
        If sSet.dReal1<>d1 Then  'create 2 textbox
            pr.Get(sSet.Current) : d1 = sSet.dReal1

            With txt
                'color box
                .text = "  "
                .color = pr.color : .DrawBorder = True
                .BorderColor = pr.color : .BackColor = pr.color
                .TextPosition(0) = wPC : .TextPosition(1) = hPC+vSpac*l1
                .Put(.NextEmptyID)

                'text box
                .DrawBorder = False : .color = v.color(FVI_VIEW_LEGEND)
                .text = " " & CStr(pr.pval(0))
                .TextPosition(0) = .TextPosition(0) + 3
                .Put(.NextEmptyID)
            End With

            'change prop color
            If l2<>0 Then App.feModifyColor(FT_PROP,pr2DSet.ID,l2,False)
            l2 = pr.color

            pr2DSet.Clear
            l1 = l1+1
        Else
            pr2DSet.Add(sSet.Current)
        End If
    Wend

    App.feAppMessage(FCM_ERROR,"Macro is done")
End Sub
