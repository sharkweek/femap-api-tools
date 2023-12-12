' Title: PickPropertyByType.BAS
' Author: Andy Perez
' Femap API Version: 12.0

Function PickPropertyByType(ParamArray type_cases()) As Long
    Dim prop_set As femap.Set
    Set prop_set = feFemap().feSet

    Dim p As femap.Prop
    Set p = feFemap().feProp

    Dim list_array() As String
    Dim id_array() As Long

    For i = 0 To UBound(type_cases) - 1
        prop_set.AddRule(type_cases(i), FGD_PROP_BYTYPE)
    Next

    ReDim list_array(prop_set.Count())
    ReDim id_array(prop_set.Count())

    prop_set.Reset()
    For i = 0 To (prop_set.Count() - 1)
        prop_set.Next()
        p.Get(prop_set.CurrentID)
        id_array(i) = p.ID
        list_array(i) = CStr(p.ID) & ".." & p.title
    Next i

    Begin Dialog UserDialog 400,170,"Select property..." ' %GRID:5,5,1,1
        ComboBox 10,20,380,125,list_array(),.property_id
        text 10,5,90,15,"Entity ID",.entity_id_text
        OKButton 80,145,110,20
        CancelButton 210,145,110,20
    End Dialog
    Dim dlg As UserDialog

    ' prompt user
    If Dialog(dlg) = 0 Then
        PickPropertyByType = FE_CANCEL
    Else
        PickPropertyByType = id_array(dlg.property_id)
    End If

End Function
