' Title: Move Annotations.BAS
' Author: Andy Perez
' Femap API Version: 12.0

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Move Annotations")
    App.feAppMessage(FCM_NORMAL, "Move annotations positioned in the model. Screen positioned annotations are skipped.")

    Dim rc As Long

    Dim txt As femap.Text
    Set txt = App.feText

    Dim txt_set As femap.Set
    Set txt_set = App.feSet

    Dim v_base As Variant
    Dim v_dir As Variant
    Dim v_length As Double

    Dim txt_drawPointer As Boolean
    Dim txt_drawBorder As Boolean
    Dim txt_modelPos As Boolean
    Dim txt_horzJust As Long
    Dim txt_vertJust As Long
    Dim txt_visView As Long
    Dim txt_layer As Long
    Dim txt_fontID As Long
    Dim txt_backColor As Long
    Dim txt_bordColor As Long
    Dim txt_textColor As Long
    Dim txt_textLoc As Variant
    Dim txt_pointerLoc As Variant
    Dim txt_text As String

    ' prompt user to select text objects
    rc = txt_set.Select(FT_TEXT, True, "Select text objects...")
    If rc = FE_CANCEL Then
        Exit Sub
    ElseIf rc = FE_FAIL Then
        App.feAppMessage(FCM_ERROR, "Selected text objects do not exist.")
    Else
        txt_set.Reset()
    End If

    ' prompt user to select vector to move by
    If App.feVectorPick("Select vector...", False, v_length, v_base, v_dir) = FE_CANCEL Then
        Exit Sub
    End If

        Do While txt.NextInSet(txt_set.ID)
            App.feTextGet(txt.ID, txt_drawPointer, txt_drawBorder, txt_modelPos, txt_horzJust, txt_vertJust, txt_visView, txt_layer, txt_fontID, txt_backColor, txt_bordColor, txt_textColor, txt_textLoc, txt_pointerLoc, txt_text)

            If txt_modelPos Then
                For i = 0 To 2
                    txt_textLoc(i) += v_dir(i) * v_length
                    txt_pointerLoc(i) += v_dir(i) * v_length
                Next i
                App.feTextPut(txt.ID, txt_drawPointer, txt_drawBorder, txt_modelPos, txt_horzJust, txt_vertJust, txt_visView, txt_layer, txt_fontID, txt_backColor, txt_bordColor, txt_textColor, txt_textLoc, txt_pointerLoc, txt_text)

            Else  ' skip over screen positioned objects
                App.feAppMessage(FCM_ERROR, "Text object " & CStr(txt.ID) & " is not model-positioned. Object skipped...")
            End If

        Loop

        App.feViewRegenerate(0)

End Sub
