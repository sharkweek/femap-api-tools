' Title: Set Transparency for All Surfaces.BAS
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 12.0

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered trade­mark
' of Siemens Digital Industries Software Inc. or its subsidiaries in the United
' States and in other countries.

Dim App As femap.model
Dim transparency As Double

Sub Main
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Set Transparency for All Surfaces")

    Dim rc As Long

    Dim s_set As femap.Set
    Set s_set = App.feSet

    Dim s As femap.Surface
    Set s = App.feSurface

    Dim v As femap.View
    Set v = App.feView

    Dim vid As Long

    Begin Dialog UserDialog 340,56,"Transparency",.DialogFunc ' %GRID:10,7,1,1
        TextBox 10,28,210,14,.percent
        text 10,7,220,14,"Percent (0=Opaque, 100=Clear)",.percent_text
        OKButton 240,7,90,21
        CancelButton 240,28,90,21
    End Dialog
    Dim dlg As UserDialog

    ' prompt user to select smoothing options
    If Dialog(dlg) = 0 Then
    	Exit Sub  ' exit if user hits 'Cancel'
    End If

    ' set transparency
    s_set.AddAll(FT_SURFACE)
    App.feModifyTransparency(FT_SURFACE, s_set.ID, transparency)

    ' prompt to enable entity colors
    App.feAppGetActiveView(vid)
    v.Get(vid)
    If v.ColorMode(FVI_SURFACE) <> 0 Then
        rc = App.feAppMessageBox(2, "Current view's surface color mode is not set to Entity Colors. Change to Entity Colors?")
        If rc = FE_OK Then
            v.ColorMode(FVI_SURFACE) = 0
            v.Put(v.ID)
        End If
    End If

    App.feViewRegenerate(0)

End Sub


Private Function DialogFunc(DlgItem$, Action%, SuppValue?) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
        DlgText "percent", "0"

	Case 3 ' TextBox or ComboBox text changed
		Select Case DlgItem
            ' validate iterations
            Case "percent"
                If IsNumeric(DlgText(DlgItem)) Then
                    transparency = CDbl(DlgText(DlgItem))

                    ' throw error if negative or non-integer
                    If transparency <= 0 Or transparency >100 Then
                        App.feAppMessageBox(0, "ERROR: Iterations must be greater than 0 and less than or equal to 100.")
                        DlgFocus(0)
                    End If

                Else  ' throw error if non-numeric value
                    App.feAppMessageBox(0, "ERROR: Iterations must be greater than 0 and less than or equal to 100.")
                    DlgFocus(0)

                End If

                DialogFunc = True
        End Select
	End Select
End Function