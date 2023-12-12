' Title: Create FREQ3 From Critical Modes Function.BAS
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

    App.feAppMessage(FCM_COMMAND, "Create FREQ3 From Critical Modes Function")

    Dim rc As Long

    Dim freq As femap.Frequency
    Set freq = App.feFreq()

    Dim fn As femap.Function
    Set fn = App.feFunction()

    Dim window As Double
    Dim f1 As Double
    Dim f2 As Double
    Dim lg As Boolean
    Dim nef As Long

    Dim fn_count As Long
    Dim x As Variant
    Dim y As Variant

    Begin Dialog UserDialog 370,70,"FREQ3 Options" ' %GRID:10,7,1,1
        Text 10,7,130,14,"Frequency window:",.freq_window_text,1
        TextBox 160,7,90,14,.window
        Text 10,28,130,14,"Number of points:",.nef_text,1
        TextBox 160,28,90,14,.nef
        Text 10,49,130,14,"Cluster:",.cluster_text,1
        TextBox 160,49,90,14,.cluster
        CheckBox 272,52,90,14,"Log",.lg
        OKButton 270,5,90,21
        CancelButton 270,26,90,21
    End Dialog
    Dim dlg As UserDialog

    ' prompt user
    If Dialog(dlg) = 0 Then
        Exit Sub  ' exit if user cancels
    Else
        window = CDbl(dlg.Window)
        nef = CLng(dlg.nef)
        cluster = CDbl(dlg.cluster)
        lg = dlg.lg
    End If

    ' pick function
    rc = fn.SelectID("Select modal frequency function...")
    If rc = FE_CANCEL Then
        Exit Sub
    Else
        fn.GetFunctionList(fn_count, x, y)
    End If

    For i = 0 to (fn_count - 1)
        f1 = x(i) - window/2
        f2 = x(i) + window/2
        freq.AddFreq3(f1, f2, nef, lg, cluster, 0, False)
    Next i

    App.feViewRegenerate(0)

End Sub
