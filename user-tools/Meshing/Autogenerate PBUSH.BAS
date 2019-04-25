' Title: Autogenerate PBUSH.BAS
' Author: Andy Perez
' Date: April 2019
' Femap API Version: 12.0

' global vars
Dim App As femap.model
Dim rc As Long

Dim p_pbush As femap.Prop
Dim shell_1 As femap.Prop
Dim shell_2 As femap.Prop
Dim mat_1 As femap.Matl
Dim mat_2 As femap.Matl
Dim propSet As femap.Set

Dim t_1 As Double
Dim e_1 As Double
Dim t_2 As Double
Dim e_2 As Double
Dim d As Double
Dim e_f As Double
Dim nu_f As Double
Dim shear As Boolean
Dim orient_csys As femap.CSys
Dim again As Boolean
Dim pi As Double
Dim shell_prop As Boolean

Sub Main
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Autogenerate PBUSH")
    App.feAppMessage(FCM_NORMAL, "Autogenerate a PBUSH property from selected stiffness method for lap joint configurations")
    App.feAppMessage(FCM_WARNING, "Stiffnesses are calculated for use in coincident CBUSHs ONLY. Properties taken from shell elements are")
    App.feAppMessage(FCM_WARNING, "taken from the thickness at the first node.")

    Set p_pbush = App.feProp
    Set shell_1 = App.feProp
    Set shell_2 = App.feProp
    Set mat_1 = App.feMatl
    Set mat_2 = App.feMatl
    Set orient_csys = App.feCSys

    Dim c As Double

    '================================================================
    ' Dialog Dropdown Box Arrays
    '================================================================
    Dim method(6) As String
    method(0) = "Boeing"
    method(1) = "Huth (Airbus)"
    method(2) = "Swift (Douglas)"
    method(3) = "Tate/Rosenfeld"
    method(4) = "Grumman"

    Dim bolt_type(3) As String
    bolt_type(0) = "Bolted metallic"
    bolt_type(1) = "Riveted metallic"
    bolt_type(2) = "Bolted graphite/epoxy"

    '================================================================
    ' Default Values
    '================================================================
    t_1 = 1
    e_1 = 1
    t_2 = 1
    e_2 = 1
    d = 1
    e_f = 1
    nu_f = 0.3
    shear = 0
    orient_csys.ID = 0
    pi = 4 * Atn(1)

    '================================================================
    ' Dialog
    '================================================================
    PROMPT_USER:
    Begin Dialog UserDialog 400,266,"Autogenerate PBUSH",.dialogfunc ' %GRID:10,7,1,1
        GroupBox 10,45,190,215,"Values",.values_groupbox
        PushButton 20,231,170,21,"From shell properties...",.from_props
        TextBox 125,63,60,14,.t_1
        TextBox 125,84,60,14,.e_1
        TextBox 125,105,60,14,.t_2
        TextBox 125,126,60,14,.e_2
        TextBox 125,147,60,14,.d
        TextBox 125,168,60,14,.e_f
        TextBox 125,189,60,14,.nu_f
        GroupBox 210,95,180,85,"Huth Options",.huth_options_groupbox
        DropListBox 220,112,160,21,bolt_type(),.bolt_type
        OptionGroup .shear
            OptionButton 230,135,120,20,"single shear",.single_shear
            OptionButton 230,155,120,20,"double shear",.double_shear
        text 20,63,95,14,"Thickness 1:",.t_1_text,1
        text 20,84,95,14,"E 1:",.e_1_text,1
        text 90,7,95,14,"Property Title",.prop_title
        text 20,105,95,14,"Thickness 2:",.t_2_text,1
        text 20,126,95,14,"E 2:",.e_2_text,1
        text 20,147,95,14,"Diameter:",.d_text,1
        text 20,168,95,14,"E (Fastener):",.e_f_text,1
        text 20,189,95,14,"nu (Fastener):",.nu_f_text,1
        OKButton 210,217,180,21
        CancelButton 210,238,180,21
        text 10,7,70,14,"PID",.pid_text
        TextBox 90,23,290,14,.title
        TextBox 10,23,70,14,.pid
        PushButton 20,210,170,21,"Coordinate System...",.csys_select
        GroupBox 210,45,180,45,"Stiffness Method",.method_groupbox
        DropListBox 220,60,160,21,method(),.method
        PushButton 210,196,180,21,"More...",.more
    End Dialog
    Dim dlg As UserDialog

    ' prompt user
    If Dialog(dlg) = 0 Then
        Exit Sub  ' exit if user hits 'Cancel'
    End If

    '================================================================
    ' Input Processing
    '================================================================
    ' convert strings to numbers
    t_1 = CDbl(dlg.t_1)
    e_1 = CDbl(dlg.e_1)
    t_2 = CDbl(dlg.t_2)
    e_2 = CDbl(dlg.e_2)
    d = CDbl(dlg.d)
    e_f = CDbl(dlg.e_f)
    nu_f = CDbl(dlg.nu_f)

    ' validate pid
    If p_pbush.Get(CLng(dlg.pid)) = FE_NOT_EXIST Then
        p_pbush.id = CLng(dlg.pid)
    Else
        p_pbush.id = p_pbush.NextEmptyID
    End If

    ' validate title
    If dlg.title = "" Then
        p_pbush.title = "Untitled"
    Else
        p_pbush.title = dlg.title
    End If

    '================================================================
    ' Calculate stiffness
    '================================================================
    Select Case dlg.method
        Case 0  ' Boeing
            c = (2^((t_1 / d)^0.85))/t_1 * (1/e_1 + 3/(8*e_f)) + (2^((t_2 / d)^0.85))/t_2 * (1/e_2 + 3/(8*e_f))

        Case 1  ' Huth (Airbus)
            Dim a As Double
            Dim b As Double
            Dim n As Integer

            Select Case dlg.bolt_type
                Case 0
                    a = 2/3
                    b = 3.0
                Case 1
                    a = 2/5
                    b = 2.2
                Case 2
                    a = 2/3
                    b = 4.2
            End Select

            n = dlg.shear + 1

            c = ((t_1 + t_2)/(2*d))^a * (b/n) * (1/(t_1 * e_1) + 1/(n*t_2*e_2) + 1/(2*t_1*e_f) + 1/(2*n*t_2*e_f))

        Case 2  ' Swift (Douglas)
            c = 5/(d*e_f) + 0.8 * (1/(t_1*e_1) + 1/(t_2 * e_2))

        Case 3  ' Tate/Rosenfeld
            c = 1/(e_f*t_1) + 1/(e_f*t_2) + 1/(e_1*t_1) + 1/(e_2*t_2) + 32/(9*e_f*pi*d^2)*(1+nu_f)*(t_1 + t_2) + 8/(5*e_f*pi*d^4)*(t_1^3 + 5*t_1^2*t_2 + 5*t_1*t_2^2 + t_2^3)

        Case 4  ' Grumman
            c = (t_1 + t_2)^2 / (e_f * d^3) + 3.7*(1/(e_1*t_1))

    End Select

    '================================================================
    ' Create PBUSH
    '================================================================
    With p_pbush
        .type = 6  ' spring/damper property
        .cbush = 1
        .flag(0) = orient_csys.ID  ' orient csys
        .refCS = orient_csys.ID
        .pval(36) = e_f * d^2 * pi / 4  ' k_x = (EA)_bolt
        .pval(37) = 1 / c  ' k_y = 1/c
        .pval(38) = 1 / c  ' k_z = 1/c
        .Put(.ID)
    End With

    If again Then
        GoTo PROMPT_USER
    End If
    App.feViewRegenerate(0)

End Sub


Private Function DialogFunc(DlgItem$, Action%, SuppValue?) As Boolean
    Select Case Action
        Case 1  ' dialog box initialization
            ' starting values
            DlgText "t_1", CStr(t_1)
            DlgText "e_1", CStr(e_1)
            DlgText "t_2", CStr(t_2)
            DlgText "e_2", CStr(e_2)
            DlgText "d", CStr(d)
            DlgText "e_f", CStr(e_f)
            DlgText "nu_f", CStr(nu_f)
            DlgText "pid", CStr(p_pbush.NextEmptyID)
            DlgText "title", "Untitled"
            DlgValue "shear", shear
            again = False

            DlgEnable "nu_f", False
            DlgEnable "nu_f_text", False
            DlgEnable "huth_options_groupbox", False
            DlgEnable "bolt_type", False
            DlgEnable "shear", False

        Case 2  ' value changing or button pressed
            Select Case DlgItem
                Case "from_props"
                    ' make sure shell properties exist
                    Set propSet = App.feSet
                    propSet.AddAll(FT_PROP)

                    ' remove non-plate properties
                    For i = 1 To 16
                        propSet.RemoveRule(i, FGD_PROP_BYTYPE)
                    Next i
                    For i = 19 To 42
                        propSet.RemoveRule(i, FGD_PROP_BYTYPE)
                    Next i

                    If propSet.Count() = 0 Then
                        App.feAppMessageBox(0, "Error: no plate properties exist in model.")

                    Else  ' prompt user to select shell properties
                        ' validate shell property 1
                        shell_1.SelectID("Select first property...")
                        shell_1.Get(shell_1.ID)
                        shell_prop = False

                        Do While not shell_prop
                            Select Case shell_1.type
                                Case 17, 18  ' plate elements only
                                    shell_prop = True

                                    ' validate isotropic material
                                    mat_1.Get(shell_1.matlID)
                                    If mat_1.type <> 0 Then
                                        App.feAppMessageBox(0, "Error: must select shell with isotropic material.")
                                        shell_prop = False
                                    End If

                                Case Else
                                    App.feAppMessageBox(0, "Error: Must select shell property.")
                                    shell_prop = False

                            End Select
                        Loop

                        ' validate shell property 2
                        shell_2.SelectID("Select second property...")
                        shell_2.Get(shell_2.ID)
                        shell_prop = False

                        Do While not shell_prop
                            Select Case shell_1.type
                                Case 17, 18  ' plate elements only
                                    shell_prop = True

                                    ' validate isotropic material
                                    mat_2.Get(shell_2.matlID)
                                    If mat_2.type <> 0 Then
                                        App.feAppMessageBox(0, "Error: must select shell with isotropic material.")
                                        shell_prop = False
                                    End If

                                Case Else
                                    App.feAppMessageBox(0, "Must select shell property.")
                                    shell_prop = False

                            End Select
                        Loop

                        ' fill in shell property values
                        DlgText "t_1", CStr(shell_1.pval(0))
                        DlgText "e_1", CStr(mat_1.Ex)
                        DlgText "t_2", CStr(shell_2.pval(0))
                        DlgText "e_2", CStr(mat_2.Ex)

                    End If

                    DialogFunc = True

                Case "csys_select"
                    orient_csys.SelectID("Select orientation coordinate system")
                    orient_csys.Get(orient_csys.ID)
                    DialogFunc = True

                Case "more"
                    again = True

                Case "method"
                    Select Case SuppValue
                        Case 0, 2 To 4  ' all but huth
                            DlgEnable "nu_f", False
                            DlgEnable "nu_f_text", False
                            DlgEnable "huth_options_groupbox", False
                            DlgEnable "bolt_type", False
                            DlgEnable "shear", False

                        Case 1  ' huth
                            DlgEnable "nu_f", False
                            DlgEnable "nu_f_text", False
                            DlgEnable "huth_options_groupbox", True
                            DlgEnable "bolt_type", True
                            DlgEnable "shear", True

                    End Select
                    DialogFunc = True

            End Select

        Case 3  ' TextBox or ComboBox text changed
            ' validate user input values
            Select Case DlgItem
                Case "t_1"
                    If IsNumeric(DlgText(DlgItem)) Then
                        t_1 = CDbl(DlgText(DlgItem))
                        If t_1 <= 0 Then
                            App.feAppMessageBox(0, "Thickness must be a numeric value greater than 0.")
                            DlgFocus(2)
                        End If
                    Else
                        App.feAppMessageBox(0, "Thickness must be a numeric value greater than 0.")
                        DlgFocus(2)
                    End If
                    DialogFunc = True
                Case "e_1"
                    If IsNumeric(DlgText(DlgItem)) Then
                        e_1 = CDbl(DlgText(DlgItem))
                        If e_1 <= 0 Then
                            App.feAppMessageBox(0, "Modulus must be a numeric value greater than 0.")
                            DlgFocus(3)
                        End If
                    Else
                        App.feAppMessageBox(0, "Modulus must be a numeric value greater than 0.")
                        DlgFocus(3)
                    End If
                    DialogFunc = True
                Case "t_2"
                    If IsNumeric(DlgText(DlgItem)) Then
                        t_2 = CDbl(DlgText(DlgItem))
                        If t_2 <= 0 Then
                            App.feAppMessageBox(0, "Thickness must be a numeric value greater than 0.")
                            DlgFocus(4)
                        End If
                    Else
                        App.feAppMessageBox(0, "Thickness must be a numeric value greater than 0.")
                        DlgFocus(4)
                    End If
                    DialogFunc = True
                Case "e_2"
                    If IsNumeric(DlgText(DlgItem)) Then
                        e_2 = CDbl(DlgText(DlgItem))
                        If e_2 <= 0 Then
                            App.feAppMessageBox(0, "Modulus must be a numeric value greater than 0.")
                            DlgFocus(5)
                        End If
                    Else
                        App.feAppMessageBox(0, "Modulus must be a numeric value greater than 0.")
                        DlgFocus(5)
                    End If
                    DialogFunc = True
                Case "d"
                    If IsNumeric(DlgText(DlgItem)) Then
                        d = CDbl(DlgText(DlgItem))
                        If d <= 0 Then
                            App.feAppMessageBox(0, "Diameter must be a numeric value greater than 0.")
                            DlgFocus(6)
                        End If
                    Else
                        App.feAppMessageBox(0, "Diameter must be a numeric value greater than 0.")
                        DlgFocus(6)
                    End If
                    DialogFunc = True
                Case "e_f"
                    If IsNumeric(DlgText(DlgItem)) Then
                        e_f = CDbl(DlgText(DlgItem))
                        If e_f <= 0 Then
                            App.feAppMessageBox(0, "Modulus must be a numeric value greater than 0.")
                            DlgFocus(7)
                        End If
                    Else
                        App.feAppMessageBox(0, "Modulus must be a numeric value greater than 0.")
                        DlgFocus(7)
                    End If
                    DialogFunc = True
                Case "nu_f"
                    If IsNumeric(DlgText(DlgItem)) Then
                        nu_f = CDbl(DlgText(DlgItem))
                        If nu_f <= 0 or nu_f >= 1 Then
                            App.feAppMessageBox(0, "Poisson's ratio must be a numeric value greater 0 and less than 1")
                            DlgFocus(8)
                        End If
                    Else
                        App.feAppMessageBox(0, "Poisson's ratio must be a numeric value greater 0 and less than 1")
                        DlgFocus(8)
                    End If
                    DialogFunc = True
            End Select

    End Select
End Function
