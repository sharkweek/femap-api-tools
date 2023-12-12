' Title: Save Images From Multiple Output Sets.BAS
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 2022.2

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

    App.feAppMessage(FCM_COMMAND, "Save Images From Multiple Output Sets")

    Dim rc As Long

    Dim os_set As femap.Set
    Set os_set = App.feSet()

    Dim v As femap.View
    Set v = App.feView()

    Dim os As femap.OutputSet
    Set os = App.feOutputSet()

    Dim ans_study As femap.AnalysisStudy
    Set ans_study = App.feAnalysisStudy()

    Dim view_id As Long
    Dim file_type As String
    Dim format As Long

    ' prompt user to select output sets
    rc = os_set.Select(FT_OUT_CASE, True, "Select output sets...")
    If rc = FE_CANCEL Then
        Exit Sub
    ElseIf rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "No output sets selected. Routine canceled...")
    End If

    ' get current view info
    App.feAppGetActiveView(view_id)
    v.Get(view_id)
    If v.Deformed = 2 Then
        file_type = ".gif"
        format = 10
    Else
        file_type = ".png"
        format = 12
    End If

    Do While os_set.Next()
        os.Get(os_set.CurrentID)
        ans_study.Get(os.study)
        v.OutputSet = os.ID
        v.Put(v.ID)
        App.feViewRegenerate(v.ID)

        file_name = ans_study.title & " - " & os.title & file_type
        App.feFilePictureSave2(0, 0, format, file_name)
    Loop

End Sub
