' Title: Rename Modal Results by Frequency.BAS
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 2301

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered trade­mark
' of Siemens Digital Industries Software Inc. or its subsidiaries in the United
' States and in other countries.

' Warning
' ----------
' This script renames by copying output sets in place. Selecting a study with
' many embedded output sets may take a very long time. To avoid this, ensure
' that the output sets being renamed are from an attached, memory mapped OP2
' or other output file.

' public declarations
Dim App As femap.model

Sub Main
    ' `Main` declarations
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Rename Modal Results by Frequency")

    Dim os As femap.OutputSet
    Set os = App.feOutputSet()

    Dim study As femap.AnalysisStudy
    Set study = App.feAnalysisStudy()

    Dim os_set As femap.Set
    Set os_set = App.feSet()

    study.SelectID("Select Analysis Study...")
    study.OutputSets(os_set.ID)
    Select Case study.AnalysisType
        Case 2, 4, 6
            os_set.Reset()
            Do While os_set.Next()
                os.Get(os_set.CurrentID)
                os.title = CStr(Round(os.value, 2)) & " Hz"
                os.Put(os_set.CurrentID)
            Loop
        Case Else
            App.feAppMessage(FCM_ERROR, "Analysis Study is not a modes/eigenvalue analysis. Routine canceled...")
    End Select

End Sub
