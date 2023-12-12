' Title: Modes Function Abscissa As Frequency.BAS
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 2306

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered trade­mark
' of Siemens Digital Industries Software Inc. or its subsidiaries in the United
' States and in other countries.

' Usage Notes
' ----------
' When performing a modal analysis, Femap will automatically read in the mass
' fraction data from the OP2 (if requested) into function objects. These
' functions are listed as being "by frequency", but are in reality by mode
' number (i.e. 1, 2, 3,...). This script will create new function objects that
' are actually mapped by frequency. It works on any results function that is
' automatically created by Femap (i.e. modal mass fraction, participation
' factor, etc.).

' The script assumes that all necessary results sets are contained within a
' single, user-selected analysis study. If the number of output sets within
' the study does not match the selected functions, an error will be thrown.

' public declarations
Dim App As femap.model

Sub Main
    ' `Main` declarations
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Modes Function Abscissa As Frequency")
    App.feAppMessage(FCM_NORMAL, "Converts abscissa of modal mass fraction, sum, etc. from mode number to frequency.")

    Dim fn As femap.Function
    Set fn = App.feFunction()

    Dim os As femap.OutputSet
    Set os = App.feOutputSet()

    Dim study As femap.AnalysisStudy
    Set study = App.feAnalysisStudy()

    Dim os_set As femap.Set
    Set os_set = App.feSet()

    Dim fn_set As femap.Set
    Set fn_set = App.feSet()

    Dim os_sort As femap.SortSet
    Set os_sort = App.feSort()

    Dim abscissa As Variant
    Dim ordinate As Variant
    Dim x As Double
    Dim y As Double
    Dim z As Double
    Dim fn_count As Long
    Dim i As Long

    ' prompt user to select functions
    rc = fn_set.Select(FT_FUNCTION_DIR, True, "Select functions...")
    If rc = FE_CANCEL Then
        Exit Sub
    ElseIf rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "No functions selected. Routine canceled...")
    End If

    study.SelectID("Select Analysis Study...")
    study.OutputSets(os_set.ID)

    fn_set.Reset()
    Do While fn_set.Next()
        fn.Get(fn_set.CurrentID())
        fn.GetFunctionList(fn_count, abscissa, ordinate)

        ' check if function and study have same number of points
        If study.CountOutputSets() <> fn_count Then
            App.feAppMessage(FCM_ERROR, "Number of data points in function not equal to number of output sets in study.")
            Exit Sub
        End If

        ' sort studies by value
        os_set.Reset()
        os_sort.Clear()
        Do While os_set.Next()
            os.Get(os_set.CurrentID())
            os_sort.AddReal(os.ID, os.value, 0, 0)
        Loop
        os_sort.Sort(True)
        os_sort.Reset()

        ' create new abscissa
        i = 0
        Do While os_sort.Next()
            os_sort.GetReal(os_sort.Current(), x, y, z)
            abscissa(i) = x
            i += 1
        Loop

        ' add function to model
        fn.title = study.title + " - " + fn.title
        fn.PutFunctionList(fn_count, abscissa, ordinate)
        fn.Put(fn.NextEmptyID)
    Loop

End Sub
