' Title: Functions from Chart Data Series.BAS
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

    App.feAppMessage(FCM_COMMAND, "Functions from Chart Data Series")

    Dim rc As Long

    Dim series_set As femap.Set
    Set series_set = App.feSet()

    Dim series As femap.ChartSeries
    Set series = App.feChartSeries()

    Dim func As femap.Function
    Set func = App.feFunction()

    Dim numPoints As Long
    Dim xyValues As Variant
    Dim xRange As Variant
    Dim yRange As Variant
    Dim numPoints_idx As Long
    Dim x_values() As Double
    Dim y_values() As Double

    ' prompt user to select entities
    rc = series_set.Select(FT_CHART_SERIES, True, "Select chart data series...")
    If rc = FE_CANCEL Then
        Exit Sub
    ElseIf rc = FE_NOT_EXIST Then
        App.feAppMessage(FCM_ERROR, "No entities selected. Routine canceled...")
    End If

    Do While series_set.Next()
        series.Get(series_set.CurrentID)
        series.GetXY(numPoints, xyValues, xRange, yRange)
        numPoints_idx = numPoints - 1
        ReDim x_values(numPoints/2)
        ReDim y_values(numPoints/2)

        For i = 0 to numPoints_idx
            idx = i - i Mod 2
            If i Mod 2 = 0 Then  ' even
                x_values(idx) = xyValues(i)
            ElseIf i Mod 2 = 1 Then  ' odd
                y_values(idx) = xyValues(i)
            End If
        Next i

        func.title = series.title
        func.PutFunctionList(numPoints/2, x_values, y_values)
        func.Put(func.NextEmptyID())
    Loop

    App.feViewRegenerate(0)

End Sub