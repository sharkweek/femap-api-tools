' Title: Clear Analysis Manager Data.BAS
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 2306

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered
' trade­mark of Siemens Digital Industries Software Inc. or its subsidiaries in
' the United States and in other countries.
'
' Usage Notes
' -----------
' This script is used to clear the analysis manager for cases where Femap
' glitches and continues to show an analysis as running when it has actually
' bombed out.

Dim App as femap.model

Sub Main
    Set App = feFemap()

    Dim am As femap.AnalysisMgr
    Set am = App.feAnalysisMgr()

    am.ClearAnalysisQueue()
    am.ClearJobMonitorData()
End Sub
