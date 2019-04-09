Attribute VB_Name = "chartmodule"
Option Explicit

Private m_objChtEvents As New cls3DViewerChartEvents

Public Sub SelectChart(ByRef wks As Worksheet, ByRef rngPopupMsgs As Range)
    Dim objChart As Chart

    If wks.ChartObjects.Count > 0 Then
        Set m_objChtEvents = New cls3DViewerChartEvents

        Set objChart = wks.ChartObjects(1).Chart
        Set m_objChtEvents.ThreeDViewerclass = objChart

        'Set m_objChtEvents.PopupMsgs = rngPopupMsgs
    End If
End Sub

