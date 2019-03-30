Attribute VB_Name = "CreateChart"
Sub ChartArray()

    Dim x(0 To 1000, 0 To 0) As Double
    Dim y(0 To 1000, 0 To 0) As Double
    x(0, 0) = 0
    y(0, 0) = 0
    For i = 1 To 1000
        x(i, 0) = i
        y(i, 0) = y(i - 1, 0) + WorksheetFunction.NormSInv(Rnd())
    Next i

    Charts.Add
    ActiveChart.ChartType = xlXYScatterLinesNoMarkers
    With ActiveChart.SeriesCollection
        If .Count = 0 Then .NewSeries
        If Val(Application.Version) >= 12 Then
            .Item(1).Values = y
            .Item(1).XValues = x
        Else
            .Item(1).Select
            Names.Add "_", x
            ExecuteExcel4Macro "series.x(!_)"
            Names.Add "_", y
            ExecuteExcel4Macro "series.y(,!_)"
            Names("_").Delete
        End If
    End With
    ActiveChart.ChartArea.Select

End Sub

