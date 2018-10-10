Attribute VB_Name = "Chart_H_"
Option Explicit

'Hardcoded: Data range, axis labels, For loops
Sub Chart_H(WS As Worksheet, top_left_corner As Range, title As String)
    
    'Select a blank cell
    WS.Activate
    WS.Range("B1").Select
    
    'Set standard attributes
    Dim cht As Object: Set cht = WS.Shapes.AddChart
    cht.Chart.Parent.Left = top_left_corner.Left
    cht.Chart.Parent.Top = top_left_corner.Top
    With cht.Chart.ChartArea.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = BACK_COLOR
        .Solid
    End With
    With cht.Chart.PlotArea.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = BACK_COLOR
        .Solid
    End With
    cht.Chart.HasTitle = True
    cht.Chart.ChartTitle.Text = title
    cht.Chart.ChartTitle.Font.Size = 14
    cht.Chart.ChartTitle.Font.Bold = True
    
    'Set custom attributes
    cht.Chart.SetSourceData Source:=WS.Range("T37:T132")
    cht.Chart.ChartType = xlLine
    cht.Chart.ChartArea.Height = 220
    cht.Chart.ChartArea.Width = 543
    cht.Chart.HasLegend = False
    
    'Additional customizations
    'Ensure axis starts at 0
    cht.Chart.Axes(xlValue).MinimumScale = 0
    
    'Thicken line
    cht.Chart.SeriesCollection(1).Format.Line.Weight = 3
    
    'Update axis labels and orient on diagonal
    cht.Chart.SeriesCollection(1).XValues = _
        Array("Apr/Avr 2011", "", "", "", "", "Sep 2011", "", "", "", "", "", "", _
              "Apr/Avr 2012", "", "", "", "", "Sep 2012", "", "", "", "", "", "", _
              "Apr/Avr 2013", "", "", "", "", "Sep 2013", "", "", "", "", "", "", _
              "Apr/Avr 2014", "", "", "", "", "Sep 2014", "", "", "", "", "", "", _
              "Apr/Avr 2015", "", "", "", "", "Sep 2015", "", "", "", "", "", "", _
              "Apr/Avr 2016", "", "", "", "", "Sep 2016", "", "", "", "", "", "", _
              "Apr/Avr 2017", "", "", "", "", "Sep 2017", "", "", "", "", "", "", _
              "Apr/Avr 2018", "", "", "", "", "Sep 2018", "", "", "", "", "", "Mar 2019")
    cht.Chart.Axes(xlCategory).TickLabels.Orientation = 45
    
    'Use school colour scheme (colour swatch from Michelle Craig on floor C4)
    Dim i As Long
    For i = 2 To 13
        cht.Chart.FullSeriesCollection(1).Points(i).Format.Line.ForeColor.RGB = RGB(164, 188, 196)
    Next
    For i = 14 To 25
        cht.Chart.FullSeriesCollection(1).Points(i).Format.Line.ForeColor.RGB = RGB(0, 82, 97)
    Next
    For i = 26 To 37
        cht.Chart.FullSeriesCollection(1).Points(i).Format.Line.ForeColor.RGB = RGB(164, 188, 196)
    Next
    For i = 38 To 49
        cht.Chart.FullSeriesCollection(1).Points(i).Format.Line.ForeColor.RGB = RGB(0, 82, 97)
    Next
    For i = 50 To 61
        cht.Chart.FullSeriesCollection(1).Points(i).Format.Line.ForeColor.RGB = RGB(164, 188, 196)
    Next
    For i = 62 To 73
        cht.Chart.FullSeriesCollection(1).Points(i).Format.Line.ForeColor.RGB = RGB(0, 82, 97)
    Next
    For i = 74 To 85
        cht.Chart.FullSeriesCollection(1).Points(i).Format.Line.ForeColor.RGB = RGB(164, 188, 196)
    Next
    For i = 86 To 96
        cht.Chart.FullSeriesCollection(1).Points(i).Format.Line.ForeColor.RGB = RGB(0, 82, 97)
    Next
    
    Set cht = Nothing
End Sub
