Attribute VB_Name = "Chart_F_"
Option Explicit

'Harcoded: Global averages: (0.24 for 2017-18, ? for 2018-19)
Sub Chart_F(WS As Worksheet, top_left_corner As Range, title As String, course_code As String)
    
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
    cht.Chart.SetSourceData Source:=WS.Range("AE6:AF7")
    cht.Chart.ChartType = xlColumnClustered
    cht.Chart.ChartArea.height = 226.5
    cht.Chart.ChartArea.width = 304
    cht.Chart.HasLegend = True
    
    'Additional customizations
    'Add additional data series for global averages
    cht.Chart.SeriesCollection.NewSeries
    cht.Chart.FullSeriesCollection(1).Name = course_code
    cht.Chart.FullSeriesCollection(2).Name = "Average All / Moyenne globale"
    cht.Chart.FullSeriesCollection(2).Values = Array(0.24, 0.24)
    
    'Ensure axis starts at 0
    cht.Chart.Axes(xlValue).MinimumScale = 0
    
    'Add data labels (formatted as percentage) and remove gridlines
    cht.Chart.SetElement (msoElementDataLabelOutSideEnd)
    cht.Chart.FullSeriesCollection(1).DataLabels.NumberFormat = "0%"
    cht.Chart.FullSeriesCollection(2).DataLabels.NumberFormat = "0%"
    cht.Chart.SetElement (msoElementPrimaryValueGridLinesNone)
    
    'Use school colour scheme (colour swatch from Michelle Craig on floor C4)
    cht.Chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(145, 8, 17)
    cht.Chart.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(84, 87, 90)
    
    Set cht = Nothing
End Sub
