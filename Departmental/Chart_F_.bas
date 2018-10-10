Attribute VB_Name = "Chart_F_"
Option Explicit

Sub Chart_F(WS As Worksheet, top_left_corner As Range, x_labels As Variant, y_labels As Variant, title As String)
    
    'Select a blank cell
    WS.Activate
    WS.Range("B1").Select
    
    'Set standard attributes
    Dim cht As Object: Set cht = WS.Shapes.AddChart
    cht.Chart.Parent.Left = top_left_corner.Left
    cht.Chart.Parent.Top = top_left_corner.Top
    With cht.Chart.ChartArea.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 255, 255)
        .ForeColor.Brightness = 0.75
        .Solid
    End With
    With cht.Chart.PlotArea.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 255, 255)
        .ForeColor.Brightness = 0.75
        .Solid
    End With
    cht.Chart.HasTitle = True
    cht.Chart.ChartTitle.Text = title
    cht.Chart.ChartTitle.Font.Size = 14
    cht.Chart.ChartTitle.Font.Bold = True
    
    'Set custom attributes
    cht.Chart.SetSourceData Source:=WS.Range("AE18:AG19")
    cht.Chart.ChartType = xlBarStacked
    cht.Chart.ChartArea.height = 220
    cht.Chart.ChartArea.width = 458.5
    cht.Chart.HasLegend = False
    
    'Additional customizations
    'Add remaining data (done seperately for different appearance)
    cht.Chart.SeriesCollection.NewSeries
    cht.Chart.FullSeriesCollection(2).Name = "=Report!$AE$20"
    cht.Chart.FullSeriesCollection(2).Values = "=Report!$AF$20:$AG$20"
    cht.Chart.SeriesCollection.NewSeries
    cht.Chart.FullSeriesCollection(3).Name = "=Report!$AE$21"
    cht.Chart.FullSeriesCollection(3).Values = "=Report!$AF$21:$AG$21"
    cht.Chart.SeriesCollection.NewSeries
    cht.Chart.FullSeriesCollection(4).Name = "=Report!$AE$22"
    cht.Chart.FullSeriesCollection(4).Values = "=Report!$AF$22:$AG$22"
    
    'Add data table
    cht.Chart.SetElement (msoElementDataTableWithLegendKeys)
    
    'Remove gridlines
    cht.Chart.SetElement (msoElementPrimaryValueGridLinesNone)
    
    'Reverse plot order
    cht.Chart.Axes(xlCategory).ReversePlotOrder = True
    
    'Remove axis labels
    cht.Chart.Axes(xlValue).Delete
    
    'Adjust axis so Total visible on in data table
    cht.Chart.Axes(xlValue).MinimumScale = 0
    cht.Chart.Axes(xlValue).MaximumScale = Application.Max(WS.Range("AF22:AG22")) * 1.2
    
    'Use school colour scheme (colour swatch from Michelle Craig on floor C4)
    cht.Chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(84, 87, 90)
    cht.Chart.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(145, 8, 17)
    cht.Chart.SeriesCollection(3).Format.Fill.ForeColor.RGB = RGB(86, 117, 130)
    cht.Chart.SeriesCollection(4).Format.Fill.Visible = msoFalse
    
    'Allow custom X- and Y-axis labels for French charts
    If Not x_labels(0) = "" Then
        cht.Chart.FullSeriesCollection(1).XValues = x_labels
    End If
    If Not y_labels(0) = "" Then
        Dim i As Long, my_series As Series
        For Each my_series In cht.Chart.SeriesCollection
            my_series.Name = y_labels(i)
            i = i + 1
        Next
    End If
    
    Set cht = Nothing
End Sub
