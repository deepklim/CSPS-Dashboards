Attribute VB_Name = "Chart_D_"
Option Explicit

Sub Chart_D(WS As Worksheet, top_left_corner As Range, x_labels As Variant, y_labels As Variant, title As String)
    
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
    cht.Chart.SetSourceData Source:=WS.Range("AB18:AC19")
    cht.Chart.ChartType = xlBarClustered
    cht.Chart.ChartArea.height = 100
    cht.Chart.ChartArea.width = 381
    cht.Chart.HasLegend = False
    
    'Additional customizations
    'Remove gridlines and add data labels
    cht.Chart.SetElement (msoElementPrimaryValueGridLinesNone)
    cht.Chart.SetElement (msoElementDataLabelOutSideEnd)
    
    'Reverse plot order
    cht.Chart.Axes(xlCategory).ReversePlotOrder = True
    
    'Remove axis labels
    cht.Chart.Axes(xlValue).Delete
    
    'Adjust axis labels
    cht.Chart.Axes(xlValue).MinimumScale = 0
    
    'Add additional data series for global averages
    cht.Chart.SeriesCollection.NewSeries
    cht.Chart.FullSeriesCollection(2).Values = Array(0.062, 0.062)
    cht.Chart.FullSeriesCollection(1).DataLabels.NumberFormat = "0.0%"
    cht.Chart.FullSeriesCollection(2).DataLabels.NumberFormat = "0.0%"
    
    'Send chart to back of Z-order to show textbook
    cht.ZOrder msoSendToBack
    
    'Tweak colours: LAST_YEAR = msoThemeColorAccent2, THIS_YEAR = msoThemeColorAccent1
    cht.Chart.FullSeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB = RGB(164, 188, 196)
    cht.Chart.FullSeriesCollection(1).Points(2).Format.Fill.ForeColor.RGB = RGB(0, 82, 97)
    cht.Chart.FullSeriesCollection(2).Points(1).Format.Fill.ForeColor.RGB = RGB(164, 188, 196)
    cht.Chart.FullSeriesCollection(2).Points(2).Format.Fill.ForeColor.RGB = RGB(0, 82, 97)
    
    'Add gap between neighbouring bars to make more readable
    cht.Chart.ChartGroups(1).Overlap = -30
    
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
