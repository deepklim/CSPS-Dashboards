Attribute VB_Name = "Chart_A_"
Option Explicit

Sub Chart_A(WS As Worksheet, top_left_corner As Range, x_labels As Variant, y_labels As Variant, title As String)
    
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
    cht.Chart.SetSourceData Source:=WS.Range("W4:Y16")
    cht.Chart.ChartType = xlLine
    cht.Chart.ChartArea.height = 220
    cht.Chart.ChartArea.width = 767.5
    cht.Chart.HasLegend = False
    
    'Additional customizations
    'Add data table and remove gridlines
    cht.Chart.SetElement (msoElementDataTableWithLegendKeys)
    cht.Chart.SetElement (msoElementPrimaryValueGridLinesNone)
    
    'Thicken line for LAST_YEAR
    cht.Chart.SeriesCollection(1).Format.Line.Weight = 4
    
    'Show THIS_YEAR as columns
    cht.Chart.SeriesCollection(2).ChartType = xlColumnClustered
    
    'Use school colour scheme (colour swatch from Michelle Craig on floor C4)
    cht.Chart.SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(164, 188, 196)
    cht.Chart.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(0, 82, 97)
    
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
