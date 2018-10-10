Attribute VB_Name = "Chart_G_"
Option Explicit

Sub Chart_G(WS As Worksheet, top_left_corner As Range, title As String)
    
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
    cht.Chart.SetSourceData Source:=WS.Range("V21:Y33")
    cht.Chart.ChartType = xlLine
    cht.Chart.ChartArea.height = 220
    cht.Chart.ChartArea.width = 543
    cht.Chart.HasLegend = True
    
    'Additional customizations
    'Ensure axis starts at 0
    cht.Chart.Axes(xlValue).MinimumScale = 0
    
    'Make axis labels bilingual and orient on diagonal
    cht.Chart.SeriesCollection(1).XValues = Array("Apr/Avr", "May/Mai", "June/Juin", "Jul/Juil", "Aug/Août", "Sept", "Oct", "Nov", "Dec/Déc", "Jan", "Feb/Fév", "Mar")
    cht.Chart.Axes(xlCategory).TickLabels.Orientation = 45
    
    'Thicken lines
    cht.Chart.SeriesCollection(1).Format.Line.Weight = 3
    cht.Chart.SeriesCollection(2).Format.Line.Weight = 3
    
    'Use school colour scheme (colour swatch from Michelle Craig on floor C4)
    cht.Chart.SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(164, 188, 196)
    cht.Chart.SeriesCollection(2).Format.Line.ForeColor.RGB = RGB(0, 82, 97)
    cht.Chart.SeriesCollection(3).Format.Line.ForeColor.RGB = RGB(145, 8, 17)
    
    Set cht = Nothing
End Sub
