Attribute VB_Name = "Chart_H_"
Option Explicit

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
    cht.Chart.SetSourceData Source:=WS.Range("AA22:AB27")
    cht.Chart.ChartType = xlColumnClustered
    cht.Chart.ChartArea.height = 220
    cht.Chart.ChartArea.width = 319
    cht.Chart.HasLegend = False
    
    'Additional customizations
    'Ensure axis starts at 0
    cht.Chart.Axes(xlValue).MinimumScale = 0
    
    'Add data labels and remove gridlines
    cht.Chart.SetElement (msoElementDataLabelOutSideEnd)
    cht.Chart.SetElement (msoElementPrimaryValueGridLinesNone)
    
    'Make axis labels bilingual and orient on diagonal
    cht.Chart.SeriesCollection(1).XValues = Array("Atl", "NCR/RCN", "Ontario", "Pac", "Prairie", "Québec")
    cht.Chart.Axes(xlCategory).TickLabels.Orientation = 45
    
    'Use school colour scheme (colour swatch from Michelle Craig on floor C4)
    cht.Chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(0, 82, 97)
    
    Set cht = Nothing
End Sub
