Attribute VB_Name = "Chart_F_"
Option Explicit

'Hardcoded: Data range + textbox numbers
Sub Chart_F(WS As Worksheet, top_left_corner As Range, title As String)
    
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
    cht.Chart.SetSourceData Source:=WS.Range("R21:T33")
    cht.Chart.ChartType = xlLine
    cht.Chart.ChartArea.Height = 220
    cht.Chart.ChartArea.Width = 543
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
    
    'Display total registrations in textbox
    DoEvents
    WS.Shapes("TextBox1").OLEFormat.Object.Object.Value = "Total, " & LAST_YEAR & ":" & vbCr & Format(Application.Sum(WS.Range("S22:S33")), "#,##0")
    WS.Shapes("TextBox2").OLEFormat.Object.Object.Value = "Total, " & THIS_YEAR & ":" & vbCr & Format(Application.Sum(WS.Range("T22:T33")), "#,##0")
    DoEvents
    
    'Send chart to back of Z-order to show textbox
    cht.ZOrder msoSendToBack
    
    'Use school colour scheme (colour swatch from Michelle Craig on floor C4)
    cht.Chart.SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(164, 188, 196)
    cht.Chart.SeriesCollection(2).Format.Line.ForeColor.RGB = RGB(0, 82, 97)
    
    Set cht = Nothing
End Sub
