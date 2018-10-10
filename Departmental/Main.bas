Attribute VB_Name = "Main"
Option Explicit

'Future proof: Update following four contstants + two tabs: "LSR_LAST_YEAR", "LSR_THIS_YEAR" + make Billing Dept Code changes listed in "Methodology" tab
Public Const LAST_YEAR As String = "2017-18"
Public Const THIS_YEAR As String = "2018-19"
'Q1/T1, Q2/T2, etc. but empty string if in between quarters as month names are too long
Public Const THIS_QUAR As String = ""
Public Const THIS_QUAR_FR As String = ""

Sub Main()
    Application.ScreenUpdating = False
    
    Dim WS As Worksheet: Set WS = ThisWorkbook.Sheets("Report")
    
    'Get dept code and dept name
    Dim splitResult As Variant: splitResult = Split(ThisWorkbook.Sheets("Instructions").Range("D3"), ":", 2)
    Dim dept_code As String: dept_code = splitResult(0)
    Dim dept_name As String: dept_name = Trim(splitResult(1))
    
    'Format page and generate raw data
    Call Pre_Format(WS, dept_name)
    Call Pre_Filter_SQL(dept_code)
    Call Calculations(WS, dept_code)
    
    'Create English charts
    Call Chart_A(WS, top_left_corner:=WS.Range("A4"), x_labels:=Array(""), y_labels:=Array(""), title:="A: Registrations by Month")
    Call Chart_B(WS, top_left_corner:=WS.Range("K4"), x_labels:=Array(""), y_labels:=Array(""), title:="B: Registrations by Business Type")
    Call Chart_C(WS, top_left_corner:=WS.Range("A19"), x_labels:=Array(""), y_labels:=Array(""), title:="C: Registrations to Leadership Programs, " & THIS_YEAR)
    Call Chart_D(WS, top_left_corner:=WS.Range("F19"), x_labels:=Array(""), y_labels:=Array(""), title:="D: No-Show Rate")
    Call Chart_E(WS, top_left_corner:=WS.Range("F26"), x_labels:=Array(""), y_labels:=Array(""), title:="E: Unique Learners per Year")
    Call Chart_F(WS, top_left_corner:=WS.Range("K19"), x_labels:=Array(""), y_labels:=Array(""), title:="F: Training Hours by Business Type")
    Call Chart_G(WS, top_left_corner:=WS.Range("A34"), x_labels:=Array(""), y_labels:=Array(""), title:="G: Top 10 Instructor-Led Courses, " & THIS_YEAR & " (Excluding Leadership Programs)")
    Call Chart_H(WS, top_left_corner:=WS.Range("I34"), x_labels:=Array(""), y_labels:=Array(""), title:="H: Top 10 Online Courses, " & THIS_YEAR & " (Excluding Leadership Programs)")
    
    'Create French charts
    Dim month_list_fr As Variant: month_list_fr = Array("Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre", "Janvier", "Février", "Mars")
    Dim business_type_list_fr As Variant: business_type_list_fr = Array("Événements", "Instructeur", "En ligne", "Total")
    Dim course_list_fr As Variant: course_list_fr = Array("PPS : G313", "PPG Phase 2 : G413", "PPG Phase 3 : G414", "PPG Phase 4 : G415", "PFD Phase 1 : E631", "PFD Phase 3 : E636", "PND Phase 1 : E632", "PND Phase 3 : E637", "NDG Phase 1 : E634", "NDG Phase 2 : E635")
    Call Chart_A(WS, top_left_corner:=WS.Range("A58"), x_labels:=month_list_fr, y_labels:=Array(""), title:="A : Inscriptions par mois")
    Call Chart_B(WS, top_left_corner:=WS.Range("K58"), x_labels:=Array(""), y_labels:=business_type_list_fr, title:="B : Inscriptions par type de livraison")
    Call Chart_C(WS, top_left_corner:=WS.Range("A73"), x_labels:=course_list_fr, y_labels:=Array(""), title:="C : Inscriptions aux programmes de leadership, " & THIS_YEAR)
    Call Chart_D(WS, top_left_corner:=WS.Range("F73"), x_labels:=Array(""), y_labels:=Array(""), title:="D : Taux d'absence")
    Call Chart_E(WS, top_left_corner:=WS.Range("F80"), x_labels:=Array(""), y_labels:=Array(""), title:="E : Apprenants uniques par année")
    Call Chart_F(WS, top_left_corner:=WS.Range("K73"), x_labels:=Array(""), y_labels:=business_type_list_fr, title:="F : Heures de formation par type de livraison")
    Call Chart_G(WS, top_left_corner:=WS.Range("A88"), x_labels:=WS.Range("W44:W53"), y_labels:=Array(""), title:="G : Top 10 des cours dirigés par un instructeur, " & THIS_YEAR & " (excluant les programmes de leadership)")
    Call Chart_H(WS, top_left_corner:=WS.Range("I88"), x_labels:=WS.Range("AA44:AA53"), y_labels:=Array(""), title:="H : Top 10 des cours en ligne, " & THIS_YEAR & " (excluant les programmes de leadership)")
    
    'Save As PDF
    If ThisWorkbook.Sheets("Instructions").OLEObjects("checkbox_pdf").Object.Value Then _
        Call PDF(WS, ThisWorkbook.path, dept_code)
    
    WS.Activate
    WS.Range("P1").Select
    
    Application.ScreenUpdating = True
End Sub


Sub Pre_Format(WS As Worksheet, dept_name As String)
    WS.Cells.ClearContents
    'Labels
    WS.Range("A2") = "Curriculum Usage Update, " & THIS_QUAR & " " & THIS_YEAR
    WS.Range("A56") = "Mise à jour sur la participation au cursus, " & THIS_QUAR_FR & " " & THIS_YEAR
    WS.Range("P2") = dept_name
    WS.Range("P56") = Application.VLookup(dept_name, ThisWorkbook.Sheets("Department Names").Columns("D:E"), 2, 0)
    WS.Range("N50") = "See Appendix 1 for Methodology"
    WS.Range("N51") = "Report generated on " & Date
    WS.Range("N52") = "Page 1/3"
    WS.Range("G51") = "Prepared by Curriculum Management"
    WS.Range("N104") = "Consulter l'Annexe 1 pour la méthodologie"
    WS.Range("N105") = "Rapport généré le " & Date
    WS.Range("N106") = "Page 2/3"
    WS.Range("G105") = "Préparé par Gestion du Curriculum"
    'Clear previous charts
    On Error Resume Next
        WS.ChartObjects.Delete
    On Error GoTo 0
End Sub


'Filter data by dept_code before hand to speed up SQL queries
Sub Pre_Filter_SQL(dept_code As String)
    ThisWorkbook.Sheets("LAST_YEAR").Cells.Clear
    ThisWorkbook.Sheets("LSR_LAST_YEAR").Rows("1:1").AutoFilter Field:=39, Criteria1:=dept_code
    ThisWorkbook.Sheets("LSR_LAST_YEAR").AutoFilter.Range.Copy
    ThisWorkbook.Sheets("LAST_YEAR").Range("A1").PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    On Error Resume Next
        ThisWorkbook.Sheets("LSR_LAST_YEAR").ShowAllData
    On Error GoTo 0
    
    ThisWorkbook.Sheets("THIS_YEAR").Cells.Clear
    ThisWorkbook.Sheets("LSR_THIS_YEAR").Rows("1:1").AutoFilter Field:=39, Criteria1:=dept_code
    ThisWorkbook.Sheets("LSR_THIS_YEAR").AutoFilter.Range.Copy
    ThisWorkbook.Sheets("THIS_YEAR").Range("A1").PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    On Error Resume Next
        ThisWorkbook.Sheets("LSR_THIS_YEAR").ShowAllData
    On Error GoTo 0
End Sub


Sub Calculations(WS As Worksheet, dept_code As String)
    Call registrations_by_month(WS.Range("W4"))
    Call registrations_by_business_type(WS.Range("AA4"))
    Call registrations_to_leadership_programs(WS.Range("W18"))
    Call no_show_rate(WS.Range("AA18"))
    Call unique_learners(WS.Range("AA23"))
    Call training_hours_by_business_type(WS.Range("AE18"))
    Call top_10(WS.Range("W33"), "Instructor-Led")
    Call top_10(WS.Range("AA33"), "Online")
End Sub
