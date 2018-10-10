Attribute VB_Name = "MEGA"
Option Explicit

'Runtime June 01 2018: 418 sec
'Runtime July 03 2018: 389 sec
'Runtime July 06 2018: 542 sec
'Runtime August 01 2018: 434 sec

Sub MEGA()
    Application.ScreenUpdating = False
    Dim t As Double: t = Timer()
    
    Dim i As Long: i = 1
    Do While i < 102
        Call Main_MEGA(ThisWorkbook.Sheets("Department Names").Range("A1").Offset(i, 0))
        i = i + 1
    Loop
    
    Debug.Print Timer() - t
    Application.ScreenUpdating = True
End Sub


Sub Main_MEGA(my_selection As String)
    
    Dim WS As Worksheet: Set WS = ThisWorkbook.Sheets("Report")
    
    'Get dept code and dept name
    Dim splitResult As Variant: splitResult = Split(my_selection, ":", 2)
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
    Call PDF(WS, ThisWorkbook.path, dept_code)
    
End Sub
