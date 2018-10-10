Attribute VB_Name = "MEGA"
Option Explicit
'Note: Don't save after running MEGA
'Runtime June 01, 2018: 1315 sec
'Runtime July 03, 2018: 1554 sec
'Runtime August 01, 2018: 1014 sec
'Runtime September 04, 2018: 1030 sec

Sub MEGA()
    Application.ScreenUpdating = False
    Dim t As Double: t = Timer()
    
    Dim i As Long: i = 1
    Do While i < 96
        Call Main_MEGA(ThisWorkbook.Sheets("Course Codes").Range("A1").Offset(i, 0))
        i = i + 1
    Loop
    
    Debug.Print Timer() - t
    Application.ScreenUpdating = True
End Sub


Sub Main_MEGA(my_selection As String)
    
    Dim WS As Worksheet: Set WS = ThisWorkbook.Sheets("Report")
    
    'Get course code and course title
    Dim course_code As String, course_title As String, splitResult As Variant
    splitResult = Split(my_selection, ":", 2)
    course_code = splitResult(0)
    course_title = Trim(splitResult(1))
    
    'Format page and generate raw data
    Call Pre_Format(WS, course_code, course_title)
    Call Pre_Filter_SQL(course_code)
    Call Calculations(WS, course_code)
    
    'Create charts
    Call Chart_F(WS, top_left_corner:=WS.Range("O5"), title:="F: % of Offerings Cancelled / % d'offres annulées", course_code:=course_code)
    Call Chart_G(WS, top_left_corner:=WS.Range("A21"), title:="G: Offerings per Month / Offres par mois")
    Call Chart_H(WS, top_left_corner:=WS.Range("I21"), title:="H: " & THIS_YEAR & ": Offerings per Region / Offres par région")
    Call Chart_I(WS, top_left_corner:=WS.Range("O21"), title:="I: Average No-Shows per Offering / Absences moyennes par offre", course_code:=course_code)
    Call Chart_J(WS, top_left_corner:=WS.Range("A36"), title:="J: Cumulative Unique Learners / Apprenants uniques cumulatif")
    Call Chart_K(WS, top_left_corner:=WS.Range("I36"), title:="K: Offerings per Language / Offres par langue")
    Call Chart_L(WS, top_left_corner:=WS.Range("O36"), title:="L: Average Class Size / Effectif de classe moyen", course_code:=course_code)
    
    'Save as PDF
    Call PDF(WS, ThisWorkbook.Path, course_code)
    
End Sub
