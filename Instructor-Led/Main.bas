Attribute VB_Name = "Main"
Option Explicit

Public Const LAST_YEAR As String = "2017-18"
Public Const THIS_YEAR As String = "2018-19"
Public Const BACK_COLOR As Long = 16777215

'To update with new data:
'Update constants LAST_YEAR and THIS_YEAR
'Update LSR, L1SR tabs
'Update hardcoding (indicated as start of each sub)
'Don't forget global averages in Charts F, I, L (updated twice yearly in April and October)

Sub Main()
    Application.ScreenUpdating = False
    
    Dim WS As Worksheet: Set WS = ThisWorkbook.Sheets("Report")
    
    'Get course code and course title
    Dim course_code As String, course_title As String, splitResult As Variant
    splitResult = Split(ThisWorkbook.Sheets("Instructions").Range("D3"), ":", 2)
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
    If ThisWorkbook.Sheets("Instructions").OLEObjects("checkbox_pdf").Object.Value Then _
        Call PDF(WS, ThisWorkbook.Path, course_code)
    
    WS.Activate
    WS.Range("R1").Select
    
    Application.ScreenUpdating = True
End Sub


Sub Pre_Format(WS As Worksheet, course_code As String, course_title As String)
    WS.Cells.ClearContents
    WS.Range("A2") = "Curriculum Usage Update, " & MonthName(Month(Date)) & " " & THIS_YEAR
    WS.Range("R2") = "Mise à jour sur l'usage du curriculum, " & Application.VLookup(MonthName(Month(Date)), ThisWorkbook.Sheets("Course Codes").Columns("J:K"), 2, 0) & " " & THIS_YEAR
    WS.Range("A3") = course_code & ": " & course_title
    WS.Range("R3") = course_code & ": " & Application.VLookup(course_code, ThisWorkbook.Sheets("Course Codes").Columns("C:E"), 3, 0)
    WS.Range("P52") = "See Appendix 1 for Methodology / Consulter l'Annexe 1 pour la méthodologie"
    WS.Range("P53") = "Report generated on / Rapport généré le " & Date
    WS.Range("P54") = "Page 1/1"
    'Clear previous charts
    On Error Resume Next
        WS.ChartObjects.Delete
    On Error GoTo 0
End Sub


'Filter data by course_code before hand to speed up SQL queries
Sub Pre_Filter_SQL(course_code As String)
    ThisWorkbook.Sheets("LAST_YEAR").Cells.Clear
    ThisWorkbook.Sheets("LSR" & Left(LAST_YEAR, 4)).Rows("1:1").AutoFilter Field:=2, Criteria1:=course_code
    ThisWorkbook.Sheets("LSR" & Left(LAST_YEAR, 4)).AutoFilter.Range.Copy
    ThisWorkbook.Sheets("LAST_YEAR").Range("A1").PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    On Error Resume Next
        ThisWorkbook.Sheets("LSR" & Left(LAST_YEAR, 4)).ShowAllData
    On Error GoTo 0
    
    ThisWorkbook.Sheets("THIS_YEAR").Cells.Clear
    ThisWorkbook.Sheets("LSR" & Left(THIS_YEAR, 4)).Rows("1:1").AutoFilter Field:=2, Criteria1:=course_code
    ThisWorkbook.Sheets("LSR" & Left(THIS_YEAR, 4)).AutoFilter.Range.Copy
    ThisWorkbook.Sheets("THIS_YEAR").Range("A1").PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    On Error Resume Next
        ThisWorkbook.Sheets("LSR" & Left(THIS_YEAR, 4)).ShowAllData
    On Error GoTo 0
End Sub


Sub Calculations(WS As Worksheet, course_code As String)
    Call first_offering(WS.Range("V5"), course_code)
    Call tombstone_data(WS.Range("A7"), course_code)
    Call cheat_total(WS)
    Call level_1_results(WS.Range("A15"), course_code)
    Call top_5_departments(WS.Range("J5"))
    Call top_5_classifications(WS.Range("J13"))
    Call offerings_cancelled(WS.Range("AE5"))
    Call offerings_per_month(WS.Range("V21"))
    Call offerings_per_region(WS.Range("AA21"))
    Call average_no_shows(WS.Range("AE21"))
    Call cumulative_unique_learners(loc:=WS.Range("V36"), course_code:=course_code)
    Call offerings_per_language(WS.Range("AA36"))
    Call average_class_size(WS.Range("AE36"))
End Sub
