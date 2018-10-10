Attribute VB_Name = "Calculations_"
Option Explicit

'Hardcoded: None
Sub registrations_by_month(loc As Range)
    'Labels
    loc.Offset(0, 0) = "Registrations by Month": loc.Offset(0, 1) = LAST_YEAR: loc.Offset(0, 2) = THIS_YEAR
    Dim month_list As Variant: month_list = Array("April", "May", "June", "July", "August", "September", "October", "November", "December", "January", "February", "March")
    loc.Offset(1, 0).Resize(UBound(month_list) + 1, 1) = Application.Transpose(month_list)
    'Formulæ
    'LAST_YEAR
    Dim i As Long
    For i = LBound(month_list) To UBound(month_list)
        loc.Offset(i + 1, 1).Formula = "=COUNTIFS(LAST_YEAR!$AD:$AD,""Confirmed"",LAST_YEAR!$K:$K,""" & month_list(i) & """)"
    Next
    'THIS_YEAR
    For i = LBound(month_list) To UBound(month_list)
        loc.Offset(i + 1, 2).Formula = "=COUNTIFS(THIS_YEAR!$AD:$AD,""Confirmed"",THIS_YEAR!$K:$K,""" & month_list(i) & """)"
    Next
End Sub


'Hardcoded: None
Sub registrations_by_business_type(loc As Range)
    'Labels
    loc.Offset(0, 0) = "Registrations by Business Type": loc.Offset(0, 1) = LAST_YEAR: loc.Offset(0, 2) = THIS_YEAR: loc.Offset(4, 0) = "Total"
    Dim biz_list As Variant: biz_list = Array("Events", "Instructor-Led", "Online")
    loc.Offset(1, 0).Resize(UBound(biz_list) + 1, 1) = Application.Transpose(biz_list)
    'Formulæ
    'LAST_YEAR
    Dim i As Long
    For i = LBound(biz_list) To UBound(biz_list)
        loc.Offset(i + 1, 1).Formula = "=COUNTIFS(LAST_YEAR!$AD:$AD,""Confirmed"",LAST_YEAR!$O:$O,""" & biz_list(i) & """)"
    Next
    'THIS_YEAR
    For i = LBound(biz_list) To UBound(biz_list)
        loc.Offset(i + 1, 2).Formula = "=COUNTIFS(THIS_YEAR!$AD:$AD,""Confirmed"",THIS_YEAR!$O:$O,""" & biz_list(i) & """)"
    Next
    'Totals
    loc.Offset(4, 1) = loc.Offset(1, 1) + loc.Offset(2, 1) + loc.Offset(3, 1)
    loc.Offset(4, 2) = loc.Offset(1, 2) + loc.Offset(2, 2) + loc.Offset(3, 2)
End Sub


'Hardcoded: None
Sub registrations_to_leadership_programs(loc As Range)
    'Labels
    loc.Offset(0, 0) = "Registrations to Leadership Programs": loc.Offset(0, 1) = THIS_YEAR
    Dim course_list As Variant: course_list = Array("G313", "G413", "G414", "G415", "E631", "E636", "E632", "E637", "E634", "E635")
    loc.Offset(1, 0) = "SDP: " & course_list(0): loc.Offset(2, 0) = "MDP Phase 2: " & course_list(1): loc.Offset(3, 0) = "MDP Phase 3: " & course_list(2): loc.Offset(4, 0) = "MDP Phase 4: " & course_list(3): loc.Offset(5, 0) = "ADP Phase 1: " & course_list(4): loc.Offset(6, 0) = "ADP Phase 3: " & course_list(5): loc.Offset(7, 0) = "NDP Phase 1: " & course_list(6): loc.Offset(8, 0) = "NDP Phase 3: " & course_list(7): loc.Offset(9, 0) = "NDG Phase 1: " & course_list(8): loc.Offset(10, 0) = "NDG Phase 2: " & course_list(9)
    'Formulæ
    'THIS_YEAR
    Dim i As Long
    For i = LBound(course_list) To UBound(course_list)
        loc.Offset(i + 1, 1).Formula = "=COUNTIFS(THIS_YEAR!$AD:$AD,""Confirmed"",THIS_YEAR!$B:$B,""" & course_list(i) & """)"
    Next
End Sub


'Hardcoded: Numerator
Sub no_show_rate(loc As Range)
    'Labels
    loc.Offset(0, 0) = "No-Show Rate": loc.Offset(0, 1) = LAST_YEAR: loc.Offset(0, 2) = THIS_YEAR
    'Formulæ
    'LAST_YEAR
    loc.Offset(1, 1).Formula = "=SUMIFS(LAST_YEAR!$AF:$AF,LAST_YEAR!$H:$H,{""Delivered - Normal"",""Open - Normal""})"
    Dim denominator As Long: denominator = loc.Offset(1, 1) + Application.Sum(ThisWorkbook.Sheets("Report").Range("AB5:AB6"))
    If denominator Then loc.Offset(1, 1) = loc.Offset(1, 1) / denominator Else loc.Offset(1, 1) = 0
    'THIS_YEAR
    loc.Offset(1, 2).Formula = "=SUMIFS(THIS_YEAR!$AF:$AF,THIS_YEAR!$H:$H,{""Delivered - Normal"",""Open - Normal""})"
    denominator = loc.Offset(1, 2) + Application.Sum(ThisWorkbook.Sheets("Report").Range("AC5:AC6"))
    If denominator Then loc.Offset(1, 2) = loc.Offset(1, 2) / denominator Else loc.Offset(1, 2) = 0
End Sub


'Hardcoded: None
Sub unique_learners(loc As Range)
    'Labels
    loc.Offset(0, 0) = "Unique Learners": loc.Offset(0, 1) = LAST_YEAR: loc.Offset(0, 2) = THIS_YEAR
    'Formulæ
    'LAST_YEAR
    Dim myQuery As String
    myQuery = "SELECT COUNT(*) FROM (SELECT DISTINCT [Student ID] FROM [LAST_YEAR$] WHERE [Reg Status] = 'Confirmed');"
    Call SQL(query:=myQuery, result_location:=loc.Offset(1, 1), header:=False, as_array:=False)
    'THIS_YEAR
    myQuery = "SELECT COUNT(*) FROM (SELECT DISTINCT [Student ID] FROM [THIS_YEAR$] WHERE [Reg Status] = 'Confirmed');"
    Call SQL(query:=myQuery, result_location:=loc.Offset(1, 2), header:=False, as_array:=False)
End Sub


'Hardcoded: None
Sub training_hours_by_business_type(loc As Range)
    'Labels
    loc.Offset(0, 0) = "Training Hours by Business Type": loc.Offset(0, 1) = LAST_YEAR: loc.Offset(0, 2) = THIS_YEAR: loc.Offset(4, 0) = "Total"
    Dim biz_list As Variant: biz_list = Array("Events", "Instructor-Led", "Online")
    loc.Offset(1, 0).Resize(UBound(biz_list) + 1, 1) = Application.Transpose(biz_list)
    'Formulæ
    'LAST_YEAR
    Dim i As Long
    For i = LBound(biz_list) To UBound(biz_list)
        loc.Offset(i + 1, 1).Formula = "=ROUND(SUMIFS(LAST_YEAR!$M:$M,LAST_YEAR!$AD:$AD,""Confirmed"",LAST_YEAR!$O:$O,""" & biz_list(i) & """),0)"
    Next
    'THIS_YEAR
    For i = LBound(biz_list) To UBound(biz_list)
        loc.Offset(i + 1, 2).Formula = "=ROUND(SUMIFS(THIS_YEAR!$M:$M,THIS_YEAR!$AD:$AD,""Confirmed"",THIS_YEAR!$O:$O,""" & biz_list(i) & """),0)"
    Next
    'Totals
    loc.Offset(4, 1) = loc.Offset(1, 1) + loc.Offset(2, 1) + loc.Offset(3, 1)
    loc.Offset(4, 2) = loc.Offset(1, 2) + loc.Offset(2, 2) + loc.Offset(3, 2)
End Sub


'Hardcoded: None
Sub top_10(loc As Range, business_type As String)
    'Labels
    loc.Offset(0, 0) = "Top 10 " & business_type: loc.Offset(0, 1) = THIS_YEAR
    'Formulæ
    'THIS_YEAR
    Dim myQuery As String
    myQuery = "SELECT TOP 10 [Course Title], COUNT([Course Title]) FROM [THIS_YEAR$] WHERE [Reg Status] = 'Confirmed' AND [Business Type] = '" & business_type & "' AND [Course Code] NOT IN ('G313','G413','G414','G415','E631','E636','E632','E637','E634','E635','E800','E801') GROUP BY [Course Title] ORDER BY 2 DESC"
    Call SQL(query:=myQuery, result_location:=loc.Offset(1, 0), header:=False, as_array:=False)
    'Clear cells below in case of overflow
    loc.Offset(11, 0).Resize(10, 2).ClearContents
    'If <10 top courses, fill in blanks to show it explicitly
    Dim i As Long
    For i = 1 To 10
        If IsEmpty(loc.Offset(i, 0)) Then
            loc.Offset(i, 0) = "N/A"
            loc.Offset(i, 1) = "0"
        End If
    Next
    'Move course code from end of title to beginning via RegExp
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    Dim rePattern As String: rePattern = "[(\[]{1}[A-Z]{1}\d{3}[)\]]{1}"
    With re
        re.IgnoreCase = True
        re.Global = True
        re.MultiLine = True
        re.Pattern = rePattern
    End With
    Dim foundCode As Object
    For i = 1 To 10
        Set foundCode = re.Execute(loc.Offset(i, 0))
        If foundCode.Count Then
            loc.Offset(i, 0) = re.Replace(loc.Offset(i, 0), "")
            loc.Offset(i, 0) = Trim(foundCode(0) & " " & loc.Offset(i, 0))
        End If
    Next
    Set re = Nothing
    Set foundCode = Nothing
    'Add French titles
    For i = 1 To 10
        loc.Offset(i + 10, 0).Formula = "=IFNA(VLOOKUP(""" & loc.Offset(i, 0) & """,'Course Names'!$A:$B,2,0),""" & loc.Offset(i, 0) & """)"
    Next
End Sub
