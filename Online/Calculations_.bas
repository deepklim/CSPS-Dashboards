Attribute VB_Name = "Calculations_"
Option Explicit

'Hardcoded: One row per LSR, limit of For loop
Sub first_reg(loc As Range, course_code As String)
    '.FormulaArray is VBA equivalent of Ctrl+Shift+Enter (i.e. array) function
    'Note intersting use of * operator: used in place of AND function within array functions
    loc.Offset(0, 0) = "First Confirmed Registration in Each LSR"
    loc.Offset(1, 0).FormulaArray = "=MIN(IF(('LSR2011'!$B:$B=""" & course_code & """)*('LSR2011'!$AA:$AA=""Confirmed""),'LSR2011'!$Y:$Y))"
    loc.Offset(2, 0).FormulaArray = "=MIN(IF(('LSR2012'!$B:$B=""" & course_code & """)*('LSR2012'!$AA:$AA=""Confirmed""),'LSR2012'!$Y:$Y))"
    loc.Offset(3, 0).FormulaArray = "=MIN(IF(('LSR2013'!$B:$B=""" & course_code & """)*('LSR2013'!$AA:$AA=""Confirmed""),'LSR2013'!$Y:$Y))"
    loc.Offset(4, 0).FormulaArray = "=MIN(IF(('LSR2014'!$B:$B=""" & course_code & """)*('LSR2014'!$AA:$AA=""Confirmed""),'LSR2014'!$Y:$Y))"
    loc.Offset(5, 0).FormulaArray = "=MIN(IF(('LSR2015'!$B:$B=""" & course_code & """)*('LSR2015'!$AA:$AA=""Confirmed""),'LSR2015'!$Y:$Y))"
    loc.Offset(6, 0).FormulaArray = "=MIN(IF(('LSR2016'!$B:$B=""" & course_code & """)*('LSR2016'!$AA:$AA=""Confirmed""),'LSR2016'!$Y:$Y))"
    loc.Offset(7, 0).FormulaArray = "=MIN(IF(('LSR2017'!$B:$B=""" & course_code & """)*('LSR2017'!$AA:$AA=""Confirmed""),'LSR2017'!$Y:$Y))"
    loc.Offset(8, 0).FormulaArray = "=MIN(IF(('LSR2018'!$B:$B=""" & course_code & """)*('LSR2018'!$AA:$AA=""Confirmed""),'LSR2018'!$Y:$Y))"
    'Discard 0 values
    Dim i As Long
    For i = 1 To 8
        If loc.Offset(i, 0).Value2 <= 0 Then loc.Offset(i, 0) = "N/A"
    Next
End Sub


'Hardcoded: Range of earliest registration dates (R6:R12), cumulative unique learners (T132)
Sub tombstone_data(loc As Range, course_code As String)
    'Labels
    loc.Offset(-2, 0) = "A: Information"
    loc.Offset(0, 0) = "Duration / Durée (hrs)"
    loc.Offset(1, 0) = "Stream / Volet"
    loc.Offset(2, 0) = "Main Topic / Sujet principal"
    loc.Offset(3, 0) = "First Registration / Première inscription"
    loc.Offset(4, 0) = "Unique Learners Since Launch / Apps. uniques depuis lancement"
    'Formulæ
    loc.Offset(0, 3).Formula = "=VLOOKUP(""" & course_code & """,'Course Codes'!$C:$I,7,0)"
    loc.Offset(1, 3).Formula = "=VLOOKUP(""" & course_code & """,'Course Codes'!$C:$I,5,0)"
    loc.Offset(2, 3).Formula = "=VLOOKUP(""" & course_code & """,'Course Codes'!$C:$I,6,0)"
    loc.Offset(3, 3) = Application.Min(ThisWorkbook.Sheets("Report").Range("R6:R13").Value2)
    loc.Offset(4, 3).Formula = "=T132"
End Sub


'Hardcoded: None
Sub level_1_results(loc As Range, course_code As String)
    'Labels
    loc.Offset(-2, 0) = "B: Level 1 Results / Résultats de niveau 1"
    loc.Offset(0, 0) = "Question": loc.Offset(0, 2) = LAST_YEAR: loc.Offset(0, 3) = THIS_YEAR
    loc.Offset(1, 0) = "Overall Satisfaction / Satisfaction globale": loc.Offset(2, 0) = "Met My Learning Needs / Répondu à mes bes.": loc.Offset(3, 0) = "Knowledge Before / Connaissances avant": loc.Offset(4, 0) = "Knowledge After / Connaissances après"
    'Formulæ
    'Overall Satisfaction
    loc.Offset(1, 2).Formula = "=IFERROR(ROUND(AVERAGEIFS('L1SR" & Left(LAST_YEAR, 4) & "'!$V:$V,'L1SR" & Left(LAST_YEAR, 4) & "'!$B:$B,""" & course_code & """,'L1SR" & Left(LAST_YEAR, 4) & "'!$U:$U,""Overall Satisfaction""),2),""N/A"")"
    loc.Offset(1, 3).Formula = "=IFERROR(ROUND(AVERAGEIFS('L1SR" & Left(THIS_YEAR, 4) & "'!$V:$V,'L1SR" & Left(THIS_YEAR, 4) & "'!$B:$B,""" & course_code & """,'L1SR" & Left(THIS_YEAR, 4) & "'!$U:$U,""Overall Satisfaction""),2),""N/A"")"
    'Learning Needs Met
    loc.Offset(2, 2).Formula = "=IFERROR(ROUND(AVERAGEIFS('L1SR" & Left(LAST_YEAR, 4) & "'!$V:$V,'L1SR" & Left(LAST_YEAR, 4) & "'!$B:$B,""" & course_code & """,'L1SR" & Left(LAST_YEAR, 4) & "'!$U:$U,""Learning Needs Met""),2),""N/A"")"
    loc.Offset(2, 3).Formula = "=IFERROR(ROUND(AVERAGEIFS('L1SR" & Left(THIS_YEAR, 4) & "'!$V:$V,'L1SR" & Left(THIS_YEAR, 4) & "'!$B:$B,""" & course_code & """,'L1SR" & Left(THIS_YEAR, 4) & "'!$U:$U,""Learning Needs Met""),2),""N/A"")"
    'Knowledge Before
    loc.Offset(3, 2).Formula = "=IFERROR(ROUND(AVERAGEIFS('L1SR" & Left(LAST_YEAR, 4) & "'!$V:$V,'L1SR" & Left(LAST_YEAR, 4) & "'!$B:$B,""" & course_code & """,'L1SR" & Left(LAST_YEAR, 4) & "'!$U:$U,""Knowledge before""),2),""N/A"")"
    loc.Offset(3, 3).Formula = "=IFERROR(ROUND(AVERAGEIFS('L1SR" & Left(THIS_YEAR, 4) & "'!$V:$V,'L1SR" & Left(THIS_YEAR, 4) & "'!$B:$B,""" & course_code & """,'L1SR" & Left(THIS_YEAR, 4) & "'!$U:$U,""Knowledge before""),2),""N/A"")"
    'Knowledge After
    loc.Offset(4, 2).Formula = "=IFERROR(ROUND(AVERAGEIFS('L1SR" & Left(LAST_YEAR, 4) & "'!$V:$V,'L1SR" & Left(LAST_YEAR, 4) & "'!$B:$B,""" & course_code & """,'L1SR" & Left(LAST_YEAR, 4) & "'!$U:$U,""Knowledge after""),2),""N/A"")"
    loc.Offset(4, 3).Formula = "=IFERROR(ROUND(AVERAGEIFS('L1SR" & Left(THIS_YEAR, 4) & "'!$V:$V,'L1SR" & Left(THIS_YEAR, 4) & "'!$B:$B,""" & course_code & """,'L1SR" & Left(THIS_YEAR, 4) & "'!$U:$U,""Knowledge after""),2),""N/A"")"
End Sub


'Hardcoded: One row per LSR, cumulative unique learners (T132)
Sub repeat_rate(loc As Range, course_code As String)
    'Labels
    loc.Offset(-2, 0) = "C: Cancellations and Repeats / Annulations et répétitions"
    loc.Offset(1, 1) = THIS_YEAR
    loc.Offset(2, 0) = "Cancellations / Annulations"
    loc.Offset(4, 0) = "Avg. Online Products / Moy. produits en-ligne"
    loc.Offset(6, 1) = "Since Launch / Depuis lancement"
    loc.Offset(7, 0) = "Repeat Rate* / Taux de reprise*"
    loc.Offset(9, 0) = "Avg. Online Products / Moy. des en ligne"
    loc.Offset(11, 0) = "*Taking the course for the 2nd+ time"
    loc.Offset(12, 0) = "*Cours suivi pour une 2e fois ou +"
    'Averages
    loc.Offset(4, 1) = 0.036
    loc.Offset(9, 1) = 0.086
    'Formulæ
    loc.Offset(2, 1).Formula = "=IFERROR(COUNTIFS(THIS_YEAR!AA:AA,""Cancelled"")/(COUNTIFS(THIS_YEAR!AA:AA,""Cancelled"")+COUNTIFS(THIS_YEAR!AA:AA,""Confirmed"")),0)"
    loc.Offset(7, 1).Formula = "=IFERROR(((COUNTIFS(LSR2011!B:B,""" & course_code & """,LSR2011!AA:AA,""Confirmed"")+" & _
                                          "COUNTIFS(LSR2012!B:B,""" & course_code & """,LSR2012!AA:AA,""Confirmed"")+" & _
                                          "COUNTIFS(LSR2013!B:B,""" & course_code & """,LSR2013!AA:AA,""Confirmed"")+" & _
                                          "COUNTIFS(LSR2014!B:B,""" & course_code & """,LSR2014!AA:AA,""Confirmed"")+" & _
                                          "COUNTIFS(LSR2015!B:B,""" & course_code & """,LSR2015!AA:AA,""Confirmed"")+" & _
                                          "COUNTIFS(LSR2016!B:B,""" & course_code & """,LSR2016!AA:AA,""Confirmed"")+" & _
                                          "COUNTIFS(LSR2017!B:B,""" & course_code & """,LSR2017!AA:AA,""Confirmed"")+" & _
                                          "COUNTIFS(LSR2018!B:B,""" & course_code & """,LSR2018!AA:AA,""Confirmed""))-T132)/T132,0)"
End Sub


'Hardcoded: None
Sub top_5_departments(loc As Range)
    'Labels
    loc.Offset(0, 0) = "D: " & THIS_YEAR & ": Top 5 Departments / Top 5 des ministères"
    loc.Offset(1, 0) = "Name / Nom": loc.Offset(1, 3) = "Regs. / Inscr."
    'Formulæ
    Dim myQuery As String
    myQuery = "SELECT TOP 5 [Billing Dept Name], COUNT([Billing Dept Name]) FROM [THIS_YEAR$] WHERE [Reg Status] = 'Confirmed' GROUP BY [Billing Dept Name] ORDER BY 2 DESC"
    Call SQL(query:=myQuery, result_location:=loc.Offset(2, 0), header:=False, as_array:=False)
    'Clear cells below in case of SQL overflow
    loc.Offset(7, 0).Resize(5, 2).ClearContents
    'Move counts two columns right
    loc.Offset(2, 1).Resize(5, 1).Copy
    loc.Offset(2, 3).PasteSpecial xlPasteValues
    loc.Offset(2, 1).Resize(5, 1).ClearContents
    Application.CutCopyMode = False
End Sub


'Hardcoded: None
Sub top_5_classifications(loc As Range)
    'Labels
    loc.Offset(0, 0) = "E: " & THIS_YEAR & ": Top 5 Classifications / Top 5 des classifications"
    loc.Offset(1, 0) = "Name / Nom": loc.Offset(1, 3) = "Regs. / Inscr."
    'Formulæ
    Dim myQuery As String
    myQuery = "SELECT TOP 5 [Learner Classif], COUNT([Learner Classif]) FROM [THIS_YEAR$] WHERE [Reg Status] = 'Confirmed' GROUP BY [Learner Classif] ORDER BY 2 DESC"
    Call SQL(query:=myQuery, result_location:=loc.Offset(2, 0), header:=False, as_array:=False)
    'Clear cells below in case of SQL overflow
    loc.Offset(7, 0).Resize(5, 2).ClearContents
    'Move counts two columns right
    loc.Offset(2, 1).Resize(5, 1).Copy
    loc.Offset(2, 3).PasteSpecial xlPasteValues
    loc.Offset(2, 1).Resize(5, 1).ClearContents
    Application.CutCopyMode = False
End Sub


'Hardcoded: None
Sub registrations_per_month(loc As Range)
    'Labels
    loc.Offset(0, 0) = "Registrations per Month": loc.Offset(0, 1) = LAST_YEAR: loc.Offset(0, 2) = THIS_YEAR
    Dim month_list As Variant: month_list = Array("April", "May", "June", "July", "August", "September", "October", "November", "December", "January", "February", "March")
    loc.Offset(1, 0).Resize(UBound(month_list) + 1, 1) = Application.Transpose(month_list)
    'Formulæ
    Dim i As Long, myQuery As String
    For i = LBound(month_list) To UBound(month_list)
        loc.Offset(i + 1, 1).Formula = "=COUNTIFS(LAST_YEAR!K:K,""" & month_list(i) & """,LAST_YEAR!AA:AA,""Confirmed"")"
        loc.Offset(i + 1, 2).Formula = "=COUNTIFS(THIS_YEAR!K:K,""" & month_list(i) & """,THIS_YEAR!AA:AA,""Confirmed"")"
    Next
End Sub


'Hardcoded: None
Sub registrations_by_learner_region(loc As Range)
    'Labels
    loc.Offset(0, 0) = "Registrations by Learner Region": loc.Offset(0, 1) = LAST_YEAR: loc.Offset(0, 2) = THIS_YEAR
    Dim region_list As Variant: region_list = Array("Atlantic", "NCR", "Ontario", "Pacific", "Prairie", "Québec")
    loc.Offset(1, 0).Resize(UBound(region_list) + 1, 1) = Application.Transpose(region_list)
    'Formulæ
    Dim i As Long, myQuery As String
    For i = LBound(region_list) To UBound(region_list)
        loc.Offset(i + 1, 1).Formula = "=COUNTIFS(LAST_YEAR!AF:AF,""" & region_list(i) & """,LAST_YEAR!AA:AA,""Confirmed"")"
        loc.Offset(i + 1, 2).Formula = "=COUNTIFS(THIS_YEAR!AF:AF,""" & region_list(i) & """,THIS_YEAR!AA:AA,""Confirmed"")"
    Next
End Sub


'Hardcoded: tab_list
Sub cumulative_unique_learners(loc As Range, course_code As String)
    
    'Labels
    loc.Offset(0, 0) = "Cumulative Unique Learners": loc.Offset(0, 1) = "Count": loc.Offset(0, 2) = "Running Total"
    
    'Instantiate a dictionary to store Student IDs
    Dim myDict As Object: Set myDict = CreateObject("Scripting.Dictionary")
    
    'Instantiate variables
    Dim tab_list As Variant: tab_list = Array("LSR2011", "LSR2012", "LSR2013", "LSR2014", _
                                              "LSR2015", "LSR2016", "LSR2017", "LSR2018")
    Dim month_list As Variant: month_list = Array("April", "May", "June", "July", "August", _
                                                  "September", "October", "November", "December", _
                                                  "January", "February", "March")
    Dim key As Variant, item As Variant, DataSource As Worksheet, myArray As Variant
    Dim r As Long, i As Long, j As Long, k As Long
    Dim first_start_date As Long, first_reg_date As Long
    
    'Loop through each LSR sheet
    For i = LBound(tab_list) To UBound(tab_list)
        
        'Clear items in dictionary
        For Each key In myDict.Keys
            myDict(key) = ""
        Next key
        
        'Set DataSource tab and sort
        Set DataSource = ThisWorkbook.Sheets(tab_list(i))
        r = DataSource.Cells(Rows.Count, 1).End(xlUp).Row
        DataSource.Columns("A:AM").Sort key1:=DataSource.Range("I1"), order1:=xlAscending, header:=xlYes
        
        'Load Student ID, Month, Reg Status, and Course Code into array (much faster than serially referencing a worksheet)
        myArray = Array(DataSource.Range("AD2:AD" & r).Value2, _
                        DataSource.Range("K2:K" & r).Value2, _
                        DataSource.Range("AA2:AA" & r).Value2, _
                        DataSource.Range("B2:B" & r).Value2)
        
        'Load array into dictionary, storing only first occurrence of each Student ID
        For j = LBound(myArray(0)) To UBound(myArray(0))
            If Not myDict.Exists(myArray(0)(j, 1)) And myArray(2)(j, 1) = "Confirmed" And myArray(3)(j, 1) = course_code Then
                myDict(myArray(0)(j, 1)) = myArray(1)(j, 1)
            End If
        Next j
        
        'Store counts of each month in month_count array
        ReDim month_count(0 To UBound(month_list), 0 To 1) As Variant
        'Initialize month names and counts
        For j = LBound(month_list) To UBound(month_list)
            month_count(j, 0) = month_list(j)
            month_count(j, 1) = 0
        Next j
        'Counts
        'Loop over items in dictionary - faster as no need to re-hash
        For j = LBound(month_count) To UBound(month_count)
            For Each item In myDict.Items
                If month_count(j, 0) = item Then month_count(j, 1) = month_count(j, 1) + 1
            Next item
        Next j
        
        'Output month_count to worksheet
        loc.Offset(1 + (i * 12), 0).Resize(UBound(month_count) + 1, 2) = month_count
        
        'Add running total
        If i = 0 Then
            loc.Offset(1, 2) = loc.Offset(1, 1)
        Else
            loc.Offset(1 + (i * 12), 2) = loc.Offset(i * 12, 2) + loc.Offset(1 + (i * 12), 1)
        End If
        For j = LBound(month_list) + 1 To UBound(month_list)
            loc.Offset(j + (i * 12) + 1, 2) = loc.Offset(j + (i * 12), 2) + loc.Offset(j + (i * 12) + 1, 1)
        Next j
        
    Next i
    
    Set myDict = Nothing
End Sub


'Hardcoded: None
Sub registrations_per_language(loc As Range)
    'Labels
    loc.Offset(0, 0) = "Registrations per Language": loc.Offset(0, 1) = "English / Anglais": loc.Offset(0, 2) = "French / Français"
    loc.Offset(1, 0) = LAST_YEAR: loc.Offset(2, 0) = THIS_YEAR
    'Formulæ
    loc.Offset(1, 1).Formula = "=COUNTIFS(LAST_YEAR!S:S,""English/Anglais"",LAST_YEAR!AA:AA,""Confirmed"")"
    loc.Offset(1, 2).Formula = "=COUNTIFS(LAST_YEAR!S:S,""French/Français"",LAST_YEAR!AA:AA,""Confirmed"")"
    loc.Offset(2, 1).Formula = "=COUNTIFS(THIS_YEAR!S:S,""English/Anglais"",THIS_YEAR!AA:AA,""Confirmed"")"
    loc.Offset(2, 2).Formula = "=COUNTIFS(THIS_YEAR!S:S,""French/Français"",THIS_YEAR!AA:AA,""Confirmed"")"
End Sub
