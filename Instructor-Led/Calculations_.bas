Attribute VB_Name = "Calculations_"
Option Explicit

'Hardcoded: One row per LSR, limit of For loop
Sub first_offering(loc As Range, course_code As String)
    '.FormulaArray is VBA equivalent of Ctrl+Shift+Enter (i.e. array) function
    'Note intersting use of (cond_1*cond_2), (cond_1+cond_2)>0 operators: used in place of AND, OR within array functions
    loc.Offset(0, 0) = "First Delivered/Open Offering in Each LSR"
    loc.Offset(1, 0).FormulaArray = "=MIN(IF(('LSR2011'!$B:$B=""" & course_code & """)*(('LSR2011'!$H:$H=""Delivered - Normal"")+('LSR2011'!$H:$H=""Open - Normal"")>0),'LSR2011'!$I:$I))"
    loc.Offset(2, 0).FormulaArray = "=MIN(IF(('LSR2012'!$B:$B=""" & course_code & """)*(('LSR2012'!$H:$H=""Delivered - Normal"")+('LSR2012'!$H:$H=""Open - Normal"")>0),'LSR2012'!$I:$I))"
    loc.Offset(3, 0).FormulaArray = "=MIN(IF(('LSR2013'!$B:$B=""" & course_code & """)*(('LSR2013'!$H:$H=""Delivered - Normal"")+('LSR2013'!$H:$H=""Open - Normal"")>0),'LSR2013'!$I:$I))"
    loc.Offset(4, 0).FormulaArray = "=MIN(IF(('LSR2014'!$B:$B=""" & course_code & """)*(('LSR2014'!$H:$H=""Delivered - Normal"")+('LSR2014'!$H:$H=""Open - Normal"")>0),'LSR2014'!$I:$I))"
    loc.Offset(5, 0).FormulaArray = "=MIN(IF(('LSR2015'!$B:$B=""" & course_code & """)*(('LSR2015'!$H:$H=""Delivered - Normal"")+('LSR2015'!$H:$H=""Open - Normal"")>0),'LSR2015'!$I:$I))"
    loc.Offset(6, 0).FormulaArray = "=MIN(IF(('LSR2016'!$B:$B=""" & course_code & """)*(('LSR2016'!$H:$H=""Delivered - Normal"")+('LSR2016'!$H:$H=""Open - Normal"")>0),'LSR2016'!$I:$I))"
    loc.Offset(7, 0).FormulaArray = "=MIN(IF(('LSR2017'!$B:$B=""" & course_code & """)*(('LSR2017'!$H:$H=""Delivered - Normal"")+('LSR2017'!$H:$H=""Open - Normal"")>0),'LSR2017'!$I:$I))"
    loc.Offset(8, 0).FormulaArray = "=MIN(IF(('LSR2018'!$B:$B=""" & course_code & """)*(('LSR2018'!$H:$H=""Delivered - Normal"")+('LSR2018'!$H:$H=""Open - Normal"")>0),'LSR2018'!$I:$I))"
    'Discard 0 values
    Dim i As Long
    For i = 1 To 8
        If loc.Offset(i, 0).Value2 <= 0 Then loc.Offset(i, 0) = "N/A"
    Next
End Sub


'Hardcoded: Range of earliest registration dates (V6:V13), cumulative unique learners (X132), denominator repeat rate (X132)
Sub tombstone_data(loc As Range, course_code As String)
    'Labels
    loc.Offset(-2, 0) = "A: Information"
    loc.Offset(0, 0) = "Duration (hours) / Durée (heures)"
    loc.Offset(1, 0) = "Stream / Volet"
    loc.Offset(2, 0) = "Main Topic / Sujet principal"
    loc.Offset(3, 0) = "First Offering / Première offre"
    loc.Offset(4, 0) = "Unique Learners Since Launch / Apps. uniques depuis lancement"
    loc.Offset(5, 0) = "Repeat Rate / Taux de reprise"
    'Formulæ
    loc.Offset(0, 3).Formula = "=IFNA(MODE.SNGL('THIS_YEAR'!$M:$M),MODE.SNGL('LAST_YEAR'!$M:$M))"
    loc.Offset(1, 3).Formula = "=VLOOKUP(""" & course_code & """,'Course Codes'!$C:$K,5,0)"
    loc.Offset(2, 3).Formula = "=VLOOKUP(""" & course_code & """,'Course Codes'!$C:$K,6,0)"
    loc.Offset(3, 3) = Application.Min(ThisWorkbook.Sheets("Report").Range("V6:V13").Value2)
    loc.Offset(4, 3).Formula = "=X132"
    loc.Offset(5, 3).Formula = "=((COUNTIFS(LSR2011!B:B,""" & course_code & """,LSR2011!AD:AD,""Confirmed"")+" & _
                                  "COUNTIFS(LSR2012!B:B,""" & course_code & """,LSR2012!AD:AD,""Confirmed"")+" & _
                                  "COUNTIFS(LSR2013!B:B,""" & course_code & """,LSR2013!AD:AD,""Confirmed"")+" & _
                                  "COUNTIFS(LSR2014!B:B,""" & course_code & """,LSR2014!AD:AD,""Confirmed"")+" & _
                                  "COUNTIFS(LSR2015!B:B,""" & course_code & """,LSR2015!AD:AD,""Confirmed"")+" & _
                                  "COUNTIFS(LSR2016!B:B,""" & course_code & """,LSR2016!AD:AD,""Confirmed"")+" & _
                                  "COUNTIFS(LSR2017!B:B,""" & course_code & """,LSR2017!AD:AD,""Confirmed"")+" & _
                                  "COUNTIFS(LSR2018!B:B,""" & course_code & """,LSR2018!AD:AD,""Confirmed""))-X132)/X132"
End Sub


'Hardcoded: All
Sub cheat_total(WS As Worksheet)
    'Labels
    WS.Range("F5") = "B: Totals / Totaux"
    WS.Range("G7") = LAST_YEAR: WS.Range("H7") = THIS_YEAR
    WS.Range("F8") = "Delivered Offerings / Offres livrées": WS.Range("F11") = "Cancelled Offs. / Offres annulées": WS.Range("F14") = "Registrations / Inscriptions": WS.Range("F17") = "No-Shows / Absences"
    'Formulæ
    'LAST_YEAR
    WS.Range("G8").Formula = "=AH6"
    WS.Range("G11").Formula = "=AG6"
    WS.Range("G14").Formula = "=AF37"
    WS.Range("G17").Formula = "=AH22"
    'THIS_YEAR
    WS.Range("H8").Formula = "=AH7"
    WS.Range("H11").Formula = "=AG7"
    WS.Range("H14").Formula = "=AF38"
    WS.Range("H17").Formula = "=AH23"
End Sub


'Hardcoded: None
Sub level_1_results(loc As Range, course_code As String)
    'Labels
    loc.Offset(-2, 0) = "C: Level 1 Results / Résultats de niveau 1"
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
Sub offerings_cancelled(loc As Range)
    'Labels
    loc.Offset(0, 1) = "Cancellation Rate": loc.Offset(0, 2) = "Cancelled Offerings": loc.Offset(0, 3) = "Open, Delivered Offerings"
    loc.Offset(1, 0) = LAST_YEAR: loc.Offset(2, 0) = THIS_YEAR
    'Formulæ
    'Cancelled offerings
    Dim myQuery As String
    myQuery = "SELECT COUNT(*) FROM (SELECT DISTINCT [Offering ID] FROM [LAST_YEAR$] WHERE [Offering Status] = 'Cancelled - Normal')"
    Call SQL(query:=myQuery, result_location:=loc.Offset(1, 2), header:=False, as_array:=False)
    myQuery = "SELECT COUNT(*) FROM (SELECT DISTINCT [Offering ID] FROM [THIS_YEAR$] WHERE [Offering Status] = 'Cancelled - Normal')"
    Call SQL(query:=myQuery, result_location:=loc.Offset(2, 2), header:=False, as_array:=False)
    'Open and Delivered Offerings
    myQuery = "SELECT COUNT(*) FROM (SELECT DISTINCT [Offering ID] FROM [LAST_YEAR$] WHERE [Offering Status] IN ('Delivered - Normal', 'Open - Normal'))"
    Call SQL(query:=myQuery, result_location:=loc.Offset(1, 3), header:=False, as_array:=False)
    myQuery = "SELECT COUNT(*) FROM (SELECT DISTINCT [Offering ID] FROM [THIS_YEAR$] WHERE [Offering Status] IN ('Delivered - Normal', 'Open - Normal'))"
    Call SQL(query:=myQuery, result_location:=loc.Offset(2, 3), header:=False, as_array:=False)
    'Cancellation Rate = Cancelled / (Delivered + Open + Cancelled)
    'LAST_YEAR
    Dim denominator As Long: denominator = loc.Offset(1, 2) + loc.Offset(1, 3)
    If denominator Then loc.Offset(1, 1) = loc.Offset(1, 2) / denominator Else loc.Offset(1, 1) = 0
    'THIS_YEAR
    denominator = loc.Offset(2, 2) + loc.Offset(2, 3)
    If denominator Then loc.Offset(2, 1) = loc.Offset(2, 2) / denominator Else loc.Offset(2, 1) = 0
End Sub


'Hardcoded: None
Sub offerings_per_month(loc As Range)
    'Labels
    loc.Offset(0, 0) = "Offerings per Month": loc.Offset(0, 1) = LAST_YEAR: loc.Offset(0, 2) = THIS_YEAR: loc.Offset(0, 3) = "Client Reqs. /" & vbCr & "Demandes des" & vbCr & "clients," & vbCr & THIS_YEAR
    Dim month_list As Variant: month_list = Array("April", "May", "June", "July", "August", "September", "October", "November", "December", "January", "February", "March")
    loc.Offset(1, 0).Resize(UBound(month_list) + 1, 1) = Application.Transpose(month_list)
    'Formulæ
    Dim i As Long, myQuery As String
    For i = LBound(month_list) To UBound(month_list)
        myQuery = "SELECT COUNT(*) FROM (SELECT DISTINCT [Offering ID] FROM [LAST_YEAR$] WHERE [Month] = '" & month_list(i) & "' AND [Offering Status] IN ('Delivered - Normal', 'Open - Normal'))"
        Call SQL(query:=myQuery, result_location:=loc.Offset(i + 1, 1), header:=False, as_array:=False)
        myQuery = "SELECT COUNT(*) FROM (SELECT DISTINCT [Offering ID] FROM [THIS_YEAR$] WHERE [Month] = '" & month_list(i) & "' AND [Offering Status] IN ('Delivered - Normal', 'Open - Normal'))"
        Call SQL(query:=myQuery, result_location:=loc.Offset(i + 1, 2), header:=False, as_array:=False)
        myQuery = "SELECT COUNT(*) FROM (SELECT DISTINCT [Offering ID] FROM [THIS_YEAR$] WHERE [Month] = '" & month_list(i) & "' AND [Client] IS NOT NULL)"
        Call SQL(query:=myQuery, result_location:=loc.Offset(i + 1, 3), header:=False, as_array:=False)
    Next
End Sub


'Hardcoded: None
Sub offerings_per_region(loc As Range)
    'Labels
    loc.Offset(0, 0) = "Offerings per Region, " & THIS_YEAR
    Dim region_list As Variant: region_list = Array("Atlantic", "NCR", "Ontario", "Pacific", "Prairie", "Québec")
    loc.Offset(1, 0).Resize(UBound(region_list) + 1, 1) = Application.Transpose(region_list)
    'Formulæ
    Dim i As Long, myQuery As String
    For i = LBound(region_list) To UBound(region_list)
        myQuery = "SELECT COUNT(*) FROM (SELECT DISTINCT [Offering ID] FROM [THIS_YEAR$] WHERE [Offering Region] = '" & region_list(i) & "' AND [Offering Status] IN ('Delivered - Normal', 'Open - Normal'))"
        Call SQL(query:=myQuery, result_location:=loc.Offset(i + 1, 1), header:=False, as_array:=False)
    Next
End Sub


'Hardcoded: None
Sub average_no_shows(loc As Range)
    'Labels
    loc.Offset(0, 1) = "Average No-Shows per Offering": loc.Offset(0, 2) = "Offerings": loc.Offset(0, 3) = "No-Shows"
    loc.Offset(1, 0) = LAST_YEAR: loc.Offset(2, 0) = THIS_YEAR
    'Formulæ
    'Offerings
    Dim myQuery As String
    myQuery = "SELECT COUNT(*) FROM (SELECT DISTINCT [Offering ID] FROM [LAST_YEAR$] WHERE [Offering Status] IN ('Delivered - Normal', 'Open - Normal'))"
    Call SQL(query:=myQuery, result_location:=loc.Offset(1, 2), header:=False, as_array:=False)
    myQuery = "SELECT COUNT(*) FROM (SELECT DISTINCT [Offering ID] FROM [THIS_YEAR$] WHERE [Offering Status] IN ('Delivered - Normal', 'Open - Normal'))"
    Call SQL(query:=myQuery, result_location:=loc.Offset(2, 2), header:=False, as_array:=False)
    'No-Shows
    loc.Offset(1, 3) = Application.Sum(ThisWorkbook.Sheets("LAST_YEAR").Columns("AF:AF"))
    loc.Offset(2, 3) = Application.Sum(ThisWorkbook.Sheets("THIS_YEAR").Columns("AF:AF"))
    'Average No-Shows per Offering
    'LAST_YEAR
    Dim denominator As Long: denominator = loc.Offset(1, 2)
    If denominator Then loc.Offset(1, 1) = loc.Offset(1, 3) / denominator Else loc.Offset(1, 1) = 0
    'THIS_YEAR
    denominator = loc.Offset(2, 2)
    If denominator Then loc.Offset(2, 1) = loc.Offset(2, 3) / denominator Else loc.Offset(2, 1) = 0
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
        
        'Set DataSource tab and sort by Start Date
        Set DataSource = ThisWorkbook.Sheets(tab_list(i))
        r = DataSource.Cells(Rows.Count, 1).End(xlUp).Row
        DataSource.Columns("A:AP").Sort key1:=DataSource.Range("I1"), order1:=xlAscending, header:=xlYes
        
        'Load Student ID, Month, Reg Status, and Course Code into array (much faster than serially referencing a worksheet)
        myArray = Array(DataSource.Range("AG2:AG" & r).Value2, _
                        DataSource.Range("K2:K" & r).Value2, _
                        DataSource.Range("AD2:AD" & r).Value2, _
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
Sub offerings_per_language(loc As Range)
    'Labels
    loc.Offset(0, 0) = "Offerings per Language": loc.Offset(0, 1) = "English / Anglais": loc.Offset(0, 2) = "French / Français"
    loc.Offset(1, 0) = LAST_YEAR: loc.Offset(2, 0) = THIS_YEAR
    'Formulæ
    Dim myQuery As String
    myQuery = "SELECT COUNT(*) FROM (SELECT DISTINCT [Offering ID] FROM [LAST_YEAR$] WHERE [Offering Status] IN ('Delivered - Normal', 'Open - Normal') AND [Offering Language] = 'English/Anglais')"
    Call SQL(query:=myQuery, result_location:=loc.Offset(1, 1), header:=False, as_array:=False)
    myQuery = "SELECT COUNT(*) FROM (SELECT DISTINCT [Offering ID] FROM [LAST_YEAR$] WHERE [Offering Status] IN ('Delivered - Normal', 'Open - Normal') AND [Offering Language] = 'French/Français')"
    Call SQL(query:=myQuery, result_location:=loc.Offset(1, 2), header:=False, as_array:=False)
    myQuery = "SELECT COUNT(*) FROM (SELECT DISTINCT [Offering ID] FROM [THIS_YEAR$] WHERE [Offering Status] IN ('Delivered - Normal', 'Open - Normal') AND [Offering Language] = 'English/Anglais')"
    Call SQL(query:=myQuery, result_location:=loc.Offset(2, 1), header:=False, as_array:=False)
    myQuery = "SELECT COUNT(*) FROM (SELECT DISTINCT [Offering ID] FROM [THIS_YEAR$] WHERE [Offering Status] IN ('Delivered - Normal', 'Open - Normal') AND [Offering Language] = 'French/Français')"
    Call SQL(query:=myQuery, result_location:=loc.Offset(2, 2), header:=False, as_array:=False)
End Sub


'Hardcoded: None
Sub average_class_size(loc As Range)
    'Labels
    loc.Offset(0, 0) = "Average Class Size": loc.Offset(0, 1) = "Regs": loc.Offset(0, 2) = "Offerings": loc.Offset(0, 3) = "Regs per Offering"
    loc.Offset(1, 0) = LAST_YEAR: loc.Offset(2, 0) = THIS_YEAR
    'Formulæ
    'Regs
    Dim myQuery As String
    myQuery = "SELECT COUNT([Reg #]) FROM [LAST_YEAR$] WHERE [Reg Status] = 'Confirmed'"
    Call SQL(query:=myQuery, result_location:=loc.Offset(1, 1), header:=False, as_array:=False)
    myQuery = "SELECT COUNT([Reg #]) FROM [THIS_YEAR$] WHERE [Reg Status] = 'Confirmed'"
    Call SQL(query:=myQuery, result_location:=loc.Offset(2, 1), header:=False, as_array:=False)
    'Offerings
    myQuery = "SELECT COUNT(*) FROM (SELECT DISTINCT [Offering ID] FROM [LAST_YEAR$] WHERE [Offering Status] IN ('Delivered - Normal', 'Open - Normal'))"
    Call SQL(query:=myQuery, result_location:=loc.Offset(1, 2), header:=False, as_array:=False)
    myQuery = "SELECT COUNT(*) FROM (SELECT DISTINCT [Offering ID] FROM [THIS_YEAR$] WHERE [Offering Status] IN ('Delivered - Normal', 'Open - Normal'))"
    Call SQL(query:=myQuery, result_location:=loc.Offset(2, 2), header:=False, as_array:=False)
    'Average Class Size
    'LAST_YEAR
    Dim denominator As Long: denominator = loc.Offset(1, 2)
    If denominator Then loc.Offset(1, 3) = loc.Offset(1, 1) / denominator Else loc.Offset(1, 3) = 0
    'THIS_YEAR
    denominator = loc.Offset(2, 2)
    If denominator Then loc.Offset(2, 3) = loc.Offset(2, 1) / denominator Else loc.Offset(2, 3) = 0
End Sub
