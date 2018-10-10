Attribute VB_Name = "SQL_"
Option Explicit

Function SQL(ByVal query As String, Optional ByVal result_location As Range, _
             Optional ByVal header As Boolean = False, Optional ByVal as_array = False) As Variant
    
    If query = "" Then Exit Function
    
    Dim cn As Object, rs As Object
    'Data tabs are within this workbook, so connects to itself
    Set cn = CreateObject("ADODB.Connection")
    With cn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .ConnectionString = "Data Source=" & ThisWorkbook.path & "\" & ThisWorkbook.Name & ";" & _
        "Extended Properties=""Excel 12.0 Xml;HDR=YES"";"
        .Open
    End With
    
    'Build and execute the SQL query
    Set rs = cn.Execute(query)
    If rs.EOF Then GoTo ExitFunction
    
    'Returning results without header
    If Not header And as_array Then
        'Return records
        SQL = rs.GetRows
    ElseIf Not header And Not as_array Then
        'Print records
        result_location.CopyFromRecordset rs
    
    'Returning results with header
    ElseIf header And as_array Then
        'Cannot return header within an array, so raise error
        rs.Close
        cn.Close
        Set rs = Nothing
        Set cn = Nothing
        Err.Raise Number:=vbObjectError + 513, Description:="Cannot return header within an array"
    ElseIf header And Not as_array Then
        'Print header
        Dim i As Long
        For i = 0 To rs.Fields.Count - 1
            result_location.Offset(0, i) = rs.Fields(i).Name
        Next i
        'Print records
        result_location.Offset(1, 0).CopyFromRecordset rs
    End If
    
ExitFunction:
    'Close the connections
    rs.Close
    cn.Close
    Set rs = Nothing
    Set cn = Nothing
    
End Function
