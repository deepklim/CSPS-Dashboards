Attribute VB_Name = "PDF_"
Option Explicit

Sub PDF(WS As Worksheet, destination_folder As String, filename As String)
    
    WS.ExportAsFixedFormat Type:=xlTypePDF, _
        filename:=destination_folder & "\" & filename, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
    
End Sub
