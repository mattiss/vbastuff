Attribute VB_Name = "WorksheetsToolbox"
Sub Test()
    ExportAsCSV ActiveSheet, "c:\temp\tada.csv", ";"
End Sub
Sub ExportAsCSV(wks As Worksheet, sFilePath As String, sDelimiter As String)
    Dim f As Integer
    f = FreeFile
    Open sFilePath For Output As f
    
    Dim textlines() As Variant, outputvar As Variant
    Set ExportRange = wks.Range("A1").CurrentRegion
    
    Set Lines = ExportRange.Rows
    For Each Line In Lines
        out = Join(Application.Transpose(Application.Transpose(Line)), sDelimiter)
        Print #f, out
    Next
    Close f
End Sub
Function AddNewSheet(sName As String, Optional wb As Workbook = Nothing) As Worksheet
    If (wb Is Nothing) Then Set wb = ActiveWorkbook
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Sheets(sName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    Set AddNewSheet = wb.Sheets.Add
    AddNewSheet.Name = sName
End Function
Function CreateFromCSV(sFilePath As String, sDelimiter As String, Optional wb As Workbook = Nothing) As Worksheet
    If (wb Is Nothing) Then Set wb = ActiveWorkbook
    Dim sName As String
    sName = GetFileNameFromFilePath(sFilePath)
    
    Set CreateFromCSV = AddNewSheet(sName, wb)
    
    Dim f As Integer
    f = FreeFile
    Open sFilePath For Input As #f
    
    Dim Line As String
    Set Rng = CreateFromCSV.Range("A1")
    While Not EOF(f)
        Line Input #f, Line ' read in data 1 line at a time
        Values = Split(Line, sDelimiter)
        Range(Rng, Rng.Offset(0, UBound(Values))) = Application.Transpose(Application.Transpose(Values))
        
        Set Rng = Rng.Offset(1)
    Wend
    Close f
End Function
