Function WorksheetExists(WorksheetName As String, Optional wb As Workbook) As Boolean
    If wb Is Nothing Then Set wb = ThisWorkbook
    With wb
        On Error Resume Next
        WorksheetExists = (.Sheets(WorksheetName).Name = WorksheetName)
        On Error GoTo 0
    End With
End Function

Sub CALCPLANEDEXCEPTIONS()

    Dim mainWB As String
        mainWB = ActiveWorkbook.Name
    Dim sheetExc As String
        sheetExc = "Exceptions"
    Dim wb As Workbook
    Dim myPath As String
    Dim myFile As String
    Dim myExtension As String
    Dim FldrPicker As FileDialog

    'Optimize Macro Speed
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    'Check if worksheet exist and create
    If WorksheetExists(sheetExc) Then
        'Delete sheets Exceptions
        Application.DisplayAlerts = False
        Worksheets(sheetExc).Delete
        Application.DisplayAlerts = True
    End If
    Sheets.Add Type:=xlWorksheet, Before:=ActiveSheet
    ActiveSheet.Name = sheetExc

    
    Range("A1").Value = "Exceptions"

    'Retrieve Target Folder Path From User
    Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)

    With FldrPicker
        .Title = "Select A Target Folder"
        .AllowMultiSelect = False
        If .Show <> -1 Then GoTo NextCode
            myPath = .SelectedItems(1) & "\"
    End With
    
    'In Case of Cancel
NextCode:
    myPath = myPath
    If myPath = "" Then GoTo ResetSettings
    
ResetSettings:
    'Target File Extension (must include wildcard "*")
    myExtension = "*.csv*"

    'Target Path with Ending Extention
    myFile = Dir(myPath & myExtension)

    'Loop through each Excel file in folder
    Do While myFile <> ""
        'Set variable equal to opened workbook
        Set wb = Workbooks.Open(Filename:=myPath & myFile)
    
        'Ensure Workbook has opened before moving on to next line of code
        DoEvents
    
        'Copy ExceptionMessages from each "Key-file" to Exceptions sheet in the main File
        
        'Workbooks(wb.Name).ActiveSheet.Range("E2:E10").SpecialCells(xlCellTypeConstants).Copy
        LastRow = Workbooks(wb.Name).ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
        'MsgBox LastRow
        
        If LastRow > 2 Then
            Let Copyrange = "E2:E" & LastRow
            'MsgBox "Copyrange: " + Copyrange
            'If Err.Number = 1004 And Err.Description = "No cells were found." Then GoTo NextFile
            On Error Resume Next
            Workbooks(wb.Name).ActiveSheet.Range(Copyrange).SpecialCells(xlCellTypeConstants).Copy
            Workbooks(mainWB).Worksheets(sheetExc).Range("A" & Rows.Count).End(xlUp).Offset(1).PasteSpecial Paste:=xlPasteValues, SkipBlanks:=False
        Else
            Let Copyrange = "E" & LastRow
            'MsgBox "Copyrange: " + Copyrange
            Workbooks(wb.Name).ActiveSheet.Range(Copyrange).Copy
            Workbooks(mainWB).Worksheets(sheetExc).Range("A" & Rows.Count).End(xlUp).Offset(1).PasteSpecial Paste:=xlPasteValues, SkipBlanks:=False
        End If
        
NextFile:
        'Disable marching ants around copied range
        Application.CutCopyMode = False

        'Close Workbook
        wb.Close SaveChanges:=False
      
        'Ensure Workbook has closed before moving on to next line of code
        DoEvents

        'Get next file name
        myFile = Dir
    Loop
    
'Reset Macro Optimization Settings
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    '======Pivot Table=======
    'Dim sheetExc As String
    '    sheetExc = "Exceptions"
    Dim myPivot As String
        myPivot = "Pivot"
    Dim DSheet As Worksheet
    Set DSheet = Worksheets(sheetExc)
    Dim PCache As PivotCache
    Dim PTable As PivotTable
    Dim PRange As Range
    Dim PLastRow As Long
    Dim PLastCol As Long
    Dim PTableName As String
        PTableName = "ExceptionPivotTable"
    
    'Check if worksheet exist and create
    If WorksheetExists(myPivot) Then
        'Delete sheets Pivot
        Application.DisplayAlerts = False
        Worksheets(myPivot).Delete
        Application.DisplayAlerts = True
    End If
    'Create New Sheet
    Sheets.Add Type:=xlWorksheet, Before:=ActiveSheet
    ActiveSheet.Name = myPivot
    
    Dim PSheet As Worksheet
    Set PSheet = Worksheets(myPivot)
    
    'Define Pivot Data Range

    PLastRow = DSheet.Cells(Rows.Count, 1).End(xlUp).Row
    PLastCol = DSheet.Cells(1, Columns.Count).End(xlToLeft).Column
    Set PRange = DSheet.Cells(1, 1).Resize(PLastRow, PLastCol)
    
    'Define Pivot Cache
    'Set PCache = ActiveWorkbook.PivotCaches.Create _
    '(SourceType:=xlDatabase, SourceData:=PRange). _
    'CreatePivotTable(TableDestination:=PSheet.Cells(2, 2), _
    'TableName:=PTableName)
    
    Set PCache = ActiveWorkbook.PivotCaches.Create _
    (SourceType:=xlDatabase, SourceData:=PRange)
    
    'Insert Blank Pivot Table
    Set PTable = PCache.CreatePivotTable _
    (TableDestination:=PSheet.Cells(1, 1), TableName:=PTableName)
    
    'Insert Row Fields
    With ActiveSheet.PivotTables(PTableName).PivotFields("Exceptions")
    .Orientation = xlRowField
    .Position = 1
    End With
    
    'Insert Data Field
    With ActiveSheet.PivotTables(PTableName).PivotFields("Exceptions")
    .Orientation = xlDataField
    .Position = 1
    .Function = xlCount
    '.NumberFormat = "#,##0"
    .Name = "Count of Exceptions"
    End With
    
    'Format Pivot
    PSheet.PivotTables(PTableName).ShowTableStyleRowStripes = True
    PSheet.PivotTables(PTableName).TableStyle2 = "PivotStyleMedium9"

    Worksheets(myPivot).Range("A1").ColumnWidth = 130
    
    '======Calc Planed=======
    Dim planedExc As Variant
planedExc = Array( _
"Comments are already  updated for the Combo type 'DETAIL' hence this record will be deferred", _
"Comments are already  updated for the Combo type 'DISCREPA' hence this record will be deferred", _
"Cut off time reached. Hence marked as Exception", _
"Member Name contains 30 or more characters, Manual processing required", _
"Multiple Membership Records with Active Status found for same Purchaser and EU", _
"Multiple records with Active Status found for same Purchaser and EU", _
"No Active Membership Record Found For Purchaser", _
"No Match Record Identified", _
"No Matching Covg. Periods Available For Input Record", _
"NO Matching Membership Record Found For Purchaser", _
"Duplicate Record for the Primary Key", _
"Unable to set Status to RESUBMITTED since there is no RESUBMITTED in the drop down", _
"Unable to set Status to DISCREPANCY since there is no DISCREPANCY in the drop down", _
"The ZipCode fetched from the Membership Screen did not match with any of the ZipCodes present in the Excel", _
"SSN Value is not present in the Membership Summary Screen", _
"No records found for processing", _
"Billing Unit not found for input record", _
"Both Terminate & Reinstate records are not available", _
"Invalid Profile ID", _
"Invalid BU ID", _
"No Matching Sequence Record Found", _
"is out of scope for RPA")

LastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row - 1
'MsgBox LastRow
CurrentRow = 2
Let myRange = "A2:A" & LastRow
Let planedRange = "C2:C" & LastRow
Dim excRange As Range
Set excRange = Range("A2:A" & LastRow)
excRange.Select

'Dim i As Variant
For Each i In ActiveSheet.Range(myRange)
    CurrentCell = "C" & CurrentRow
    For Each j In planedExc
        'MsgBox "i = " + i
        'MsgBox "j = " + j
        If InStr(i, j) > 0 Then
            'MsgBox j + " = Planned"
            Range(CurrentCell) = "Planned"
            Range(CurrentCell).Offset(0, 1) = Range(CurrentCell).Offset(0, -1)
        End If
    Next j
    CurrentRow = CurrentRow + 1
Next i

Range("C" & LastRow).Offset(-(LastRow - 1), 1).Value = "TotalPlaned"
Range("C" & LastRow).Offset(-(LastRow - 1), 1).Font.Bold = True
Range("D" & (LastRow + 1)).Formula = "=SUM(D2:D" & LastRow & ")"
Range("D" & (LastRow + 1)).Font.Bold = True

'=== Put the Total value in D1 cell
LastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "D").End(xlUp).Row
ActiveSheet.Range("D1").Value = ActiveSheet.Range("D" & LastRow)


'=== Write data to the Database RPA_Reports

Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
Dim query As String
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset

Dim Daily_Total As Integer
Dim Daily_E2E As Integer
Dim Daily_Kickout_Planed As Integer
Dim Daily_Kickout_Unexpected As Integer

Dim ToDate_Total As Integer
Dim ToDate_E2E As Integer
Dim ToDate_Kickout_Planed As Integer
Dim ToDate_Kickout_Unexpected As Integer

Daily_Total = Worksheets("Dashboard").Range("C7").Value
Daily_E2E = Worksheets("Dashboard").Range("C9").Value
Daily_Kickout_Planed = Worksheets("Pivot").Range("D1").Value
Daily_Kickout_Unexpected = Daily_Kickout_Planed - Worksheets("Dashboard").Range("C14").Value

'---- Replace below highlighted names with the corresponding values

strCon = "Provider=SQLOLEDB; " & _
        "Data Source=csecpccad005.ecp.ca.kp.org; " & _
        "Initial Catalog=BPA_Reports;" & _
        "Integrated Security=SSPI"

'---  Open   the above connection string.

con.Open (strCon)

uSQL = "SELECT [TotalTransactionsAttempted],[AutomatedEndToEnd],[PlannedKickout],[UnexpectedKickout]" & _
      "FROM [BPA_Reports].[dbo].[BotM_Term]" & _
      "WHERE [Daily_ToDateReport] = 'ToDate';"

rs.Open uSQL, con

ToDate_Total = rs(0) + Daily_Total
ToDate_E2E = rs(1) + Daily_E2E
ToDate_Kickout_Planed = rs(2) + Daily_Kickout_Planed
ToDate_Kickout_Unexpected = ToDate_Total - ToDate_E2E - ToDate_Kickout_Planed

uSQL = "UPDATE BotM_Term " & _
      "Set [TotalTransactionsAttempted] = '" & Daily_Total & "' " & _
      ",[AutomatedEndToEnd] = '" & Daily_E2E & "' " & _
      ",[PlannedKickout] = '" & Daily_Kickout_Planed & "' " & _
      ",[UnexpectedKickout] = '" & Daily_Kickout_Unexpected & "' " & _
      "WHERE [Daily_ToDateReport] = 'Daily';"

con.Execute uSQL

uSQL = "UPDATE BotM_Term " & _
      "Set [TotalTransactionsAttempted] = '" & ToDate_Total & "' " & _
      ",[AutomatedEndToEnd] = '" & ToDate_E2E & "' " & _
      ",[PlannedKickout] = '" & ToDate_Kickout_Planed & "' " & _
      ",[UnexpectedKickout] = '" & ToDate_Kickout_Unexpected & "' " & _
      "WHERE [Daily_ToDateReport] = 'ToDate';"
         
con.Execute uSQL

con.Close

End Sub


