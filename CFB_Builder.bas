Attribute VB_Name = "CFB_Builder"
Option Explicit
'---------------------------------------------------------------------------------------
' Module    : CFBFormBuilder
' Author    : Janakiraman Murugesan
' Purpose   : Generate feedback forms from input workbook data
' Version History:
'   1.0.0 - 08-March-2025 - Initial version created
'---------------------------------------------------------------------------------------
Sub CFBFormBuilder()
    On Error GoTo ErrorHandler
    
    Dim fieldValues As Variant
    Dim ws As Worksheet
    Dim mismatches As Collection
    Dim inputWb As Workbook
    Dim templateWb As Workbook
    Dim dataArray As Variant
    Dim i As Long
    Dim requestedDate As String
    Dim folderPath As String
    
    ' Initialize variables
    fieldValues = Array("Sr. No", "SOW No", "SOW Description", "Cyient-Team Member's Name", _
                       "Cyient Team Lead Name", "WEC Manager Details")
    
    ' Performance optimizations
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' Get input workbook and data
    Set inputWb = InputWorkbooK
    If inputWb Is Nothing Then GoTo CleanUp
    
    Set ws = inputWb.Sheets(1)
    Set mismatches = CheckFieldsInInputsheet(ws, fieldValues)
    
    ' Check for missing fields
    If mismatches.Count > 0 Then
        Dim mismatchMsg As String: mismatchMsg = "The following table titles were not found in First Row:" & vbNewLine
        Dim mismatch As Variant
        For Each mismatch In mismatches
            mismatchMsg = mismatchMsg & "- " & mismatch & vbNewLine
        Next mismatch
        MsgBox mismatchMsg, vbExclamation
        GoTo CleanUp
    End If
    
    ' Create output folder
    folderPath = inputWb.Path & "\OutputForms"
    If Len(Dir(folderPath, vbDirectory)) = 0 Then MkDir folderPath
    
    ' Get all data at once
    Dim lastRow As Long: lastRow = ws.UsedRange.Rows.Count
    dataArray = ws.Range("A2:" & ws.Cells(2, ws.Columns.Count).End(xlToLeft).Address).Resize(lastRow - 1).Value
    
    ' Cache column indices
    Dim colSowNo As Long: colSowNo = GetColumnIndex("SOW No", ws)
    Dim colWecMgr As Long: colWecMgr = GetColumnIndex("WEC Manager Details", ws)
    Dim colSowDesc As Long: colSowDesc = GetColumnIndex("SOW Description", ws)
    Dim colTeamMem As Long: colTeamMem = GetColumnIndex("Cyient-Team Member's Name", ws)
    Dim colTeamLead As Long: colTeamLead = GetColumnIndex("Cyient Team Lead Name", ws)
    Dim colSrNo As Long: colSrNo = GetColumnIndex("Sr. No", ws)
    
    ' Create single template copy
    ThisWorkbook.Sheets(Array("Covering letter", "Feedback Form")).Copy
    Set templateWb = ActiveWorkbook
    requestedDate = Format(Now, "MM-DD-YYYY")
    
    ' Process all records using array
    For i = 1 To UBound(dataArray, 1)
        With templateWb.Worksheets("Feedback Form")
            .Cells(4, 4).Value = dataArray(i, colSowNo)
            .Cells(5, 4).Value = dataArray(i, colWecMgr)
            .Cells(6, 4).Value = dataArray(i, colSowDesc)
            .Cells(7, 4).Value = dataArray(i, colTeamMem)
            .Cells(6, 17).Value = dataArray(i, colTeamLead)
            .Cells(8, 17).Value = requestedDate
        End With
        
        ' Save with unique filename
        Dim outFilename As String
        outFilename = dataArray(i, colSrNo) & "_" & dataArray(i, colSowNo) & ".xlsx"
        templateWb.SaveAs folderPath & "\" & outFilename, FileFormat:=xlOpenXMLWorkbook
    Next i
    MsgBox "Files are Generated in : " & vbCrLf & folderPath
CleanUp:
    If Not templateWb Is Nothing Then templateWb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Set inputWb = Nothing
    Set ws = Nothing
    Set templateWb = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    GoTo CleanUp
End Sub

Function SelectedPath() As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select a CFB Input File"
        .Filters.Add "Excel Files", "*.xls;*.xlsx;*.xlsm"
        .AllowMultiSelect = False
        If .Show = -1 Then SelectedPath = .SelectedItems(1)
    End With
End Function

Function InputWorkbooK() As Workbook
    Dim filepath As String: filepath = SelectedPath
    If Len(filepath) = 0 Then Exit Function
    
    Dim wb As Workbook
    Dim filename As String: filename = Mid(filepath, InStrRev(filepath, "\") + 1)
    
    On Error Resume Next
    Set wb = Workbooks(filename)
    On Error GoTo 0
    
    If wb Is Nothing Then Set wb = Workbooks.Open(filepath)
    Set InputWorkbooK = wb
End Function

Public Function CheckFieldsInInputsheet(ws As Worksheet, fieldValues As Variant) As Collection
    Dim mismatch As New Collection
    If ws Is Nothing Or IsEmpty(fieldValues) Then Exit Function
    
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim headers As Variant: headers = ws.Range("A1", ws.Cells(1, lastCol)).Value
    Dim i As Long, j As Long
    Dim found As Boolean
    
    For i = LBound(fieldValues) To UBound(fieldValues)
        found = False
        For j = 1 To lastCol
            If Trim(LCase(headers(1, j))) = Trim(LCase(fieldValues(i))) Then
                found = True
                Exit For
            End If
        Next j
        If Not found Then mismatch.Add fieldValues(i)
    Next i
    
    Set CheckFieldsInInputsheet = mismatch
End Function

Function GetColumnIndex(searchValue As String, Optional ws As Worksheet) As Long
    If ws Is Nothing Then Set ws = ActiveSheet
    
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim i As Long
    searchValue = Trim(LCase(searchValue))
    
    For i = 1 To lastCol
        If Trim(LCase(ws.Cells(1, i).Value)) = searchValue Then
            GetColumnIndex = i
            Exit Function
        End If
    Next i
    GetColumnIndex = -1
End Function
