Sub makeStatusReport()
    Dim wb As Workbook
    Dim wsExport As Worksheet, wsData As Worksheet
    Dim tblExport As ListObject
    Dim maxExportRow As Long, maxExportCol As Long
    Dim maxExportRowFil As Long
    Dim cellError As Range
    Dim lastRowData As Long
    Dim visibleRange As Range
    Dim i As Long, j As Long, k As Long
    
    Set wb = ActiveWorkbook
    
    Set wsExport = ActiveSheet
    Set wsData = wb.Sheets("Data")
    
    Set tblExport = wsExport.ListObjects(1)
    
    maxExportRow = tblExport.ListRows.Count
    maxExportCol = tblExport.ListColumns.Count
    
    Application.ScreenUpdating = False
    
    'Remove Active Filters from Export
    tblExport.AutoFilter.ShowAllData
    
    'Remove empty rows from the end of Export Table
    For i = maxExportRow To 1 Step -1
        Set cellError = tblExport.DataBodyRange.Cells(i, 1)
        If IsError(cellError.Value) Then
            tblExport.ListRows(i).Delete
        ElseIf Trim(cellError.Value) = "" Then
            tblExport.ListRows(i).Delete
        Else
            Exit For
        End If
    Next i
    
    'Clear Data range
    lastRowData = wsData.Cells.SpecialCells(xlCellTypeLastCell).Row
    With wsData
        .Range("A" & 2 & ":" & Split(.Cells(1, 9).Address, "$")(1) & lastRowData).ClearContents
    End With
    
    'Copy Export table onto Data range
    k = 0 'Offset index
    tblExport.Range.AutoFilter Field:=3, Criteria1:="<>BEI LLC SB"  'Filter out Seal Beach
    For i = 8 To maxExportCol
        tblExport.Range.AutoFilter Field:=i, Criteria1:="<>0"   'Filter out 0-value per SS
        On Error Resume Next
        maxExportRowFil = tblExport.DataBodyRange.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count
        On Error GoTo 0
        'In Data Sheet:
        'Column A
        wsData.Range("A2:A" & maxExportRowFil + 1).Offset(k, 0).Value = tblExport.ListColumns(i).Name
        'Column B:H
        For j = 2 To 8
            On Error Resume Next
            Set visibleRange = tblExport.ListColumns(j - 1).DataBodyRange.SpecialCells(xlCellTypeVisible)
            On Error GoTo 0
            If Not visibleRange Is Nothing Then
                visibleRange.Copy _
                Destination:=wsData.Cells(2 + k, j)
            End If
        Next j
        'Column I
        tblExport.ListColumns(i).DataBodyRange.Copy _
            Destination:=wsData.Range("I2").Offset(k, 0)
        'Prepare for next iteration
        k = k + maxExportRowFil
        tblExport.Range.AutoFilter Field:=i
    Next i
    tblExport.Range.AutoFilter Field:=3
    
    Application.ScreenUpdating = True
    
    MsgBox "Task Complete: Status Report Created."
    
End Sub
