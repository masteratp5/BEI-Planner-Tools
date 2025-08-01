Sub makeShortageReport()
    Dim wb As Workbook
    Dim wsExport As Worksheet, wsNewS As Worksheet, wsOldS As Worksheet
    Dim tblExport As ListObject, tblNewS As ListObject, tblOldS As ListObject
    Dim maxExportCol As Long, maxNewSCol As Long, maxOldSCol As Long
    Dim maxNewSRow As Long
    Dim i As Long, j As Long
    Dim resizeRange As Range
    Dim maxMax As Long
    Dim cellError As Range
    Dim lookupCol As Range, lookupRange As Range, returnRange As Range, cellResult As Range, newCellCol As Range
    Dim colIndex As Long
    
    Set wb = ActiveWorkbook
    
    Set wsExport = wb.Sheets("Export")
    Set wsNewS = wb.Sheets("Shortages New")
    
    Set tblExport = wsExport.ListObjects("ExportTable")
    Set tblNewS = wsNewS.ListObjects("ShortagesNewTable")
    
    maxExportCol = tblExport.ListColumns.Count
    maxExportRow = tblExport.ListRows.Count
    maxNewSCol = tblNewS.ListColumns.Count
    maxNewSRow = tblNewS.ListRows.Count
    
    Application.ScreenUpdating = False
    
    'Delete Shortages Old
    Application.DisplayAlerts = False
    Sheets("Shortages Old").Delete
    Application.DisplayAlerts = True
    
    'Copy Shortages New Onto Shortages Old
    wsNewS.Copy Before:=wsNewS
    Sheets("Shortages New (2)").Name = "Shortages Old"
    
    Set wsOldS = wb.Sheets("Shortages Old")
    Set tblOldS = wsOldS.ListObjects(1)
    tblOldS.Name = "ShortagesOldTable"
    
    
    'Clear existing entries in New Shortage
    For i = 1 To maxNewSCol - 1
        If InStr(tblNewS.ListColumns(i).Name, "MS_") = 0 Then
            If InStr(tblNewS.ListColumns(i).Name, "CD_") = 0 Then
                If InStr(tblNewS.ListColumns(i).Name, "TS_") = 0 Then
                    tblNewS.ListColumns(i).DataBodyRange.ClearContents
                End If
            End If
        End If
    Next i
    
    'Resize Shortages New Table to accomadate Export if Shortages New Table is too small
    If maxExportRow > maxNewSRow Then
        Set resizeRange = tblNewS.Range.Resize(maxExportRow + 1, maxNewSCol)
        tblNewS.Resize resizeRange
    End If
    
    'Copy Export onto Shortages New
    For i = 1 To maxExportCol
        For j = 1 To maxNewSCol
            If tblExport.ListColumns(i).Name = tblNewS.ListColumns(j).Name Then
                tblNewS.DataBodyRange.Columns(j).Value = tblExport.DataBodyRange.Columns(i).Value
                Exit For
            End If
        Next j
    Next i
    
    'Delete empty rows in Shortages Old Table
    If maxNewSRow > maxOldSRow Then
        maxMax = maxNewSRow
    Else
        maxMax = maxOldSRow
    End If
    
    For i = maxMax To 1 Step -1
        Set cellError = tblOldS.DataBodyRange.Cells(i, 1)
        If IsError(cellError.Value) Then
            tblOldS.ListRows(i).Delete
        Else
            Exit For
        End If
    Next i
    
    'Delete empty rows in Shortages New Table
    If maxExportRow > maxNewSRow Then
        maxMax = maxExportRow
    Else
        maxMax = maxNewSRow
    End If
    
    For i = maxMax To 1 Step -1
        Set cellError = tblNewS.DataBodyRange.Cells(i, 1)
        If IsError(cellError.Value) Then
            tblNewS.ListRows(i).Delete
        Else
            Exit For
        End If
    Next i
        
    
    'Match Owner from Shortages Old to Shortages New
    Set lookupCol = tblNewS.ListColumns("Supply Order").DataBodyRange
    Set lookupRange = tblOldS.ListColumns("Supply Order").DataBodyRange
    Set returnRange = tblOldS.ListColumns("Owner").DataBodyRange
    Set newCellCol = tblNewS.ListColumns("Owner").DataBodyRange
    For Each cellResult In lookupCol
        newCellCol.Cells(cellResult.Row - tblNewS.Range.Row, 1).Value = Application.WorksheetFunction.XLookup(cellResult.Value, lookupRange, returnRange, "")
    Next cellResult
    
    'Match Comments from Shortages Old to Shortages New
    Set lookupCol = tblNewS.ListColumns("StockCode").DataBodyRange
    Set lookupRange = tblOldS.ListColumns("StockCode").DataBodyRange
    Set returnRange = tblOldS.ListColumns("Comments").DataBodyRange
    Set newCellCol = tblNewS.ListColumns("Comments").DataBodyRange
    For Each cellResult In lookupCol
        newCellCol.Cells(cellResult.Row - tblNewS.Range.Row, 1).Value = Application.WorksheetFunction.XLookup(cellResult.Value, lookupRange, returnRange, "")
    Next cellResult
    
    'Check past-due dates and fix them
    'Copy Latest Status onto Previous Status
    tblNewS.ListColumns("Supply Date Previous Status").DataBodyRange.Value = _
        tblNewS.ListColumns("Supply Date Latest Status").DataBodyRange.Value
    'Change BFE dates to TBD
    tblNewS.Range.AutoFilter Field:=3, Criteria1:="BFE"
    colIndex = tblNewS.ListColumns("Supply Date Latest Status").Index
    tblNewS.Range.AutoFilter Field:=colIndex, Criteria1:="<>TBD"
    On Error Resume Next 'Edge case - 0 cells after filter
    tblNewS.ListColumns("Supply Date Latest Status").DataBodyRange.SpecialCells(xlCellTypeVisible) = "TBD"
    tblNewS.ListColumns("Date Confirmed?" & vbLf & "(Y/N)").DataBodyRange.SpecialCells(xlCellTypeVisible) = "N"
    On Error GoTo 0
    tblNewS.Range.AutoFilter Field:=colIndex
    tblNewS.Range.AutoFilter Field:=3
    'Clear Previous Status from on-time dates
    colIndex = tblNewS.ListColumns("Check").Index
    tblNewS.Range.AutoFilter Field:=colIndex, Criteria1:=""
    tblNewS.ListColumns("Supply Date Previous Status").DataBodyRange.SpecialCells(xlCellTypeVisible) = ""
    tblNewS.Range.AutoFilter Field:=colIndex
    'Extract dates from Gebo?
    
    Application.ScreenUpdating = True
    
    MsgBox "Task Complete: Shortage Report Created."
    
End Sub
