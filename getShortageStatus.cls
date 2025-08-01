Sub getShortageStatus()
    Dim wbExport As Workbook, wbGebo As Workbook
    Dim geboPath As String
    Dim wsNewS As Worksheet, wsGeboP As Worksheet
    Dim tblNewS As LisObject
    Dim commentCol As ListColumn
    Dim maxNewSRow As Long
    Dim cell As Range, commentColRange As Range
    
    geboPath = _
        "C:\Users\cf875f\Documents\Shortage Reports\Expeditor SR Multi-Level Explorer (BEI-HB) v1.003 - COPY.xlsm"
    
    Set wbExport = ActiveWorkbook
    Set wbGebo = Workbooks.Open(geboPath)
    
    Set wsNewS = wbExport.Sheets("Shortages New")
    Set wsGeboP = wbGebo.Sheets("Pivot")
    
    Set tblNewS = wsNewS.ListObjects("ShortagesNewTable")
    
    Set commentCol = tblNewS.ListColumns("Comments")
    
    Set commentColRange = commentCol.DataBodyRange
    
    maxNewSRow = tblNewS.ListRows.Count
    
    Applicaton.ScreenUpdating = False
    
    'Fill every cell in "Comments" with Gebo Pivot
    For Each cell In commentColRange
        cell.Value = Application.WorksheetFunction.XLookup()
    Next cell
    
    Application.ScreenUpdating = True
    
    
    MsgBox "Task Complete: Shortages' statuses updated."
End Sub
