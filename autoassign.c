Sub autoassign()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastrow As Long
    Dim outRow As Long
    Dim i As Long, colCom As Long, colAssign As Long
    Dim summary As Worksheet
    
    Dim Arr As Variant                  ' Collector IDs (0-based)
    Dim quotas() As Long                ' Max assignments (0-based)
    Dim used() As Long                  ' Track usage (0-based)
    Dim data As Variant                 ' Entire sheet in memory
    
    Dim sortedIndex() As Long
    Dim tempIdx As Long
    Dim r As Long, c As Long
    Dim pick As Long
    Dim assigned As Boolean
    Dim attempts As Long
    Dim collDict As Object
    Dim collector As String
    Dim ckey As Variant
    
    Set wb = ActiveWorkbook
    Set ws = ActiveSheet
    
    colCom = Application.Match("Commission", ws.Rows(1), 0)
    If IsError(colCom) Then MsgBox "Commission column not found", vbCritical: Exit Sub
    
    colAssign = Application.Match("Collector ID", ws.Rows(1), 0)
    If IsError(colAssign) Then
        MsgBox "Collector ID column not found", vbCritical
        Exit Sub
    End If

    lastrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Load data to memory
    data = ws.Range("A1:AJ" & lastrow).Value
    
    ' Collector IDs (0-based)
    Arr = Array("SVS", "TSG", "DDR", "SL1", "SSS", "EB1", "SVW", "TTD", _
                "RW2", "SB5", "SKR", "SMK", "WB1", "NB2", "CJW", "ANT")
    
    ' Make quotas/used 0-based to match Arr
    ReDim quotas(0 To UBound(Arr))
    ReDim used(0 To UBound(Arr))
    
    ' Assign quotas (0-based)
    quotas(0) = 500
    quotas(1) = 409
    quotas(2) = 409
    quotas(3) = 409
    quotas(4) = 408
    quotas(5) = 250
    quotas(6) = 409
    quotas(7) = 359
    quotas(8) = 359
    quotas(9) = 359
    quotas(10) = 400
    quotas(11) = 400
    quotas(12) = 200
    quotas(13) = 200
    quotas(14) = 200
    quotas(15) = 250
    
    ' Prepare sorted index (holds worksheet row numbers)
    ReDim sortedIndex(2 To lastrow)
    For i = 2 To lastrow
        sortedIndex(i) = i
    Next i
    
    ' Sort by commission descending (simple bubble; replace with quicksort for huge datasets)
    For r = 2 To lastrow - 1
        For c = r + 1 To lastrow
            If val(data(sortedIndex(c), colCom)) > val(data(sortedIndex(r), colCom)) Then
                tempIdx = sortedIndex(r)
                sortedIndex(r) = sortedIndex(c)
                sortedIndex(c) = tempIdx
            End If
        Next c
    Next r
    
    Randomize
    
    ' Assign collectors - highest balances first
    For r = 2 To lastrow
        Dim rowIndex As Long
        rowIndex = sortedIndex(r)
        
        assigned = False
        attempts = 0
        
        Do While Not assigned And attempts < 500
            pick = Int((UBound(Arr) + 1) * Rnd)   ' 0..UBound(Arr)
            
            If used(pick) < quotas(pick) Then
                data(rowIndex, colAssign) = Arr(pick)
                used(pick) = used(pick) + 1
                assigned = True
            End If
            
            attempts = attempts + 1
        Loop
        
        If Not assigned Then
            data(rowIndex, colAssign) = "UNASSIGNED"
        End If
    Next r
    
    ' Write back
    ws.Range("A1:AJ" & lastrow).Value = data
    
    ' Remove old Summary first
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Sheets("Summary").delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Create new summary
    Set wsOut = wb.Worksheets.Add
    wsOut.Name = "Summary"
    
    ' Summarize totals
    Set collDict = CreateObject("Scripting.Dictionary")
    Dim amount As Double
    
    For i = 2 To UBound(data, 1)   ' skip header
        collector = Trim(data(i, colAssign))
        If IsNumeric(data(i, colCom)) Then amount = CDbl(data(i, colCom)) Else amount = 0
        If collector <> "" Then
            If collDict.exists(collector) Then
                collDict(collector) = collDict(collector) + amount
            Else
                collDict.Add collector, amount
            End If
        End If
    Next i
    
    ' Output dictionary to array
    ReDim outArr(1 To collDict.Count + 1, 1 To 2)
    outArr(1, 1) = "Collector"
    outArr(1, 2) = "Total"
    
    outRow = 2
    For Each ckey In collDict.Keys
        outArr(outRow, 1) = ckey
        outArr(outRow, 2) = collDict(ckey)
        outRow = outRow + 1
    Next ckey
    
    ' Write to sheet
    wsOut.Range("A1").Resize(collDict.Count + 1, 2).Value = outArr
    
    ' Sort by total descending
    With wsOut.Sort
        .SortFields.Clear
        .SortFields.Add Key:=wsOut.Range("B2:B" & collDict.Count + 1), Order:=xlDescending
        .SetRange wsOut.Range("A1:B" & collDict.Count + 1)
        .Header = xlYes
        .Apply
    End With

    
    MsgBox "Auto-assignment completed.", vbInformation
End Sub


