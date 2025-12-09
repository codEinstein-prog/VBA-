Sub Splitconsumer()
'
' Splitconsumer Macro
' Splits consumer & LTE customers from Business & ICT customers
'
' Keyboard Shortcut: Ctrl+Shift+N
'

    Dim ws As Worksheet
    Dim i As Long, lastrow As Long
    Dim val As Variant
    Dim j As Long
    
    ' Speed up
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Set ws = ActiveSheet
    
    ' Find the last row
    ActiveSheet.UsedRange
    lastrow = ws.UsedRange.Rows(ws.UsedRange.Rows.Count).Row
    
    ' --- Sort by column Z (26th column) ---
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=ws.Range("Z2:Z" & lastrow), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    
    With ws.Sort
        .SetRange ws.Range("A1:Z" & lastrow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Exclusion list
    val = Array("BusinessICT", "VIP Employee", "Third Party Employee", "TELSTAR", "Business")
    
    ' --- Loop bottom to top for column C cleanup ---
    For i = lastrow To 2 Step -1
        For j = LBound(val) To UBound(val)
            ' Use exact match (case-insensitive, trimmed)
            If Trim(UCase(ws.Cells(i, 3).Value)) = UCase(val(j)) Then
                ws.Rows(i).delete
                Exit For
            End If
        Next j
    Next i
    
    ' Restore settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    MsgBox "Split complete!", vbInformation
End Sub


