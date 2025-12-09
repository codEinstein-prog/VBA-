Sub easytextpro()
'
' easytextpro Macro
' automated organization for easytextpro system.
'
' Keyboard Shortcut: Ctrl+Shift+E
'

Dim ws As Worksheet
Dim i As Long
Dim lastrow As Long
Dim cellstatus As Variant
Dim status As Variant
Dim actstatus As Variant
Dim wrongStatus As Variant

    Set ws = ActiveSheet
    
    '--- Sort by Column C
    '--- Loop bottom to top for column C cleanup
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "([0-9])\1{3,}"
    regex.IgnoreCase = True
    
    ActiveSheet.UsedRange
    lastrow = ws.UsedRange.Rows(ws.UsedRange.Rows.Count).Row
    
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=ws.Range("C2:C" & lastrow), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ws.Sort
        .SetRange ws.Range("A1:Z" & lastrow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    For i = lastrow To 2 Step -1
        If Len(Trim(ws.Cells(i, 3).Value)) > 13 Or Len(Trim(ws.Cells(i, 3).Value)) < 7 Then
            ws.Rows(i).delete
        ElseIf IsEmpty(ws.Cells(i, 3).Value) Then
            ws.Rows(i).delete
        ElseIf ws.Cells(i, 3).Value Like "*[A-Za-z]*" Then
            ws.Rows(i).delete
        ElseIf regex.Test(ws.Cells(i, 3).Value) Then
            ws.Rows(i).delete
        End If
    Next i
    
    Call RemoveRowsByPrefixAutomatic
    
    ActiveSheet.UsedRange
    lastrow = ws.UsedRange.Rows(ws.UsedRange.Rows.Count).Row
    
    For i = lastrow To 2 Step -1
        Dim phoneVal As String
        phoneVal = Trim(ws.Cells(i, 3).Value)
        
        If Len(phoneVal) <= 7 Then
            ws.Cells(i, 3).Value = "1876" & phoneVal
        ElseIf Len(phoneVal) >= 9 Then
            ' Only add "1" if number does NOT already start with "1" or "1876"
            If Left(phoneVal, 1) <> "1" And Left(phoneVal, 4) <> "1876" Then
                ws.Cells(i, 3).Value = "1" & phoneVal
            End If
        End If
    Next i
    
    '--- Sort by Column E
    ActiveSheet.UsedRange
    lastrow = ws.UsedRange.Rows(ws.UsedRange.Rows.Count).Row
    
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=ws.Range("E2:E" & lastrow), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ws.Sort
        .SetRange ws.Range("A1:AA" & lastrow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    '--- Delete rows by cell status (Column E)
    cellstatus = Array("NOSVC", "NLIS", "ACNTXT", "NTASS", "NO TXT", "DON TXT", "NO SERV", "NO SERVICE", "NLI", "NOT ASS")
    For i = lastrow To 2 Step -1
        For Each status In cellstatus
            If LCase(Trim(ws.Cells(i, 5).Value)) = LCase(status) Then
                ws.Rows(i).delete
                Exit For
            End If
        Next status
    Next i
    
    '--- Sort by Column F
    ActiveSheet.UsedRange
    lastrow = ws.UsedRange.Rows(ws.UsedRange.Rows.Count).Row
    
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=ws.Range("F2:F" & lastrow), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ws.Sort
        .SetRange ws.Range("A1:AA" & lastrow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    '--- Delete by actstatus (Column F)
    wrongStatus = Array("WRNG", "WRNG#", "wrong", "wrong#")
    actstatus = Array("CLO", "SIF", "PIF", "ARR", "WDRA", "DEC", "LEG", "PPA")
    
    For i = lastrow To 2 Step -1
        For Each status In actstatus
            If LCase(Trim(ws.Cells(i, 6).Value)) = LCase(status) Then
                ws.Rows(i).delete
                Exit For
            End If
        Next status
        
        ' delete SKP only if col E = WRNG variants
        If LCase(Trim(ws.Cells(i, 6).Value)) = "skp" Then
            For Each status In wrongStatus
                If InStr(1, LCase(Trim(ws.Cells(i, 5).Value)), LCase(status)) > 0 Then
                    ws.Rows(i).delete
                    Exit For
                End If
            Next status
        End If
    Next i
    
    
    
    ws.Columns("G:G").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

    ActiveSheet.UsedRange
    lastrow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=Range("G2:G" & lastrow) _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ws.Sort
        .SetRange Range("A1:AA" & lastrow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ActiveSheet.UsedRange
    lastrow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    
    ws.Columns("G:G").NumberFormat = "0.00"
    ws.Range("G2:G" & lastrow).Value = ws.Range("G2:G" & lastrow).Value
    For i = lastrow To 2 Step -1
        If Abs(CDbl(ws.Cells(i, "G").Value)) < 1000 Then
            ws.Rows(i).delete
        End If
    Next i
    
    ActiveSheet.UsedRange
    lastrow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    
    ws.Columns("H:H").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ws.Range("H2:H" & lastrow).FormulaR1C1 = "=(DOLLAR(RC[-1]))"
    
    ' Copy values back to G and delete H
    ws.Range("H2:H" & lastrow).Copy
    ws.Range("G2").PasteSpecial Paste:=xlPasteValues
    ws.Columns("H:H").delete
    
    ' Reformat G as plain number
    ws.Columns("G:G").NumberFormat = "0.00"
    
    ActiveSheet.UsedRange
    lastrow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    For i = lastrow To 2 Step -1
        If (Not IsError(ws.Cells(i, 1).Value)) And _
           (Not IsError(ws.Cells(i, 2).Value)) And _
           (Not IsError(ws.Cells(i, 3).Value)) Then
    
            If Trim(CStr(ws.Cells(i, 1).Value)) = "" And _
               Trim(CStr(ws.Cells(i, 2).Value)) = "" And _
               Trim(CStr(ws.Cells(i, 3).Value)) = "" Then
                ws.Rows(i).delete
            End If
        End If
    Next i
    
    ActiveSheet.UsedRange
    lastrow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    
    ws.Range("$A$1:$AA" & lastrow).RemoveDuplicates _
    Columns:=Array(1, 2, 3), Header:=xlYes
    
End Sub
