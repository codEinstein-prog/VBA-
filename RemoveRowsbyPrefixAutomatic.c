Sub RemoveRowsByPrefixAutomatic()
    Dim ws As Worksheet
    Dim wb As Workbook, ww As Worksheet
    Dim lastrow As Long, lastPrefix As Long
    Dim dataRange As Range, prefixes As Range
    Dim i As Long, prefix As Range
    Dim checkCol As String
    Dim hasHeader As VbMsgBoxResult
    Dim startRow As Long
    Dim Deletedcount As Long
    Dim filePath As String
    
    Set ws = ActiveSheet
    
    ' === Open prefix workbook ===
    filePath = "C:\Users\SS01\Documents\JCRS DOCCUMENTATION\876 " & "&" & " 658 prefix.xlsx"
    Set wb = GetObject(filePath)
    Set ww = wb.Sheets(1)   ' Prefixes are assumed in col A of sheet1
    
    ' Find last prefix row in column A of prefix file
    lastPrefix = ww.Cells(ww.Rows.Count, "A").End(xlUp).Row
    If lastPrefix < 1 Then
        MsgBox "No prefixes found in column A of prefix file", vbExclamation
        wb.Close False
        Exit Sub
    End If
    
    ' === Ask user for phone number column ===
    checkCol = InputBox("Enter the column letter that contains the phone numbers (e.g., A):", "Phone Number Column")
    If checkCol = "" Then
        wb.Close False
        Exit Sub
    End If
    
    ' === Ask if header row exists ===
    hasHeader = MsgBox("Does your phone number column have a header in row 1?", vbYesNo + vbQuestion, "Header Row")
    If hasHeader = vbYes Then
        startRow = 2
    Else
        startRow = 1
    End If
    
    ' Find last row of data in active sheet
    lastrow = ws.Cells(ws.Rows.Count, checkCol).End(xlUp).Row
    If lastrow < startRow Then
        MsgBox "No phone numbers found in column " & checkCol, vbExclamation
        wb.Close False
        Exit Sub
    End If
    
    Set dataRange = ws.Range(checkCol & startRow & ":" & checkCol & lastrow)
    Set prefixes = ww.Range("A1:A" & lastPrefix)
    
    Deletedcount = 0
    
    ' === Loop bottom to top ===
    For i = dataRange.Rows.Count To 1 Step -1
        For Each prefix In prefixes
            If prefix.Value <> "" Then
                If Left(CStr(dataRange.Cells(i, 1).Value), Len(prefix.Value)) = CStr(prefix.Value) Then
                    dataRange.Cells(i, 1).EntireRow.delete
                    Deletedcount = Deletedcount + 1
                    Exit For
                End If
            End If
        Next prefix
    Next i
    
    wb.Close SaveChanges:=False
    
    MsgBox Deletedcount & " row(s) with listed prefixes were removed.", vbInformation
End Sub

