Sub ConcatenateColumnsCAndD_AB()
    Dim ws As Worksheet
    Dim lastRowC As Long, lastRowD As Long
    Dim i As Long
    Dim combinedString As String
    Dim fixedText As String
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("BL-BID") ' Change the sheet index or name as needed

    ' Find the last row with data in columns C and D
    lastRowC = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    lastRowD = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row

    ' Find the maximum of the two last rows
    Dim lastRow As Long
    lastRow = Application.WorksheetFunction.Max(lastRowC, lastRowD)
    
    ' Loop through the rows starting from row 3
    For i = 3 To lastRow
        ' Concatenate data from columns C and D with a comma separator
        If ws.Cells(i, "C").Value <> "" And ws.Cells(i, "D").Value <> "" Then
            combinedString = combinedString & ws.Cells(i, "C").Value & "(" & ws.Cells(i, "D").Value & "), "
        End If
    Next i

    ' Remove the trailing comma and space
    If Len(combinedString) > 0 Then
        combinedString = Left(combinedString, Len(combinedString) - 2)
    End If

    ' Combine the fixed text and the concatenated string
    combinedString = fixedText & combinedString

    ' Merge cells in column F from row 3 to row 30
    With ws.Range("F3:F30")
        .Merge
        .Value = combinedString
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
    End With

End Sub

