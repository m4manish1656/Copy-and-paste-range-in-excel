# Copy-and-paste-range-in-excel vba
Sub repet()
Dim ws As Worksheet: Set ws = Sheets("Shirt")
'declare and set your worksheet, amend as required

ws.Range("F2:F8").Copy

For i = 1 To 15
    NextRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row + 1
    'get the Next Empty Row
    ws.Range("F" & NextRow).PasteSpecial xlPasteAll 'paste
Next i
End Sub
#this code will copy a raqnge of f2 to f8 and find a blank cell and paste it 15 times in vba.
#happy to help.
