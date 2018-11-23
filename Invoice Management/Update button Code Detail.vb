Private Sub CommandButton2_Click()
Worksheets("Invoice Entry").Select
    Range("D8:D19").Select
    Selection.Copy
Worksheets("Invoice Details").Select
Worksheets("Invoice Details").Range("D3").Select
If Worksheets("Invoice Details").Range("D3").Offset(1, 0) <> "" Then
Worksheets("Invoice Details").Range("D3").End(xlDown).Select
End If
ActiveCell.Offset(1, -2).Select
ActiveCell.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Worksheets("Invoice Entry").Select
    Range("C8:C19").Select
    Selection.Copy
Worksheets("Invoice Details").Select
Worksheets("Invoice Details").Range("D3").Select
If Worksheets("Invoice Details").Range("D3").Offset(1, 0) <> "" Then
Worksheets("Invoice Details").Range("D3").End(xlDown).Select
End If
ActiveCell.Offset(1, -1).Select
ActiveCell.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Worksheets("Invoice Entry").Select
    Range("G8:G19").Select
    Selection.Copy
Worksheets("Invoice Details").Select
Worksheets("Invoice Details").Range("D3").Select
If Worksheets("Invoice Details").Range("D3").Offset(1, 0) <> "" Then
Worksheets("Invoice Details").Range("D3").End(xlDown).Select
End If
ActiveCell.Offset(1, 1).Select
ActiveCell.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Worksheets("Invoice Entry").Select
    Range("H8:H19").Select
    Selection.Copy
Worksheets("Invoice Details").Select
Worksheets("Invoice Details").Range("D3").Select
If Worksheets("Invoice Details").Range("D3").Offset(1, 0) <> "" Then
Worksheets("Invoice Details").Range("D3").End(xlDown).Select
End If
ActiveCell.Offset(1, 2).Select
ActiveCell.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
 Worksheets("Invoice Entry").Select
    Range("I8:I19").Select
    Selection.Copy
Worksheets("Invoice Details").Select
Worksheets("Invoice Details").Range("D3").Select
If Worksheets("Invoice Details").Range("D3").Offset(1, 0) <> "" Then
Worksheets("Invoice Details").Range("D3").End(xlDown).Select
End If
ActiveCell.Offset(1, 3).Select
ActiveCell.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Worksheets("Invoice Entry").Select
    Range("J8:J19").Select
    Selection.Copy
Worksheets("Invoice Details").Select
Worksheets("Invoice Details").Range("D3").Select
If Worksheets("Invoice Details").Range("D3").Offset(1, 0) <> "" Then
Worksheets("Invoice Details").Range("D3").End(xlDown).Select
End If
ActiveCell.Offset(1, 4).Select
ActiveCell.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Worksheets("Invoice Entry").Select
    Range("B8:B19").Select
    Selection.Copy
Worksheets("Invoice Details").Select
Worksheets("Invoice Details").Range("D3").Select
If Worksheets("Invoice Details").Range("D3").Offset(1, 0) <> "" Then
Worksheets("Invoice Details").Range("D3").End(xlDown).Select
End If
ActiveCell.Offset(1, 0).Select
ActiveCell.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False




End Sub
