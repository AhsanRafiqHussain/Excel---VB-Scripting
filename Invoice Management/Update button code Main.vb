Private Sub CommandButton1_Click()
Dim Invoiceno As String, Company As String, Note As String
Dim Date1 As Date
Dim Detail As Single, Discount As Single, Other As Single, Deposit As Single
    Worksheets("Invoice Entry").Select
    Invoiceno = Range("I4")
    Company = Range("C4")
    Note = Range("C22")
    Date1 = Range("I5")
    Detail = Range("I20")
    Discount = Range("I21")
    Other = Range("I22")
    Deposit = Range("I24")
    Worksheets("Invoices - Main").Select
    Worksheets("Invoices - Main").Range("B3").Select
    If Worksheets("Invoices - Main").Range("B3").Offset(1, 0) <> "" Then
    Worksheets("Invoices - Main").Range("B3").End(xlDown).Select
    End If
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Value = Invoiceno
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = Company
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = Date1
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = Discount
    ActiveCell.Value = Detail
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = Other
    ActiveCell.Offset(0, 2).Select
    ActiveCell.Value = Deposit
    ActiveCell.Offset(0, 2).Select
    ActiveCell.Value = Note
    Worksheets("Invoice Entry").Select
    Worksheets("Invoice Entry").Range("I4").Select