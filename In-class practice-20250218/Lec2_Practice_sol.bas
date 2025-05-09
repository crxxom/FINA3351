Attribute VB_Name = "practice"
Sub exercise1()

    MsgBox ThisWorkbook.Worksheets("Record2").Range("G1:I15").Rows.Count
    MsgBox ThisWorkbook.Worksheets("Record2").Range("G1:I15").Columns.Count
    ThisWorkbook.Worksheets("Record2").Range("G1:I15").Copy _
        ThisWorkbook.Worksheets("MySheet").Range("A1")

End Sub

Sub exercise2()

    ThisWorkbook.Worksheets("Record2").Activate
    MsgBox Range("G1:I15").Rows.Count
    MsgBox Range("G1:I15").Columns.Count
    Range("G1:I15").Copy ThisWorkbook.Worksheets("MySheet").Range("A1")

End Sub
