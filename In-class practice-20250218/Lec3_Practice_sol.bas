Attribute VB_Name = "Practice_sol"
Option Explicit

'***************************************************************************
' Q1 SUGGESTED CODES:
'***************************************************************************
Sub Q1()

    Dim S, K, T, r, sig
    Dim dOne As Double
    ThisWorkbook.Worksheets("Q1").Activate
    S = Range("B4")
    K = Range("B5")
    T = Range("B6")
    r = Range("B7")
    sig = Range("B8")
    dOne = (Log(S / K) + (r + sig ^ 2 / 2) * T) / (sig * Sqr(T))
    'Should not use WorksheetFunction.Sqrt, WorksheetFunction.Ln
    Range("B10") = dOne

End Sub

'***************************************************************************
' Q2 SUGGESTED CODES:
'***************************************************************************
Sub Q2()

    Dim Var(1) As String
    'Dim Var(0 to 1) as string
    Var(0) = "Lecture"
    Var(1) = "One"
    ThisWorkbook.Worksheets("Sheet2").Activate
    Rows("1:100").ClearContents
    Range("A1:B1").Value = Var

End Sub

Sub Q2_v2()

    Dim Var()
    Var = Array("Lecture", "One")
    ThisWorkbook.Worksheets("Sheet2").Activate
    Rows("1:100").ClearContents
    Range("A1:B1").Value = Var

End Sub

'***************************************************************************
' Q3 SUGGESTED CODES:
'***************************************************************************
Sub Q3()

    Dim FilmName(1 To 4), FilmType(), RngObject(1 To 3) 'As Object
    Dim i
    Worksheets("Q3").Activate

    FilmName(1) = Range("A2")
    FilmName(2) = Range("A3")
    FilmName(3) = Range("A4")
    FilmName(4) = Range("A5")
'    For i = 1 To 4
'        FilmName(i) = Cells(i + 1, 1)
'    Next

    FilmType = Range("B2:B5")

    Set RngObject(1) = Range("A2:A5")
    Set RngObject(2) = Range("B2:B5")
    Set RngObject(3) = Range("C2:C5")

End Sub



