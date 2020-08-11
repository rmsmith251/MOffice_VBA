'Loops through columns in the row that's clicked and opens all hyperlinks.

Option Explicit

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

Call openisos

End Sub

Public Sub openisos()
Dim CurrentHT As Variant
Dim HTRow As Variant
Dim LinkColumn As Variant
Dim k As Integer
Dim currentcolumn As Variant
Dim CurrentRow As Variant

' Call column and row to make sure that the double clicks are only in the iso column
' and only after row 2 to ensure no errors.
currentcolumn = ActiveCell.Column
CurrentRow = ActiveCell.Row

If currentcolumn = 2 And CurrentRow > 2 Then
' Commented section below is for lookups in a different sheet. Could be handy.
'    CurrentHT = Cells(ActiveCell.Row, 3).Value
'    HTRow = WorksheetFunction.Match(CurrentHT, Sheet2.Range("B:B"), 0)
    'Debug.Print HTRow
' Specify the first column to look in.
    LinkColumn = 15
' Start for loop to pull hyperlinks. Max of 4 in this workbook.
    For k = 1 To 4
        Dim WorkRng As Range
        On Error Resume Next
        Set WorkRng = Sheet1.Cells(CurrentRow, LinkColumn)
        'Debug.Print WorkRng
        ActiveWorkbook.FollowHyperlink Address:=WorkRng, NewWindow:=True
        LinkColumn = LinkColumn + 1
    Next k
End If
End Sub
