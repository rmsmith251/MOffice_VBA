# The following function is used to create a double click action that filters a separate table based on the row/column.

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

Call FilterTPbySubsystem

End Sub

Sub FilterTPbySubsystem()
Application.ScreenUpdating = False
Dim CurrentSS As Variant
Dim currentcolumn As Variant
Dim i As Long
Dim Table1 As ListObject
Dim Table2 As ListObject
Dim Table3 As ListObject
Dim Table4 As ListObject
'Dim Table5 As ListObject

CurrentSS = Cells(ActiveCell.Row, 6).Value
currentcolumn = ActiveCell.Column
'MsgBox CurrentColumn
'Debug.Print CurrentColumn
    
'Filter spools table (table1) if clicking in the spools section
If currentcolumn >= 14 And currentcolumn <= 18 Then
    ActiveWorkbook.Sheets("Spool Tracker").Activate
    Set Table1 = ActiveSheet.ListObjects(1)
    'Debug.Print Table1
    If (ActiveSheet.AutoFilterMode And ActiveSheet.FilterMode) Or ActiveSheet.FilterMode Then
        Table1.AutoFilter.ShowAllData
    End If
    Table1.Range. _
        AutoFilter Field:=21, Criteria1:=CurrentSS
Else
'Next is filter the test pack tracker if clicking in that region.
    If (currentcolumn >= 22 And currentcolumn <= 26) Or (currentcolumn >= 33 And currentcolumn <= 35) Then
        ActiveWorkbook.Sheets("Test Pack Tracker").Activate
        Set Table2 = ActiveSheet.ListObjects(1)
        'Debug.Print Table2
        If (ActiveSheet.AutoFilterMode And ActiveSheet.FilterMode) Or ActiveSheet.FilterMode Then
            Table2.AutoFilter.ShowAllData
        End If
        Table2.Range. _
            AutoFilter Field:=12, Criteria1:=CurrentSS
    Else
'Next is the 3 week look-ahead
        If currentcolumn >= 27 And currentcolumn <= 32 Then
            ActiveWorkbook.Sheets("3 Week Look-Ahead").Activate
            Set Table3 = ActiveSheet.ListObjects(1)
            'Debug.Print Table3
            If (ActiveSheet.AutoFilterMode And ActiveSheet.FilterMode) Or ActiveSheet.FilterMode Then
                Table3.AutoFilter.ShowAllData
            End If
            Table3.Range. _
                AutoFilter Field:=5, Criteria1:=CurrentSS
        Else
'Next is the pre-comm tracker.
            If currentcolumn >= 36 And currentcolumn <= 38 Then
                ActiveWorkbook.Sheets("Pre-Comm Tracker").Activate
                Set Table4 = ActiveSheet.ListObjects(1)
                'Debug.Print Table4
                If (ActiveSheet.AutoFilterMode And ActiveSheet.FilterMode) Or ActiveSheet.FilterMode Then
                    Table4.AutoFilter.ShowAllData
                End If
                Table4.Range. _
                    AutoFilter Field:=8, Criteria1:=CurrentSS
            End If
        End If
    End If
End If
Application.ScreenUpdating = True
End Sub
