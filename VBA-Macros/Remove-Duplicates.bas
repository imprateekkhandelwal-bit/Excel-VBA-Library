
Attribute VB_Name = "RemoveDuplicates"
Sub RemoveDuplicateRecords()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long

    Set ws = ActiveSheet

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If lastRow < 2 Then
        MsgBox "No data available.", vbInformation
        Exit Sub
    End If

    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)) _
        .RemoveDuplicates Columns:=Evaluate("ROW(1:" & lastCol & ")"), _
        Header:=xlYes

    MsgBox "Duplicate records removed successfully!", vbInformation

End Sub
