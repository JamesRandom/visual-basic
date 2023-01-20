Option Explicit

Sub SplitTablesAtBlankRows()
'
' Split all table at blank rows
'
    Dim t As Table
    Dim r As Row
    
    ' Should be just the selected table
    For Each t In Selection.Tables
        For Each r In t.Rows
            ' A row is empty if there is just a paragraph mark plus
            ' an end mark for each cell and the row.
            If Len(r.Range) = 2 * (r.Cells.Count + 1) Then
                'r.ConvertToText wdSeparateByParagraphs, False
                r.Split
            End If
        Next r
    Next t
End Sub

