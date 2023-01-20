Option Explicit

Sub ConvertTableToText()
'
' Convert current table to text then delete spare line at the end
'
    Selection.Rows.ConvertToText Separator:=wdSeparateByParagraphs, _
        NestedTables:=False
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeBackspace
End Sub
