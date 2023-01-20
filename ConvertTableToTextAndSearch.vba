Option Explicit

Sub ConvertTableToTextAndSearch()
'
' Convert current table to text, then search for next paragraph with 
' specified formatting
'
    Selection.Rows.ConvertToText Separator:=wdSeparateByParagraphs, _
        NestedTables:=False
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeBackspace
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Heading 4")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
End Sub
