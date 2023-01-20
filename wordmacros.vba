Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''
Private Sub MarkSelectionAsInserted()
'
' Macro to mark a selection as inserted text
'
    Dim trackChanges As Boolean
    
    ' Record current state of track changes
    trackChanges = ActiveDocument.TrackRevisions
	'Error handling: restore state
    On Error GoTo BombOut
    
    ' If text selected then
    If Selection.Type = wdSelectionNormal Then
        ' Turn off track chnages
        ActiveDocument.TrackRevisions = False
        ' Remove the selected text
        Selection.Cut
        ' Turn on track changes
        ActiveDocument.TrackRevisions = True
        ' Insert the text
        Selection.Paste
    End If
BombOut:
    ' Restore track changes to original state
    ActiveDocument.TrackRevisions = trackChanges
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''
Sub NoRowSplit()
    ' Change *all* table rows in document so they do not split
    ' across pages - use with care!
    
    Dim t As Table
    For Each t In ActiveDocument.Tables
        t.Rows.AllowBreakAcrossPages = False
    Next t
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''
Sub ConvertTableToText()
'
' ConvertTableToText Macro
'
' Convert current table to text then delete spare line at the end
'
    Selection.Rows.ConvertToText Separator:=wdSeparateByParagraphs, _
        NestedTables:=False
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeBackspace
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''
Sub ConvertTableToTextAndSearch()
'
' ConvertTableToTextAndSearch Macro
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
