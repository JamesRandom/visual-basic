Option Explicit

Private Sub MarkTextAsInserted()
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

