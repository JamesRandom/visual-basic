Option Explicit

Sub TableNoRowSplit()
'
' Change *all* table rows in document so they do not split
' across pages - use with care!
''
    
    Dim t As Table
    For Each t In ActiveDocument.Tables
        t.Rows.AllowBreakAcrossPages = False
    Next t
End Sub
