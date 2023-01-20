Option Explicit

Sub FormatAllTables()
'
' Sets a sensible default format for all tables
'
    Dim i As Integer
    Dim t As Table
    
    For Each t In ActiveDocument.Tables ' All tables
    'For Each t In Selection.Tables ' Selected tables 
		' Set a predefined style
        t.Style = "Style1"
		' Set vertical alignment (check if this breaks horizontal alignment?)
        t.Range.Cells.VerticalAlignment = wdCellAlignVerticalCenter
        ' Make sure no text wrapping (tables drifting out of position)
        t.Rows.WrapAroundText = False
        ' Don't let rows split across pages
        t.Rows.AllowBreakAcrossPages = False
		
		' Note: following don't work if there are vertically merged/split cells
		
        ' Set the first row to repeat if the table splits across pages
        t.Rows(1).HeadingFormat = True
        ' Set first row to heading style
        t.Rows(1).Range.Style = "Cell Heading"
        ' Set remaining rows to body style
        For i = 2 To t.Rows.Count
            t.Rows(i).Range.Style = "Cell Body"
        Next i
    Next t
End Sub
