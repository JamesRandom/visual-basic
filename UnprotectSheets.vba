Option Explicit

Private Sub UnprotectSheets()
'
' remove protection from sheets
'
    Dim s As Worksheet
    For Each s In ActiveWorkbook.Worksheets
        s.Unprotect "infineon"
        s.Activate
        s.Cells(1, 1).Select
    Next
    ActiveWorkbook.Worksheets("Data").Visible = xlSheetVisible
    ActiveWorkbook.Worksheets(1).Activate
End Sub
