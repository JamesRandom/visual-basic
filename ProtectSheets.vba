Option Explicit

Private Sub ProtectSheets()
'
' protect all sheets
    Dim s As Worksheet
    For Each s In ActiveWorkbook.Worksheets
        s.Protect "infineon"
        s.Activate
        s.Cells(1, 1).Select
    Next
    ActiveWorkbook.Worksheets("Data").Visible = xlSheetHidden
    ActiveWorkbook.Worksheets(1).Activate
End Sub
