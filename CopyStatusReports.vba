' Copy status reports from individual files for each WP
' File names should begin with WPn and the corresponding
' sheet is copied into this workbook.
'
Sub CopyReports()
    Dim WPs() As String
    Dim wp As Variant
    Dim wb As Workbook
    Dim main As Sheets
    
    Application.ScreenUpdating = False
    
    ' List of work package numbers (prefixes to file and sheet names
    WPs = Split("WP1,WP2,WP3,WP4,WP5,WP6,WP7,TS1,TS2,TS3,TS4,TS5", ",")
    
    ' The destination file
    Set main = ActiveWorkbook.Worksheets
    ' enable overwriting of contents
    UnprotectSheets
    
    ' Iterate through all WPs
    For Each wp In WPs
        Dim str, path As String
        Dim sheet As String
        Dim link As Variant
        
        ' Name of the sheet containing the status for this WP
        sheet = wp & " Status"
        
        ' Find file to import
        path = ActiveWorkbook.path & "\"
        str = path & "*" & wp & "*"
        str = Dir(str)
        
        ' If file exists then copy appropriate page
        If Len(str) > 0 Then
            ' Open file
            Set wb = Workbooks.Open(path & str, False, True)
            
            ' Remove links
            ' Some workbooks contain links (for some reason).
            ' Some of these are to ishare and so cause an error.
            If Not IsEmpty(wb.LinkSources) Then
                Call UnprotectSheets(wb)
                ' If there are any links, remove them
                For Each link In wb.LinkSources(Type:=xlLinkTypeExcelLinks)
                    wb.BreakLink Name:=link, Type:=xlLinkTypeExcelLinks
                Next link
            End If
            
            ' Find relevant sheet
            Dim ws As Worksheet
            Set ws = wb.Worksheets(sheet)
            
            ' Get the cells with information in
            ws.UsedRange.Copy
            ' Copy to destination file
            main(sheet).Paste
            ' clear copy buffer before closing
            main(sheet).Cells(1, 1).Copy
            ' Close file
            wb.Close SaveChanges:=False
        Else
            ' Mark missing status reports as "unknown"
            
            Dim target As Range
            
            ' Find the "status" cells in the sheet
            Set target = main(sheet).UsedRange.Find("Status", LookIn:=xlFormulas, LookAt:=xlWhole, MatchCase:=True)
            ' Insert the default value
            If (Not target Is Nothing) Then
                target.Offset(1, 0).SpecialCells(xlCellTypeSameValidation).Value = "Unknown"
            End If
        End If
    Next wp
    
    ' Now break the links with the source files this may have created
    ' (This seems to vary with version of Excel)
    BreakAllLinks
    
    ' Finally protect the sheets again and restore everything to normal
    ProtectSheets
    Application.ScreenUpdating = True
    
    main("Project Overview").Activate
End Sub