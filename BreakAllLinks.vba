' Break all links to other files
Private Sub BreakAllLinks()
    Dim aLinks As Variant
    
    ' Get a list of links in the file
    aLinks = ActiveWorkbook.LinkSources(xlExcelLinks)
    If Not IsEmpty(aLinks) Then
        Dim s As Variant
        ' If the list is not empty then got through and break each link
        For Each s In aLinks
            ActiveWorkbook.BreakLink Name:=s, Type:=xlExcelLinks
        Next s
    End If
End Sub
