Private Const StandardLeftIndent = 13 ' millimitres
Private Const StandardBodyFontName = "Verdana"
Private Const StandardBodyFontSize = 9
Private Const StandardHeaderFontName = "Arial"
Private Const CompanyName = "XXXXXX Ltd"
Private Const ConfidentialNotice = CompanyName & " Company Confidential"



' Update all the styles to the standard
Public Sub UpdateStyles()

    If ActiveDocument.Path = "" Then
        Err.Raise Number:=1, Description:="Document not saved; save the file in the same directory as the image files before continuing"
    End If

    SetDocumentProperties
    SetPageLayout

    ' Start new page with table of contents
    FormatTOCHeading
    Dim p As Paragraph
    ' Add TOC in new paragraph
    Set p = ActiveDocument.Paragraphs.Add
    Set p = ActiveDocument.Paragraphs.Add
    ' Insert table of contents
    ActiveDocument.Fields.Add Range:=p.Range, Type:=wdFieldTOC, Text:="\o ""1-3"" \h \z", PreserveFormatting:=False

    ' Format the styles we need for the front page
    FormatHeading1
    FormatNormal
    FormatTitle
    FormatSubTitle
    FormatHeader
    FormatFooter

    FillPageHeadersFooters
    CreateFrontPage

    ' Format all the remaining paragraphs
    FormatHeadingNumbering
    FormatHeading2
    FormatHeading3
    FormatHeading4
    FormatHeading5
    FormatHeading6
    FormatHeading7
    FormatHeading8
    FormatHeading9
    FormatQuote
    FormatList
    FormatListBullet
    FormatListNumber
    FormatListContinue
    FormatCode
    FormatFootnoteText
    FormatCellBody
    FormatCellHeading
    FormatFigure
    FormatTableTitle
    FormatTOC

End Sub


' Generate a list of all the styles in the document, in their own format
Public Sub ExportStylesInDocument()

    Dim oDoc As Document
    Dim oPara As Paragraph
    Dim oStyle As Style

    With ActiveDocument
        For Each oStyle In .Styles
            If oStyle.Type = wdStyleTypeParagraph Then
                Set oPara = .Paragraphs.Add
                Set oPara = .Paragraphs.Add
                oPara.Range.Text = oStyle.NameLocal
                oPara.Style = oStyle
            End If
        Next
    End With
End Sub


' Rename all Framemaker styles to match the appropriate standard Word ones
' Remove Framemaker styles from document
Public Sub ConvertFrameStyles()
    Dim p As Paragraph
    Dim toBeDeleted As New Collection
    With ActiveDocument
        For Each p In .Paragraphs
            Select Case p.Style
                Case "ChapterTitle", "Preface", "Contents"
                    p.Style = .Styles("Heading 1")
                Case "DocumentTitle"
                    p.Style = .Styles("Title")
                Case "DocumentSubTitle"
                    p.Style = .Styles("Subtitle")
                Case "Body", "Note", "Quote"
                    p.Style = .Styles("Normal")
                Case "Heading1", "Heading1Page", "PrefaceHeading1"
                    p.Style = .Styles("Heading 2")
                Case "Heading2", "PrefaceHeading2"
                    p.Style = .Styles("Heading 3")
                Case "Heading3"
                    p.Style = .Styles("Heading 4")
                Case "Heading4", "HeadingRunIn"
                    p.Style = .Styles("Heading 5")
                Case "Bulleted", "Bulleted+"
                    p.Style = .Styles("List Bullet")
                Case "Bulleted2"
                    p.Style = .Styles("List Bullet 2")
                Case "Bulleted3"
                    p.Style = .Styles("List Bullet 3")
                Case "Alpha", "Alpha Start", "Roman Start", "Numbered Start"
                    p.Style = .Styles("List Number")
                    p.Style.ListTemplate.ListLevels(1).StartAt = 1
                Case "Alpha", "Roman", "Numbered"
                    p.Style = .Styles("List Number")
                Case "Alpha2", "Roman2", "Numbered2"
                    p.Style = .Styles("List Number 2")
                Case "Alpha3", "Roman3", "Numbered3"
                    p.Style = .Styles("List Number 3")
                Case "Level1Cont"
                    p.Style = .Styles("List Continue")
                Case "Level2Cont"
                    p.Style = .Styles("List Continue 2")
                Case "Level3Cont"
                    p.Style = .Styles("List Continue 3")
                Case "TableFootnote", "Footnote"
                    p.Style = .Styles("Footnote Text")
            End Select
        Next

        On Error Resume Next
        .Styles("ChapterTitle").Delete
        .Styles("Preface").Delete
        .Styles("Contents").Delete
        .Styles("DocumentTitle").Delete
        .Styles("Body").Delete
        .Styles("Note").Delete
        .Styles("Quote").Delete
        .Styles("DocumentSubTitle").Delete
        .Styles("Heading1").Delete
        .Styles("Heading1Page").Delete
        .Styles("PrefaceHeading1").Delete
        .Styles("Heading2").Delete
        .Styles("PrefaceHeading2").Delete
        .Styles("Heading3").Delete
        .Styles("Heading4").Delete
        .Styles("HeadingRunIn").Delete
        .Styles("Bulleted").Delete
        .Styles("Bulleted+").Delete
        .Styles("Bulleted2").Delete
        .Styles("Bulleted3").Delete
        .Styles("Alpha").Delete
        .Styles("Alpha Start").Delete
        .Styles("Roman").Delete
        .Styles("Roman Start").Delete
        .Styles("Numbered").Delete
        .Styles("Numbered Start").Delete
        .Styles("Alpha2").Delete
        .Styles("Roman2").Delete
        .Styles("Numbered2").Delete
        .Styles("Alpha3").Delete
        .Styles("Roman3").Delete
        .Styles("Numbered3").Delete
        .Styles("Level1Cont").Delete
        .Styles("Level2Cont").Delete
        .Styles("Level3Cont").Delete
        .Styles("TableFootnote").Delete
        .Styles("Footnote").Delete
    End With
End Sub



' Rename all Framemaker styles to match the appropriate standard Word ones
' Remove Framemaker styles from document
Public Sub ConvertOldWordStyles()
    Dim p As Paragraph
    Dim toBeDeleted As New Collection
    With ActiveDocument
        For Each p In .Paragraphs
            Select Case p.Style
                Case "TOC"
                    p.Style = .Styles("TOC Heading")

                Case "Section Heading", "Appendix Heading"
                    p.Style = .Styles("Heading 1")

                Case "Front Page Heading 1"
                    p.Style = .Styles("Title")

                Case "Standard Table Title"
                    p.Style = .Styles("CellHeading")

                Case "Standard Table Text"
                    p.Style = .Styles("CellBody")

                Case "Standard Body"
                    p.Style = .Styles("Normal")

                Case "Section Sub Heading"
                    p.Style = .Styles("Heading 2")

                Case "Section Sub Sub Heading"
                    p.Style = .Styles("Heading 3")

                Case "Standard Code"
                    p.Style = .Styles("Code")

                Case "Standard Bullet", "Bullet List"
                    p.Style = .Styles("List Bullet")

                Case "Standard Bullet 2"
                    p.Style = .Styles("List Bullet 2")

                Case "Standard List"
                    p.Style = .Styles("List Number")

                Case "Standard Indent"
                    p.Style = .Styles("List")

            End Select
        Next

        On Error Resume Next
        .Styles("TOC").Delete
        .Styles("Section Heading").Delete
        .Styles("Appendix Heading").Delete
        .Styles("Front Page Heading 1").Delete
        .Styles("Standard Body").Delete
        .Styles("Standard Code").Delete
        .Styles("Standard Table Title").Delete
        .Styles("DocumentSubTitle").Delete
        .Styles("Standard Table Text").Delete
        .Styles("Section Sub Heading").Delete
        .Styles("Section Sub Sub Heading").Delete
        .Styles("Standard Bullet").Delete
        .Styles("Bullet List").Delete
        .Styles("Standard List").Delete
        .Styles("Standard Indent").Delete
    End With
End Sub


' Debugging function
Private Sub getStyle()
    On Error Resume Next
    ExportStylesInDocument
    Dim s As Style
    Set s = ActiveDocument.Styles("Heading 2")
    Dim lt As ListTemplates
    Set lt = s.ListTemplate

End Sub


Private Sub SetDocumentProperties()
    Dim p As DocumentProperty
    On Error Resume Next

    ' If Revision property does not exist, create it
    Set p = ActiveDocument.CustomDocumentProperties("Revision")
    If p Is Nothing Then
        ActiveDocument.CustomDocumentProperties.Add name:="Revision", LinkToContent:=False, Type:=msoPropertyTypeString, Value:="1.00.00"
    End If

    ' If Copyright property does not exist, create it
    Set p = ActiveDocument.CustomDocumentProperties("Copyright")
    If p Is Nothing Then
        ActiveDocument.CustomDocumentProperties.Add name:="Copyright", LinkToContent:=False, Type:=msoPropertyTypeString, Value:=CompanyName
    End If

    ' If Title property is empty, then set it
    Set p = ActiveDocument.BuiltInDocumentProperties(wdPropertyTitle)
    If p.Value = "" Then
        p.Value = "Document Title"
    End If
    ' Set Company
    ActiveDocument.BuiltInDocumentProperties(wdPropertyCompany) = CompanyName
    ActiveDocument.BuiltInDocumentProperties(wdPropertyAuthor) = CompanyName
End Sub


Private Sub FillPageHeadersFooters()
   ' Set the content and headers and footers
    With ActiveDocument.Sections(1)
        Dim f As Field
        Dim r As Range

        ' set header for odd pages
        Set r = .Headers(wdHeaderFooterPrimary).Range
        r.Delete
        ActiveDocument.Fields.Add Range:=r, Type:=wdFieldStyleRef, Text:="""Heading 1""", PreserveFormatting:=True
        r.InsertBefore (vbTab & vbTab)

        ' set header for even pages
        Set r = .Headers(wdHeaderFooterEvenPages).Range
        r.Delete
        ActiveDocument.Fields.Add Range:=r, Type:=wdFieldDocProperty, Text:="Title", PreserveFormatting:=True

        ' set footer for odd pages
        Set r = .Footers(wdHeaderFooterPrimary).Range
        r.Delete
        ' Insert footer text
        r.InsertAfter (vbTab & vbTab & ConfidentialNotice)
        ' Insert page number after first tab
        r.SetRange Start:=r.Start + 1, End:=r.Start + 1
        ActiveDocument.Fields.Add Range:=r, Type:=wdFieldPage

        ' set footer for even pages
        Set r = .Footers(wdHeaderFooterEvenPages).Range
        r.Delete
        ' Insert footer text
        r.InsertBefore ("Revision " & vbTab)
        ' Followed by page number
        r.SetRange Start:=r.End, End:=r.End
        ActiveDocument.Fields.Add Range:=r, Type:=wdFieldPage
        ' Insert revision number before tab
        r.SetRange Start:=r.End - 1, End:=r.End - 1
        ActiveDocument.Fields.Add Range:=r, Type:=wdFieldDocProperty, Text:="Revision", PreserveFormatting:=True

        ' Remove border from first page
        .Footers(wdHeaderFooterFirstPage).Range.ParagraphFormat.Borders.Enable = False

        ' Insert images in front page header
        Set r = .Headers(wdHeaderFooterFirstPage).Range
        r.Delete
        ' Insert Border
        With .Headers(wdHeaderFooterFirstPage).Shapes.AddPicture(FileName:=ActiveDocument.Path & "\border.png", linktofile:=False, savewithdocument:=True)
            .name = "Border Image"
            ' Set absolute page positioning
            .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
            .RelativeVerticalPosition = wdRelativeVerticalPositionPage
            ' Format shape
            .WrapFormat.Type = wdWrapNone
            .WrapFormat.AllowOverlap = True
            .LockAnchor = False
            .ZOrder msoSendBehindText
            ' Set position and size
            .Width = MillimetersToPoints(184)
            .Height = MillimetersToPoints(255)
            .Left = MillimetersToPoints(17.5)
            .Top = MillimetersToPoints(19)
        End With
        ' Insert Logo
        With .Headers(wdHeaderFooterFirstPage).Shapes.AddPicture(FileName:=ActiveDocument.Path & "\Logo.eps", linktofile:=False, savewithdocument:=True)
            .name = "Logo"
            ' Set absolute page positioning
            .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
            .RelativeVerticalPosition = wdRelativeVerticalPositionPage
            ' Format shape
            .WrapFormat.Type = wdWrapNone
            .WrapFormat.AllowOverlap = True
            .LockAnchor = False
            .ZOrder msoSendBehindText
            ' Set position and size
            .Width = MillimetersToPoints(62)
            .Height = MillimetersToPoints(16)
            .Left = MillimetersToPoints(137)
            .Top = MillimetersToPoints(22)
        End With
    End With

    ' Switch back to text view (close header/footer view)
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
End Sub


Private Sub SetPageLayout()
    ' Set page size and margings, headers and footers
    With ActiveDocument.PageSetup
        .DifferentFirstPageHeaderFooter = True
        .OddAndEvenPagesHeaderFooter = True
        .MirrorMargins = True
        .LeftMargin = Application.MillimetersToPoints(25)
        .RightMargin = Application.MillimetersToPoints(20)
        .PaperSize = wdPaperA4
    End With
End Sub

Private Sub CreateFrontPage()
    Dim s As Shape
    ' Create text box for title
    Set s = ActiveDocument.Shapes.AddTextbox(msoTextOrientationHorizontal, MillimetersToPoints(102), MillimetersToPoints(148), MillimetersToPoints(78), MillimetersToPoints(64))
    With s
        '.name = "Title Text Box"
        ' Set absloute positioning
        .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
        .RelativeVerticalPosition = wdRelativeVerticalPositionPage
        ' Set layout on page
        .WrapFormat.Type = wdWrapNone
        .WrapFormat.AllowOverlap = True
        .LockAnchor = True
        .Line.Visible = msoFalse
        With .TextFrame
            Dim p As Paragraph
            ' First paragraph
            Set p = .TextRange.Paragraphs(1)
            ' Insert document title property
            ActiveDocument.Fields.Add Range:=p.Range, Type:=wdFieldDocProperty, Text:="Title", PreserveFormatting:=True
            ' Set to Title style
            p.Style = ActiveDocument.Styles("Title")

            ' Add next paragraph
            Set p = .TextRange.Paragraphs.Add
            Set p = .TextRange.Paragraphs.Add
            ' Add placeholder text
            p.Range.Text = "Subtitle"
            ' Set to Subtitle style
            p.Style = ActiveDocument.Styles("SubTitle")

            ' Add next paragraph
            Set p = .TextRange.Paragraphs.Add
            Set p = .TextRange.Paragraphs.Add
            ' Insert revision number document property
            ActiveDocument.Fields.Add Range:=p.Range, Type:=wdFieldDocProperty, Text:="Revision", PreserveFormatting:=True
            ' Set to Normal style
            p.Style = ActiveDocument.Styles("Normal")
            ' But without indent
            p.LeftIndent = 0

            ' Set a left border
            With .ContainingRange.Borders(wdBorderLeft)
                .LineStyle = wdLineStyleSingle
                .Color = wdBlack
                .LineWidth = wdLineWidth050pt
            End With
            .ContainingRange.Borders.DistanceFromLeft = 12
        End With
    End With

    ' Create text box for confidential notice
    Set s = ActiveDocument.Shapes.AddTextbox(msoTextOrientationHorizontal, MillimetersToPoints(20), MillimetersToPoints(260), MillimetersToPoints(178), MillimetersToPoints(10))
    With s
        '.name = "Confidential Text Box"
        ' Set absloute positioning
        .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
        .RelativeVerticalPosition = wdRelativeVerticalPositionPage
        ' Set layout on page
        .WrapFormat.Type = wdWrapNone
        .WrapFormat.AllowOverlap = True
        .LockAnchor = True
        .Line.Visible = msoFalse

        With .TextFrame
            ' First paragraph
            With .TextRange.Paragraphs(1)
                .Range.Text = ConfidentialNotice
                ' Set to Normal style
                .Style = ActiveDocument.Styles("Normal")
                ' But without indent
                .LeftIndent = 0
                .Alignment = wdAlignParagraphCenter
                .Range.Font.Size = 10
                .Range.Font.Bold = True
            End With
        End With
    End With
End Sub



' Update Normal style to our defaults
Private Sub FormatNormal()
    Dim s As Style
    Set s = GetParaStyle("Normal")
    With s
        ' Font
        With .Font
            .Size = StandardBodyFontSize
            .name = StandardBodyFontName
            .Bold = False
            .Italic = False
            .Color = wdColorBlack
        End With
        ' Paragraph layout
        With .ParagraphFormat
            .Alignment = wdAlignParagraphLeft
            .FirstLineIndent = 0
            .LeftIndent = MillimetersToPoints(StandardLeftIndent)
            .LineSpacingRule = wdLineSpaceAtLeast
            .LineSpacing = 12
            .SpaceBefore = 11
            .SpaceAfter = 5
        End With
        .LanguageID = wdEnglishUS
    End With

End Sub

' Set the bulleted list style
Private Sub FormatListBullet()
    Dim n As String
    Dim lt As ListTemplate
    Set lt = ListGalleries(wdBulletGallery).ListTemplates(1)

    For i = 1 To 5
        n = "List Bullet"
        If i > 1 Then n = n & " " & i

        ' Set up the bullet format
        With lt.ListLevels(1)
            .Font.Color = RGB(249, 152, 6) ' Orange
            .Font.name = "Wingdings"
            .Font.Size = 7
            .NumberStyle = wdListNumberStyleBullet
            ' Choose a character for each level
            .NumberFormat = Mid("nluonlu", i, 1)
        End With

        ' then the paragraph styles
        Dim s As Style
        Set s = GetParaStyle(n)
        With s
            .BaseStyle = ActiveDocument.Styles("Normal")
            Call .LinkToListTemplate(lt)
            With .ParagraphFormat
                ' calculate appropriate indents for each level:
                ' bullet indented by 3mm
                ' +5mm at each level
                .LeftIndent = MillimetersToPoints(StandardLeftIndent + 3 + i * 8)
                .FirstLineIndent = -MillimetersToPoints(8)
                .TabStops.ClearAll
                .TabStops.Add (MillimetersToPoints(StandardLeftIndent + 3 + i * 8))
                .LineSpacingRule = wdLineSpaceAtLeast
                .LineSpacing = 12
                .SpaceBefore = 6
                .SpaceAfter = 4
            End With
        End With
    Next
End Sub


' Set format for unnumbered lsist
Private Sub FormatList()
    Dim n As String
    For i = 1 To 5
        n = "List"
        If i > 1 Then n = n & " " & i

        Dim s As Style
        Set s = GetParaStyle(n)
        s.BaseStyle = ActiveDocument.Styles("Normal")
        With s.ParagraphFormat
            ' calculate appropriate indents for each level:
            ' +8mm at each level
            .LeftIndent = MillimetersToPoints(StandardLeftIndent + 3 + i * 8)
            .FirstLineIndent = 0
            .LineSpacingRule = wdLineSpaceAtLeast
            .LineSpacing = 12
            .SpaceBefore = 6
            .SpaceAfter = 4
        End With
    Next
End Sub


' Set format for list continuation paragraphs
Private Sub FormatListContinue()
    Dim n As String
    For i = 1 To 5
        n = "List Continue"
        If i > 1 Then n = n & " " & i

        Dim s As Style
        Set s = GetParaStyle(n)
        s.BaseStyle = ActiveDocument.Styles("Normal")
        With s.ParagraphFormat
            .LeftIndent = MillimetersToPoints(StandardLeftIndent + 3 + i * 8)
            .FirstLineIndent = 0
            .LineSpacingRule = wdLineSpaceAtLeast
            .LineSpacing = 12
            .SpaceBefore = 2
            .SpaceAfter = 6
        End With
    Next
End Sub


' Set formatting and numbering for numbered lists
Private Sub FormatListNumber()
    Dim n As String
    Dim lt As ListLevel

    For i = 1 To 5
        n = "List Number"
        If i > 1 Then n = n & " " & i

        Dim s As Style
        Set s = GetParaStyle(n)
        With s
            .BaseStyle = ActiveDocument.Styles("Normal")
            ' Set up numbering
            Set lt = .ListTemplate.ListLevels(1)
            With lt
                .NumberStyle = wdListNumberStyleArabic
                .NumberFormat = "%1."
                .TrailingCharacter = wdTrailingTab
                .Alignment = wdListLevelAlignLeft
                .NumberPosition = MillimetersToPoints(StandardLeftIndent + 3 + (i - 1) * 8)
                .TabPosition = MillimetersToPoints(StandardLeftIndent + 3 + i * 8)
                .TextPosition = MillimetersToPoints(StandardLeftIndent + 3 + i * 8)
            End With
            'Call .LinkToListTemplate(lt)
            ' The paragraph styles
            With .ParagraphFormat
                ' calculate appropriate indents for each level:
                ' number indented by 3mm
                ' +8mm at each level
                .LeftIndent = MillimetersToPoints(StandardLeftIndent + 3 + i * 8)
                .FirstLineIndent = -MillimetersToPoints(8)
                .TabStops.ClearAll
                .TabStops.Add (MillimetersToPoints(StandardLeftIndent + 3 + i * 8))
                .LineSpacingRule = wdLineSpaceAtLeast
                .LineSpacing = 12
                .SpaceBefore = 6
                .SpaceAfter = 4
            End With
        End With
    Next
End Sub

' Set format for table of content paragraphs
Private Sub FormatTOC()
    Dim n As String
    For i = 1 To 9
        n = "TOC " & i

        With GetParaStyle(n)
            .BaseStyle = ActiveDocument.Styles("Normal")
            If i = 1 Then
                With .Font
                    .Size = 12
                    .Bold = True
                End With
            Else
                With .Font
                    .Size = 10
                    .Bold = False
                End With
            End If
            With .ParagraphFormat
                .TabStops.ClearAll
                .FirstLineIndent = 0
                If i > 2 Then
                    .LeftIndent = MillimetersToPoints(StandardLeftIndent + (i - 2) * 8)
                    .TabStops.Add (MillimetersToPoints(StandardLeftIndent + (i - 1) * 8))
                Else
                    .LeftIndent = MillimetersToPoints(StandardLeftIndent)
                    .TabStops.Add (MillimetersToPoints(StandardLeftIndent + 8))
                End If
                Call .TabStops.Add(MillimetersToPoints(165), wdAlignTabRight, wdTabLeaderDots)
                .LineSpacingRule = wdLineSpaceAtLeast
                .LineSpacing = 12
                If i = 1 Then
                    .SpaceBefore = 18
                Else
                    .SpaceBefore = 6
                End If
                .SpaceAfter = 5
            End With
        End With
    Next

    n = "Table of Figures"
    With GetParaStyle(n)
        .BaseStyle = ActiveDocument.Styles("Normal")
        .Font.Size = 10
        With .ParagraphFormat
            .TabStops.ClearAll
            .FirstLineIndent = 0
            .LeftIndent = MillimetersToPoints(StandardLeftIndent)
            Call .TabStops.Add(MillimetersToPoints(165), wdAlignTabRight, wdTabLeaderDots)
            .LineSpacingRule = wdLineSpaceAtLeast
            .LineSpacing = 12
            .SpaceBefore = 6
            .SpaceAfter = 5
        End With
    End With
End Sub


' Define the numbering template to be used for headings
Private Sub FormatHeadingNumbering()
    s = "%1"
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1)
        For i = 1 To .ListLevels.Count
            With .ListLevels(i)
                .NumberStyle = wdListNumberStyleArabic
                .NumberFormat = s
                s = s & ".%" & i + 1
                .TrailingCharacter = wdTrailingTab
                .NumberPosition = 0
                .TabPosition = MillimetersToPoints(StandardLeftIndent)
                .TextPosition = MillimetersToPoints(StandardLeftIndent)
                .Alignment = wdListLevelAlignLeft
                .StartAt = 1
                .LinkedStyle = "Heading " & i
            End With
        Next
    End With
End Sub


' Update TOC Heading style to our defaults
Private Sub FormatTOCHeading()
    Dim s As Style
    Set s = GetParaStyle("TOC Heading")
    With s
        .BaseStyle = ActiveDocument.Styles("Normal")
        .NextParagraphStyle = ActiveDocument.Styles("Normal")
        .LanguageID = wdEnglishUS
        ' Font
        With .Font
            .Size = 16
            .name = StandardHeaderFontName
            .Bold = True
            .Italic = False
            .Color = RGB(77, 77, 77)
        End With
        ' Paragraph layout
        With .ParagraphFormat
            .LeftIndent = 2 * MillimetersToPoints(StandardLeftIndent)
            .TabStops.Add (MillimetersToPoints(StandardLeftIndent))
            .FirstLineIndent = -MillimetersToPoints(StandardLeftIndent)
            .LineSpacingRule = wdLineSpaceAtLeast
            .LineSpacing = 15
            .SpaceBefore = 30
            .SpaceAfter = 7
            .KeepWithNext = True
            .KeepTogether = True
            .OutlineLevel = wdOutlineLevelBodyText
            .PageBreakBefore = True
        End With
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .Color = RGB(249, 152, 6) ' Orange
            .LineWidth = wdLineWidth050pt
        End With
        .Borders.DistanceFromBottom = 12
    End With

End Sub


' Update Heading 1 style to our defaults
Private Sub FormatHeading1()
    Dim s As Style
    Set s = GetParaStyle("Heading 1")
    With s
        .BaseStyle = ActiveDocument.Styles("Normal")
        .NextParagraphStyle = ActiveDocument.Styles("Normal")
        .LanguageID = wdEnglishUS
        Call .LinkToListTemplate(ListGalleries(wdOutlineNumberGallery).ListTemplates(1), 1)
        .ListTemplate.OutlineNumbered = True
        ' Font
        With .Font
            .Size = 16
            .name = StandardHeaderFontName
            .Bold = True
            .Italic = False
            .Color = RGB(77, 77, 77)
        End With
        ' Paragraph layout
        With .ParagraphFormat
            .LeftIndent = MillimetersToPoints(StandardLeftIndent)
            .FirstLineIndent = -MillimetersToPoints(StandardLeftIndent)
            .TabStops.ClearAll
            .TabStops.Add (MillimetersToPoints(StandardLeftIndent))
            .LineSpacingRule = wdLineSpaceAtLeast
            .LineSpacing = 15
            .SpaceBefore = 30
            .SpaceAfter = 7
            .KeepWithNext = True
            .KeepTogether = True
            .OutlineLevel = wdOutlineLevel1
            .PageBreakBefore = True
        End With
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .Color = RGB(249, 152, 6) ' Orange
            .LineWidth = wdLineWidth050pt
        End With
        .Borders.DistanceFromBottom = 12
    End With

End Sub


' Update Heading 2 style to our defaults
Private Sub FormatHeading2()
    Dim s As Style
    Set s = GetParaStyle("Heading 2")
    With s
        .BaseStyle = ActiveDocument.Styles("Heading 1")
        .Borders.Enable = False
        .LanguageID = wdEnglishUS
        'Call .LinkToListTemplate(ListGalleries(wdOutlineNumberGallery).ListTemplates(1), 2)
        ' Font
        With .Font
            .Size = 13
            .name = StandardHeaderFontName
            .Bold = True
            .Italic = False
            .Color = RGB(77, 77, 77)
        End With
        ' Paragraph layout
        With .ParagraphFormat
            .LeftIndent = MillimetersToPoints(StandardLeftIndent)
            .FirstLineIndent = -MillimetersToPoints(StandardLeftIndent)
            .TabStops.ClearAll
            .TabStops.Add (MillimetersToPoints(StandardLeftIndent))
            .LineSpacingRule = wdLineSpaceAtLeast
            .LineSpacing = 15
            .SpaceBefore = 24
            .SpaceAfter = 7
            .KeepWithNext = True
            .KeepTogether = True
            .OutlineLevel = wdOutlineLevel2
            .PageBreakBefore = False
        End With
    End With

End Sub


' Update Heading 3 style to our defaults
Private Sub FormatHeading3()
    Dim s As Style
    Set s = GetParaStyle("Heading 3")
    With s
        .BaseStyle = ActiveDocument.Styles("Heading 1")
        .Borders.Enable = False
        .LanguageID = wdEnglishUS
        ' Font
        With .Font
            .Size = 11
            .name = StandardHeaderFontName
            .Bold = True
            .Italic = False
            .Color = RGB(77, 77, 77)
        End With
        ' Paragraph layout
        With .ParagraphFormat
            .LeftIndent = MillimetersToPoints(StandardLeftIndent)
            .FirstLineIndent = -MillimetersToPoints(StandardLeftIndent)
            .TabStops.ClearAll
            .TabStops.Add (MillimetersToPoints(StandardLeftIndent))
            .LineSpacingRule = wdLineSpaceAtLeast
            .LineSpacing = 15
            .SpaceBefore = 20
            .SpaceAfter = 7
            .KeepWithNext = True
            .KeepTogether = True
            .OutlineLevel = wdOutlineLevel3
            .PageBreakBefore = False
        End With
        'Call .LinkToListTemplate(ListGalleries(wdOutlineNumberGallery).ListTemplates(1), 3)
    End With

End Sub


' Update Heading 4 style to our defaults
Private Sub FormatHeading4()
    Dim s As Style
    Set s = GetParaStyle("Heading 4")
    With s
        .BaseStyle = ActiveDocument.Styles("Heading 1")
        .LinkToListTemplate ListTemplate:=Nothing
        .Borders.Enable = False
        .LanguageID = wdEnglishUS
        ' Font
        With .Font
            .Size = 11
            .name = StandardHeaderFontName
            .Bold = True
            .Italic = False
            .Color = RGB(77, 77, 77)
        End With
        ' Paragraph layout
        With .ParagraphFormat
            .LeftIndent = MillimetersToPoints(StandardLeftIndent)
            .FirstLineIndent = 0
            .LineSpacingRule = wdLineSpaceAtLeast
            .LineSpacing = 15
            .SpaceBefore = 18
            .SpaceAfter = 7
            .KeepWithNext = True
            .KeepTogether = True
            .OutlineLevel = wdOutlineLevel4
            .PageBreakBefore = False
        End With
    End With

End Sub


' Update Heading 5 style to our defaults
Private Sub FormatHeading5()
    Dim s As Style
    Set s = GetParaStyle("Heading 5")
    With s
        .BaseStyle = ActiveDocument.Styles("Heading 1")
        .LinkToListTemplate ListTemplate:=Nothing
        .Borders.Enable = False
        .LanguageID = wdEnglishUS
        ' Font
        With .Font
            .Size = 9
            .name = StandardHeaderFontName
            .Bold = True
            .Italic = False
            .Color = RGB(77, 77, 77)
        End With
        ' Paragraph layout
        With .ParagraphFormat
            .LeftIndent = MillimetersToPoints(StandardLeftIndent)
            .FirstLineIndent = 0
            .LineSpacingRule = wdLineSpaceAtLeast
            .LineSpacing = 15
            .SpaceBefore = 15
            .SpaceAfter = 7
            .KeepWithNext = True
            .KeepTogether = True
            .OutlineLevel = wdOutlineLevel5
            .PageBreakBefore = False
        End With
    End With

End Sub




' Update Heading 6 style to our defaults
Private Sub FormatHeading6()
    Dim s As Style
    Set s = GetParaStyle("Heading 6")
    With s
        .BaseStyle = ActiveDocument.Styles("Heading 1")
        .LinkToListTemplate ListTemplate:=Nothing
        .Borders.Enable = False
        .LanguageID = wdEnglishUS
        ' Font
        With .Font
            .Size = 9
            .name = StandardHeaderFontName
            .Bold = True
            .Italic = False
            .Color = RGB(77, 77, 77)
        End With
        ' Paragraph layout
        With .ParagraphFormat
            .LeftIndent = MillimetersToPoints(StandardLeftIndent)
            .FirstLineIndent = 0
            .LineSpacingRule = wdLineSpaceAtLeast
            .LineSpacing = 15
            .SpaceBefore = 15
            .SpaceAfter = 7
            .KeepWithNext = True
            .KeepTogether = True
            .OutlineLevel = wdOutlineLevel5
            .PageBreakBefore = False
        End With
    End With

End Sub

' Update Heading 7 style to our defaults
Private Sub FormatHeading7()
    Dim s As Style
    Set s = GetParaStyle("Heading 7")
    With s
        .BaseStyle = ActiveDocument.Styles("Heading 1")
        .LinkToListTemplate ListTemplate:=Nothing
        .Borders.Enable = False
        .LanguageID = wdEnglishUS
        ' Font
        With .Font
            .Size = 9
            .name = StandardHeaderFontName
            .Bold = True
            .Italic = False
            .Color = RGB(77, 77, 77)
        End With
        ' Paragraph layout
        With .ParagraphFormat
            .LeftIndent = MillimetersToPoints(StandardLeftIndent)
            .FirstLineIndent = 0
            .LineSpacingRule = wdLineSpaceAtLeast
            .LineSpacing = 15
            .SpaceBefore = 15
            .SpaceAfter = 7
            .KeepWithNext = True
            .KeepTogether = True
            .OutlineLevel = wdOutlineLevel5
            .PageBreakBefore = False
        End With
    End With

End Sub

' Update Heading 8 style to our defaults
Private Sub FormatHeading8()
    Dim s As Style
    Set s = GetParaStyle("Heading 8")
    With s
        .BaseStyle = ActiveDocument.Styles("Heading 1")
        .LinkToListTemplate ListTemplate:=Nothing
        .Borders.Enable = False
        .LanguageID = wdEnglishUS
        ' Font
        With .Font
            .Size = 9
            .name = StandardHeaderFontName
            .Bold = True
            .Italic = False
            .Color = RGB(77, 77, 77)
        End With
        ' Paragraph layout
        With .ParagraphFormat
            .LeftIndent = MillimetersToPoints(StandardLeftIndent)
            .FirstLineIndent = 0
            .LineSpacingRule = wdLineSpaceAtLeast
            .LineSpacing = 15
            .SpaceBefore = 15
            .SpaceAfter = 7
            .KeepWithNext = True
            .KeepTogether = True
            .OutlineLevel = wdOutlineLevel5
            .PageBreakBefore = False
        End With
    End With

End Sub

' Update Heading 9 style to our defaults
Private Sub FormatHeading9()
    Dim s As Style
    Set s = GetParaStyle("Heading 9")
    With s
        .BaseStyle = ActiveDocument.Styles("Heading 1")
        .LinkToListTemplate ListTemplate:=Nothing
        .Borders.Enable = False
        .LanguageID = wdEnglishUS
        ' Font
        With .Font
            .Size = 9
            .name = StandardHeaderFontName
            .Bold = True
            .Italic = False
            .Color = RGB(77, 77, 77)
        End With
        ' Paragraph layout
        With .ParagraphFormat
            .LeftIndent = MillimetersToPoints(StandardLeftIndent)
            .FirstLineIndent = 0
            .LineSpacingRule = wdLineSpaceAtLeast
            .LineSpacing = 15
            .SpaceBefore = 15
            .SpaceAfter = 7
            .KeepWithNext = True
            .KeepTogether = True
            .OutlineLevel = wdOutlineLevel5
            .PageBreakBefore = False
        End With
    End With

End Sub


' Set the document title format
Private Sub FormatTitle()
    Dim s As Style
    Set s = GetParaStyle("Title")
    With s
        .BaseStyle = ActiveDocument.Styles("Normal")
        .NextParagraphStyle = ActiveDocument.Styles("Subtitle")
        With .Font
            .Size = 24
            .name = StandardHeaderFontName
            .Bold = True
            .Italic = False
            .Color = RGB(102, 13, 20) ' Brown
        End With
        With .ParagraphFormat
            .LeftIndent = 0
            .FirstLineIndent = 0
            .LineSpacingRule = wdLineSpaceAtLeast
            .LineSpacing = 24
            .SpaceBefore = 12
            .SpaceAfter = 10
            .KeepWithNext = True
            .KeepTogether = True
            .PageBreakBefore = False
            .Alignment = wdAlignParagraphLeft
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
End Sub

' Set document subtitle format
Private Sub FormatSubTitle()
    Dim s As Style
    Set s = GetParaStyle("Subtitle")
    With s
        .BaseStyle = ActiveDocument.Styles("Normal")
        .NextParagraphStyle = ActiveDocument.Styles("Normal")
        With .Font
            .Size = 18
            .name = StandardHeaderFontName
            .Bold = True
            .Italic = False
            .Color = RGB(249, 152, 6) ' Orange
        End With
        With .ParagraphFormat
            .LeftIndent = 0
            .FirstLineIndent = 0
            .LineSpacingRule = wdLineSpaceAtLeast
            .LineSpacing = 18
            .SpaceBefore = 0
            .SpaceAfter = 21
            .KeepWithNext = True
            .KeepTogether = True
            .PageBreakBefore = False
            .Alignment = wdAlignParagraphLeft
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
End Sub


' Set program code style
' Do we want to set the backgroud/border for this?
Private Sub FormatCode()
    Dim s As Style
    Set s = GetParaStyle("Code")
    With s
        .BaseStyle = ActiveDocument.Styles("Normal")
        With .Font
            .Size = 9
            .name = "Courier New"
            .Bold = False
            .Italic = False
        End With
        With .ParagraphFormat
            .LeftIndent = Application.MillimetersToPoints(StandardLeftIndent + 5)
            .FirstLineIndent = 0
            .LineSpacingRule = wdLineSpaceSingle
            .SpaceBefore = 0
            .SpaceAfter = 0
            .KeepWithNext = False
            .KeepTogether = True
            .OutlineLevel = wdOutlineLevelBodyText
            .PageBreakBefore = False
        End With
    End With

End Sub



' Set block quote style
Private Sub FormatQuote()
    Dim s As Style
    Set s = GetParaStyle("Quote")
    With s
        .BaseStyle = ActiveDocument.Styles("Normal")
        With .Font
            .Size = 9
            .name = StandardBodyFontName
            .Bold = False
            .Italic = True
        End With
        With .ParagraphFormat
            .LeftIndent = Application.MillimetersToPoints(StandardLeftIndent + 3 + 8)
            .RightIndent = Application.MillimetersToPoints(StandardLeftIndent + 3 + 8)
            .FirstLineIndent = 0
            .LineSpacingRule = wdLineSpaceSingle
            .SpaceBefore = 16
            .SpaceAfter = 8
            .KeepWithNext = False
            .KeepTogether = True
            .OutlineLevel = wdOutlineLevelBodyText
            .PageBreakBefore = False
        End With
    End With

End Sub



' Style footnote text: smaller and italic
Private Sub FormatFootnoteText()
    Dim s As Style

    Set s = GetParaStyle("Footnote Text")
    With s
        .BaseStyle = ActiveDocument.Styles("Normal")
        .NextParagraphStyle = ActiveDocument.Styles("Footnote Text")
        With .Font
            .Size = 8.5
            .name = StandardBodyFontName
            .Bold = False
            .Italic = True
        End With
        With .ParagraphFormat
            '.LeftIndent = Application.MillimetersToPoints(StandardLeftIndent)
            '.FirstLineIndent = 0
            .LineSpacingRule = wdLineSpaceSingle
            .SpaceBefore = 2
            .SpaceAfter = 0
        End With
    End With
End Sub


' Set table cell body text
Private Sub FormatCellBody()
    Dim s As Style

    Set s = GetParaStyle("CellBody")
    With s
        .BaseStyle = ActiveDocument.Styles("Normal")
        .NextParagraphStyle = ActiveDocument.Styles("CellBody")
        With .Font
            .Size = 9
            .name = StandardBodyFontName
            .Bold = False
            .Italic = False
        End With
        With .ParagraphFormat
            .LeftIndent = 0
            .FirstLineIndent = 0
            .LineSpacingRule = wdLineSpaceSingle
            .SpaceBefore = 4
            .SpaceAfter = 4
        End With
    End With
End Sub

' Set table cell heading format
Private Sub FormatCellHeading()
    Dim s As Style

    Set s = GetParaStyle("CellHeading")
    With s
        .BaseStyle = ActiveDocument.Styles("Normal")
        .NextParagraphStyle = ActiveDocument.Styles("CellHeading")
        With .Font
            .Size = 9
            .name = StandardHeaderFontName
            .Color = RGB(77, 77, 77)
            .Bold = True
            .Italic = False
        End With
        With .ParagraphFormat
            .LeftIndent = 0
            .FirstLineIndent = 0
            .LineSpacingRule = wdLineSpaceSingle
            .SpaceBefore = 4
            .SpaceAfter = 4
        End With
    End With
End Sub

' Set page heading format
Private Sub FormatHeader()
    Dim s As Style

    Set s = GetParaStyle("Header")
    With s
        .BaseStyle = ActiveDocument.Styles("Normal")
        With .Font
            .Size = 8
            .name = StandardHeaderFontName
            .Color = RGB(77, 77, 77)
            .Bold = True
            .Italic = False
        End With
        With .ParagraphFormat
            .LeftIndent = 0
            .FirstLineIndent = 0
            .LineSpacingRule = wdLineSpaceSingle
            .SpaceBefore = 0
            .SpaceAfter = 0
            With .TabStops
                .ClearAll
                .Add Position:=MillimetersToPoints(82.5), Alignment:=wdAlignTabCenter
                .Add Position:=MillimetersToPoints(165), Alignment:=wdAlignTabRight
            End With
        End With
    End With
End Sub


' Set page footer format
Private Sub FormatFooter()
    Dim s As Style

    Set s = GetParaStyle("Footer")
    With s
        .BaseStyle = ActiveDocument.Styles("Normal")
        With .Font
            .Size = 8
            .name = StandardHeaderFontName
            .Color = RGB(77, 77, 77)
            .Bold = True
            .Italic = False
        End With
        With .ParagraphFormat
            .LeftIndent = 0
            .FirstLineIndent = 0
            .LineSpacingRule = wdLineSpaceSingle
            .SpaceBefore = 0
            .SpaceAfter = 0
            With .TabStops
                .ClearAll
                .Add Position:=MillimetersToPoints(82.5), Alignment:=wdAlignTabCenter
                .Add Position:=MillimetersToPoints(165), Alignment:=wdAlignTabRight
            End With
        End With
        .Borders.DistanceFromTop = 8
        With .Borders(wdBorderTop)
            .LineStyle = wdLineStyleSingle
            .Color = wdColorBlack
            .LineWidth = wdLineWidth050pt
        End With
    End With
End Sub


' Get a list template for table title numbering
Private Function GetTableListTemplate() As ListTemplate
    Dim lt As ListTemplate
    On Error Resume Next
    ' Get existing template
    Set lt = ActiveDocument.ListTemplates("TableListTemplate")
    If (lt Is Nothing) Then
        ' Or create if doesn't exist
        ' Note: name has to be string literal, not a variable (in Word 2000, at least)
        Set lt = ActiveDocument.ListTemplates.Add(False, "TableListTemplate")
    End If
    Set GetTableListTemplate = lt
End Function


' Get a list template for figure title numbering
Private Function GetFigureListTemplate() As ListTemplate
    Dim lt As ListTemplate
    On Error Resume Next
    ' Get existing template
    Set lt = ActiveDocument.ListTemplates("FigureListTemplate")
    If (lt Is Nothing) Then
        ' Or create if doesn't exist
        ' Note: name has to be string literal, not a variable (Word 2000)
        Set lt = ActiveDocument.ListTemplates.Add(False, "FigureListTemplate")
    End If
    Set GetFigureListTemplate = lt
End Function

' Set figure title format
Private Sub FormatFigure()
    Dim lt As ListTemplate

    Dim s As Style
    Set s = GetParaStyle("Figure")
    With s
        ' Get or create the numbering template
        If (.ListTemplate Is Nothing) Then
            Set lt = GetFigureListTemplate()
            lt.OutlineNumbered = False
            Call .LinkToListTemplate(lt)
        Else
            Set lt = .ListTemplate
        End If

        ' Set up the number format
        With lt.ListLevels(1)
            .NumberStyle = wdListNumberStyleArabic
            .NumberFormat = "Figure %1."
            .TrailingCharacter = wdTrailingSpace
            .LinkedStyle = "Figure"
        End With

        ' Set the paragraph style
        .BaseStyle = ActiveDocument.Styles("Normal")
        .NextParagraphStyle = ActiveDocument.Styles("Normal")
        With .Font
            .Size = 9
            .name = StandardHeaderFontName
            .Color = RGB(77, 77, 77)
            .Bold = True
            .Italic = False
        End With
        With .ParagraphFormat
            .LeftIndent = Application.MillimetersToPoints(StandardLeftIndent)
            .FirstLineIndent = 0
            .LineSpacingRule = wdLineSpaceSingle
            .SpaceBefore = 19
            .SpaceAfter = 14
        End With
    End With
End Sub


' Set table title format
Private Sub FormatTableTitle()
    Dim lt As ListTemplate
    Dim s As Style

    Set s = GetParaStyle("TableTitle")
    With s
        ' Set the paragraph style
        .BaseStyle = ActiveDocument.Styles("Normal")
        .NextParagraphStyle = ActiveDocument.Styles("Normal")
        Call .LinkToListTemplate(lt)
        With .Font
            .Size = 9
            .name = StandardHeaderFontName
            .Color = RGB(77, 77, 77)
            .Bold = True
            .Italic = False
        End With
        With .ParagraphFormat
            .LeftIndent = Application.MillimetersToPoints(StandardLeftIndent)
            .FirstLineIndent = 0
            .LineSpacingRule = wdLineSpaceSingle
            .SpaceBefore = 19
            .SpaceAfter = 14
        End With

        ' Get or create the numbering template
        If (.ListTemplate Is Nothing) Then
            Set lt = GetTableListTemplate()
        Else
            Set lt = .ListTemplate
        End If

        ' Set up the number format
        With lt.ListLevels(1)
            .NumberStyle = wdListNumberStyleArabic
            .NumberFormat = "Table %1."
            .TrailingCharacter = wdTrailingSpace
            '.LinkedStyle = "Table Title"
        End With
        Call .LinkToListTemplate(lt)

    End With
End Sub


' Get a style by name; create it if it doesn't already exist
Private Function GetParaStyle(name As String) As Style
    Dim s As Style
    Dim p As Paragraph
    ' Turn off error handling so we can tell when a style doesn't already exist
    On Error Resume Next
    Set s = ActiveDocument.Styles(name)
    If (s Is Nothing) Or (s = Empty) Or (s.Type <> wdStyleTypeParagraph) Then
        ' Create the style if it doesn't exist
        Set s = ActiveDocument.Styles.Add(name, wdStyleTypeParagraph)
    End If
    ' Create an example of the style: partly for reference, also
    ' some features don't get set if the style isn't used
    With ActiveDocument.Paragraphs
        Set p = .Add
        Set p = .Add
        p.Range.Text = name
        p.Style = s
        If name = "Quote" Then
            p.Range.InsertAfter (": Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.")
        ElseIf name = "Code" Then
            p.Range.InsertAfter (":" & vbNewLine & "Private Sub FormatCode()" & vbNewLine & "  Dim s As Style" & vbNewLine & "  Set s = GetParaStyle(""Code"")" & vbNewLine & "  With s")
        End If
    End With
    Set GetParaStyle = s
End Function


' Duplicate an existing style with a new name
' Not currently used
Private Sub CopyStyle(dest As String, src As String)
    Dim destStyle As Style
    Dim srcStyle As Style
    Set srcStyle = ActiveDocument.Styles(src)
    Set destStyle = GetParaStyle(dest)

    destStyle.BaseStyle = srcStyle
    destStyle.NextParagraphStyle = srcStyle.NextParagraphStyle

    ' Newly created style so leave verything to be inherited
    If (Not destStyle.BuiltIn) Then Exit Sub

    ' Builtin style so set as much as possible from srcStyle.style
    With destStyle
        .BaseStyle = srcStyle.BaseStyle
        .NextParagraphStyle = srcStyle.NextParagraphStyle
        .Font = srcStyle.Font
        .ParagraphFormat = srcStyle.ParagraphFormat.Duplicate
        .Borders = srcStyle.Borders
        If Not .ListTemplate Is Nothing Then
            Call .LinkToListTemplate(.ListTemplate)
        End If
        With .Frame
            If srcStyle.Frame.Height > 0 Then Height = srcStyle.Frame.Height
            .HeightRule = srcStyle.Frame.HeightRule
            .HorizontalDistanceFromText = srcStyle.Frame.HorizontalDistanceFromText
            .HorizontalPosition = srcStyle.Frame.HorizontalPosition
            .LockAnchor = srcStyle.Frame.LockAnchor
            .RelativeHorizontalPosition = srcStyle.Frame.RelativeHorizontalPosition
            .RelativeVerticalPosition = srcStyle.Frame.RelativeVerticalPosition
            '.Shading = srcStyle.Frame.Shading
            .TextWrap = srcStyle.Frame.TextWrap
            .VerticalDistanceFromText = srcStyle.Frame.VerticalDistanceFromText
            .VerticalPosition = srcStyle.Frame.VerticalPosition
            If srcStyle.Frame.Width > 0 Then .Width = srcStyle.Frame.Width
            .WidthRule = srcStyle.Frame.WidthRule
        End With
        With .Shading
            .BackgroundPatternColor = srcStyle.Shading.BackgroundPatternColor
            .BackgroundPatternColorIndex = srcStyle.Shading.BackgroundPatternColorIndex
            .ForegroundPatternColor = srcStyle.Shading.ForegroundPatternColor
            .ForegroundPatternColorIndex = srcStyle.Shading.ForegroundPatternColorIndex
            .Texture = srcStyle.Shading.Texture
        End With
    End With
End Sub

