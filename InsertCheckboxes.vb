'
' Insert checkboxes in a number of cells and link the cell contents to its checkbox
'
Sub AddCheckBoxes()
    Dim c As Range, myRange As Range
    
    ' Ignore errors
    On Error Resume Next
    ' Turn off screen updating
    Application.ScreenUpdating = False
    
    Set myRange = Selection
    For Each c In myRange.Cells
        Dim checked As Boolean
        Dim box As CheckBox
        
        ' Set checked state based on some existing data
        checked = c.Offset(0, 1).Value <> "N"
        ' Insert the checkbox
        ActiveSheet.CheckBoxes.Add(c.Left, c.Top, c.Width, c.Height).Select
        Set box = Selection
        With box
            ' Link the checkbox to the corresponding cell
            .LinkedCell = c.Address
            ' Remove the CheckBox label
            .Characters.Text = ""
            ' Set name of control to the cell name
            .Name = c.Address
            ' Check/uncheck as appropriate
            .Value = checked
        End With
        
       ' Set the text colour to the background so we can't see the TRUE/FALSE
        c.Font.Color = c.Interior.Color
    Next
    
    ' Restore original selection
    myRange.Select
    
    ' Restore screen updating
    Application.ScreenUpdating = True
End Sub


' Delete checkboxes
Sub RemoveCheckboxes()
    Dim c As CheckBox
    Dim n, i As Integer
	' Get number in collection
    n = ActiveSheet.CheckBoxes.Count
	' For each doesn't work for some reason
    For i = n To 1 Step -1 ' have to count down 
        ActiveSheet.CheckBoxes(i).Delete
    Next i
End Sub
