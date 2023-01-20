VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UpdateBeforeSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents App As Word.Application
Attribute App.VB_VarHelpID = -1

Private Sub App_DocumentBeforeSave(ByVal Doc As Document, SaveAsUI As Boolean, Cancel As Boolean)

    Call UpdateScores

End Sub

Private Sub UpdateScores()
    Dim f As Field
    
    For Each f In ActiveDocument.Fields
        f.Update
    Next f
End Sub
