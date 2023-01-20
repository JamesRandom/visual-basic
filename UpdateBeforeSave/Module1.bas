Attribute VB_Name = "Module1"
Option Explicit
Dim X As New UpdateBeforeSave
Public Sub Register_Event_Handler()
    Set X.App = Word.Application
End Sub



