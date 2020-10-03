Attribute VB_Name = "DataTest"
Option Explicit
'tested for filling up all data correct
Sub Test(Results As Boolean)
Results = True
Dim cont As Control
Dim text As String
For Each cont In MAIN.Controls
    If TypeName(cont) = "TextBox" Then
        text = cont.text
        If Not IsNumeric(text) Then Results = False
    End If

Next cont
If MAIN.SelectedFileLabel = "Nothing" Then Results = False
If MAIN.SelectedWordLabel = "Nothing" Then Results = False
End Sub
