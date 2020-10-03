VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Description 
   Caption         =   "Create description"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   14004
   OleObjectBlob   =   "Description.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Description"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub OkButton_Click()
    On Error Resume Next
    StartDate = CDate(Me.TextBox5.text)
    If Err Then
    MsgBox "Wrong Date"
    Exit Sub
    End If
    
    Dim tests As Boolean ' test about correct data
    Call DataTest.Test(tests)
    
    If tests Then
        If Me.TextBox1.text <> "This text will appear before <date>" Then
        Text1 = Me.TextBox1.text
        Else
        Text1 = ""
        End If
        Text2 = Me.TextBox2.text
        Text3 = Me.TextBox3.text
        Text4 = Me.TextBox4.text
        Call START.Generator
    Else
        MsgBox "Check information in columns"
    End If
    
End Sub

Private Sub UserForm_Initialize()
'fill the data textbox with current date
Dim LDate As String
LDate = Date
Me.TextBox5.text = LDate

End Sub
