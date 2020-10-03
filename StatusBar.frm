'Inspired from book: "Excel 2016 Power Programming with VBA", John Wiley & Sons
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StatusBar 
   Caption         =   "Loading"
   ClientHeight    =   1692
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4512
   OleObjectBlob   =   "StatusBar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "StatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub UserForm_Initialize()
With Me
    .lblProgress.BackColor = vbGreen
    .lblProgress.Width = 1

End With
End Sub
Public Sub SetDescription(Description As String)
    Me.ProcessName.Caption = Description

End Sub

Public Sub UpdateProgress(PctDone As Double)
    Dim PctInt As Integer
    PctInt = PctDone
    With Me
        .frmProgress.Caption = PctInt & "%"
        .lblProgress.Width = PctInt * 2
        .Repaint
    End With
End Sub
