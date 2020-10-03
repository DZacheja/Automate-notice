VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MAIN 
   Caption         =   "GENERATOR"
   ClientHeight    =   6756
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9108.001
   OleObjectBlob   =   "MAIN.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ExcelSelect_Click()
'select exel and fill comboboxes
    Dim ExApp As Excel.Application
    Dim wb As Workbook
    Dim FileName As String 'fileSelected
    Dim Filter As String
    Dim Title As String
    Dim ws As Excel.Worksheet
    Dim FD As FileDialog
    Filter = "Excel file (*.xl**),*.*"
    Title = "Select Excel file"
    Set ExApp = New Excel.Application
    Set FD = Application.FileDialog(msoFileDialogFilePicker)
    With FD
        .Filters.Clear
        .Filters.Add "Excel file", "*.xl**", 1
        .Title = Title
        .AllowMultiSelect = False
        If .Show = -1 Then FileName = .SelectedItems(1)
        
    End With
    
    If Not FileName = "" Then
   'fill the label contains excel path
    Set wb = Excel.Workbooks.Open(FileName, , True)
    Me.SelectedFileLabel.Caption = FileName
    Me.SelectedFileLabel.ForeColor = &HFF00&
    Else
    MsgBox "Nothing selected!"
    ExApp.Quit
    Exit Sub
    End If
    'fill up ComboBoxe's
    For Each ws In wb.Worksheets
        Me.SheetPeopleDataComboBox.AddItem ws.Name
        Me.SheetLandPlotsComboBox.AddItem ws.Name
    Next ws
    ExApp.Quit
    Set wb = Nothing
    
    'select first item in ComboBox
    Me.SheetLandPlotsComboBox.ListIndex = 0
    Me.SheetPeopleDataComboBox.ListIndex = 0
    
End Sub

Private Sub GENERATE_Click()
    MAIN.Hide
    Description.Show
    
End Sub

Private Sub WordSelect_Click()
      'select the orginal word witch contains all bookmarks
    Dim wdApp As Word.Application
    Dim selected As Variant
    Dim FileName As String 'fileSelected
    Dim FileDg As FileDialog
    Dim CorrectFile As Boolean
    Set FileDg = Application.FileDialog(msoFileDialogFilePicker)

    With FileDg
    .Filters.Clear
    .Filters.Add "Word document", "*.do**", 1
    .Title = "Select word file"
    .AllowMultiSelect = False
    If FileDg.Show = -1 Then FileName = FileDg.SelectedItems(1)
    End With
    
    If FileName = "" Then
    MsgBox "Nothing selected"
    Exit Sub
    End If
         'fill the label contains word path
    FileName = FileDg.SelectedItems(1)
    Me.SelectedWordLabel.Caption = FileName
    Me.SelectedWordLabel.ForeColor = &HFF00&
    Call WordTest.BookmarksTest(CorrectFile)
    If Not CorrectFile Then
        MsgBox "Check bookmarks in word file!"
        Me.SelectedWordLabel.Caption = "Nothing"
    Else
    Me.SelectedWordLabel.ForeColor = &HFF00&
    End If
End Sub
