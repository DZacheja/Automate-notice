Attribute VB_Name = "WordTest"
Option Explicit

Sub BookmarksTest(Results As Boolean)
    Results = True

    Dim wdDoc As Document
    Set wdDoc = ThisDocument
    Selection.InsertFile FileName:=MAIN.SelectedWordLabel.Caption

    Dim TestRange As Word.Range

    On Error Resume Next

    Set TestRange = wdDoc.Bookmarks("contents").Range
    If Err Then Results = False

    Set TestRange = wdDoc.Bookmarks("ID").Range
    If Err Then Results = False

    Set TestRange = wdDoc.Bookmarks("Name").Range
    If Err Then Results = False

    Set TestRange = wdDoc.Bookmarks("Postcode").Range
    If Err Then Results = False

    Set TestRange = wdDoc.Bookmarks("Street").Range
    If Err Then Results = False

    wdDoc.Content.Delete
    On Error GoTo 0
End Sub



