Attribute VB_Name = "START"


Sub Generator()
Description.Hide
Application.ScreenUpdating = False
On Error GoTo ActivityError
    Dim CurrentActivity As String
    Dim CopyWord As Word.Application
    Dim WdRange As Word.Range
    Dim ThisWord As Document
    
    Dim PersonRange As Excel.Range
    Dim LandPlotsRange As Excel.Range
    Dim ws_DATA As Excel.Worksheet
    Dim ws_TIMETABLE As Excel.Worksheet
    
    Dim OrginalFile As String
    Dim CopyFile As String
    Dim objExcel As Object
    Dim exWB As Workbook
    Dim Persons() As Person
    Dim LandPlots() As LandPlot
    Dim i As Long
    Dim prc As Double
    Dim CName As Integer: CName = MAIN.NameBox.text
    Dim CStreet As Integer: CStreet = MAIN.StreetBox.text
    Dim CPostcode As Integer: CPostcode = MAIN.PostcodeBox.text
    Dim CID As Integer: CID = MAIN.IdBox.text
    Dim CLand As Integer: CLand = MAIN.LandplotsBox
    Dim CNumber As Integer: CNumber = MAIN.NumberBox
    Dim ClDate As Integer: ClDate = MAIN.DateBox.text
    Dim ClTime As Integer: ClTime = MAIN.TimeBox.text
    
    Set ThisWord = ThisDocument
    OrginalFile = MAIN.SelectedWordLabel.Caption 'orginal Filename
    
    CurrentActivity = "Opening excel file"
    Set objExcel = CreateObject("Excel.Application")
    Set exWB = objExcel.Workbooks.Open(MAIN.SelectedFileLabel.Caption) 'Excel open
    
    CurrentActivity = "Setup range in selected Sheets"
    Set ws_DATA = exWB.Sheets(MAIN.SheetPeopleDataComboBox.text) 'Sheet with people data
    Set ws_TIMETABLE = exWB.Sheets(MAIN.SheetLandPlotsComboBox.text) 'Sheet with LandPlots data
    ThisWord.Activate
    
    Set LandPlotsRange = ws_TIMETABLE.Range("A1").CurrentRegion
    
'status bar setup
    StatusBar.Show vbModeless
    StatusBar.SetDescription "Load informations about Land Plots..."
    
CurrentActivity = "Load information about Land Plots..."
    
'Get the information about Land plots
    For i = 2 To LandPlotsRange.Rows.Count
        
            ReDim Preserve LandPlots(i - 2) 'resize array of land plots
            LandPlots(i - 2).Number = LandPlotsRange.Cells(i, CNumber) ' Land plot number
            LandPlots(i - 2).Date = CStr(LandPlotsRange.Cells(i, ClDate)) ' Land plot date
            LandPlots(i - 2).Hour = LandPlotsRange.Cells(i, ClTime) 'land plot Time
            prc = (i / LandPlotsRange.Rows.Count) * 100
            StatusBar.UpdateProgress (prc)
    Next i
    StatusBar.Hide
    StatusBar.Show vbModeless
    StatusBar.SetDescription "Load informations about person..."
    
'People's data information
    Set PersonRange = ws_DATA.Range("A1").CurrentRegion
    For i = 2 To PersonRange.Rows.Count
        ReDim Preserve Persons(i - 2)
        Persons(i - 2).Name = PersonRange.Cells(i, CName) 'person name
        Persons(i - 2).Street = PersonRange.Cells(i, CStreet) 'person street
        Persons(i - 2).Postcode = PersonRange.Cells(i, CPostcode) 'person postcode
        Persons(i - 2).ID = PersonRange.Cells(i, CID) 'person id
        Persons(i - 2).LandPlots() = Split(PersonRange.Cells(i, CLand), ",") ' person LandPlot
        'call function for making a description
        Persons(i - 2).Description = _
        StringForPlot(Persons(i - 2).LandPlots, LandPlots())
        prc = (i / PersonRange.Rows.Count) * 100
        StatusBar.UpdateProgress (prc)
   Next i
StatusBar.Hide
objExcel.Quit
Set exWB = Nothing
'clear curr document
ThisWord.Content.Delete

    StatusBar.Show vbModeless
    StatusBar.SetDescription "Filling up the file..."

    For i = 0 To UBound(Persons()) 'create documents for all peoples
'insert new page and orginal file to current file
    If i <> 0 Then
    Selection.EndKey Unit:=wdStory
    Selection.InsertBreak Type:=wdPageBreak
    Selection.EndKey Unit:=wdStory
    End If
    Selection.InsertFile FileName:=OrginalFile

'insert name
    Set WdRange = ThisWord.Bookmarks("Name").Range
    WdRange.InsertAfter text:=Persons(i).Name
'insert sreet
    Set WdRange = ThisWord.Bookmarks("Street").Range
    WdRange.InsertAfter Persons(i).Street
'insert postcode
    Set WdRange = ThisWord.Bookmarks("Postcode").Range
    WdRange.InsertAfter Persons(i).Postcode
'insert ID
    Set WdRange = ThisWord.Bookmarks("ID").Range
    WdRange.InsertAfter Persons(i).ID
'insert decription
    Set WdRange = ThisWord.Bookmarks("contents").Range
    WdRange.InsertAfter Persons(i).Description
'remove all bookmarks
    ThisWord.Bookmarks("Name").Delete
    ThisWord.Bookmarks("Street").Delete
    ThisWord.Bookmarks("Postcode").Delete
    ThisWord.Bookmarks("ID").Delete
    ThisWord.Bookmarks("contents").Delete
'calulate precents
    prc = (i / UBound(Persons())) * 100
    StatusBar.UpdateProgress (prc)
    
    Next i
StatusBar.Hide

Application.ScreenUpdating = True
Exit Sub
ActivityError:
    On Error Resume Next
        objExcel.Quit
    On Error GoTo 0
    MsgBox "Error in: " & CurrentActivity
End Sub
Function StringForPlot(plot() As String, ByRef LandPlots() As LandPlot) As String
'Description
    Dim curr_data As Date
    StringForPlot = ""
    Dim DaysTable() As String
    Dim TimeTable() As String
    Dim idx As Integer
    For i = 0 To UBound(plot())
        
        idx = index(LandPlots, plot(i))
        DaysTable() = Split(LandPlots(idx).Date, ",")
        TimeTable() = Split(LandPlots(idx).Hour, ",")
        If UBound(DaysTable()) > 0 Then
            For j = 0 To UBound(DaysTable())
            curr_data = DateAdd("d", CInt(DaysTable(j)), StartDate)
                StringForPlot = StringForPlots & _
        Text1 & " " & curr_data & " " & Text2 & " " & TimeTable(j) & ":00 " & Text3 & " " & plot(i) & " " & Text4 & Chr(10)
            Next j
        Else
        curr_data = DateAdd("d", CInt(LandPlots(i).Date), StartDate)
        StringForPlot = StringForPlot & _
        Text1 & " " & curr_data & " " & Text2 & " " & TimeTable(0) & ":00 " & Text3 & " " & plot(i) & " " & Text4 & Chr(10)
        End If
    Next i
    StringForPlot = Mid(StringForPlot, 1, Len(StringForPlot) - 1)
End Function
Function index(tablica() As LandPlot, szukana As String) As Integer
    Dim FoundTest As Boolean: FoundTest = False
    'Found the index in land plots array
    szukana = Trim(szukana)
    For i = LBound(tablica()) To UBound(tablica())
        If Trim(tablica(i).Number) = szukana Then
            index = i
            FoundTest = True
            Exit For
        End If
    Next i
    If FoundTest = False Then
    MsgBox "Not Found!: " & LandPlot
    End
    End If
End Function

