Attribute VB_Name = "Variables"
Option Explicit
Public Text1 As String
Public Text2 As String
Public Text3 As String
Public Text4 As String
Public StartDate As Date
Type LandPlot
    Number As String
    Hour As String
    Date As String
End Type

Type Person
    Name As String
    Street As String
    Postcode As String
    ID As String
    LandPlots() As String
    Description As String
End Type
