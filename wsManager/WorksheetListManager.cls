VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetListManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Function returnCurrentWorksheets() As String()
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    Dim wsCount As Integer
    Let wsCount = wb.Worksheets.Count()
    
    Dim wsArr() As String
    ReDim wsArr(1 To wsCount)
    
    Dim i As Integer
    Dim wsCurr As Worksheet
    For i = 1 To wb.Worksheets.Count
        Set wsCurr = wb.Worksheets(i)
        wsArr(i) = wsCurr.Name
    Next
    
    returnCurrentWorksheets = wsArr()
End Function
