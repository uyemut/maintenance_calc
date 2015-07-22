Attribute VB_Name = "Module1"
Sub ScanDates()
'************************************************************************************
'***                                                                              ***
'***  Program - ScanDates()                                                       ***
'***  Author - Thomas Uyemura                                                     ***
'***             Calendar Coordinator                                             ***
'***  Written - Jan 9th, 2015                                                     ***
'***  This script will format the Data Spreadsheet so that each date consists of  ***
'***           eight lines and NO MORE than EIGHT lines .                         ***
'***                                                                              ***
'************************************************************************************
Dim iRow As Long
Dim iRowStart As Long
Dim iRowPrev As Long

Dim iRowMaxOneDate As Long
Dim iRowCnt As Long
Dim iRowMaxLimit As Long
Dim iRowFirst As Long

Dim wkInsertCnt As Long
Dim wkRowDiff As Long

Dim ws As Worksheet
Dim wkDate As Date
Dim wkDatePrev As Date

Dim Rng As Range
Dim wsHardSearch As Boolean
Dim wkDebugg As Boolean

wkHardSearch = True
iRowFirst = 2
iRowStart = iRowFirst
iRow = iRowFirst
iRowPrev = iRowFirst
iRowMaxOneDate = 8
iRowCnt = 8
iRowMaxLimit = 3333

Set ws = Worksheets("Data")
wkDatePrev = ws.Cells(iRow, 6)

Set Rng = ws.Columns("F:F").Find(What:=wkDatePrev + 1, After:=ws.Range("F" & iRowStart), _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlNext, LookIn:=xlFormulas)

iRowStart = Rng.Row
iRow = Rng.Row
wkDatePrev = ws.Cells(iRow, 6)

wkRowDiff = Rng.Row - iRowPrev
wkInsertCnt = iRowMaxOneDate - wkRowDiff
If wkRowDiff < iRowMaxOneDate Then
   ws.Range(Cells(iRow, 1), Cells(iRow, 1)).EntireRow.Resize(wkInsertCnt).Insert
End If

iRowStart = iRow + wkInsertCnt
iRowPrev = iRowStart

'Add Columns so that all dates have 8 lines ( rows ) !

    Do While ws.Cells(iRowStart, 6) <> ""
        ws.Range(Cells(iRow, 1), Cells(iRow, 1)).Select
        
        On Error Resume Next
        Set Rng = ws.Columns("F:F").Find(What:=wkDatePrev + 1, After:=ws.Range("F" & iRowStart), _
                  SearchOrder:=xlByRows, _
                  SearchDirection:=xlNext, LookIn:=xlFormulas)
        If Err.Number <> 0 Then
           MsgBox "Very Very Strange, No Null Account Numbers found after Micco " & _
                  " Lot! Investigation Needed."
           MsgBox " This is a very fatal error, abort mission ! "
           iRow = 1
           Exit Sub
        End If

        iRow = Rng.Row
        wkDatePrev = ws.Cells(iRow, 6)

        wkRowDiff = Rng.Row - iRowPrev
        wkInsertCnt = iRowMaxOneDate - wkRowDiff
        If wkRowDiff < iRowMaxOneDate Then
           ws.Range(Cells(iRow, 1), Cells(iRow, 1)).EntireRow.Resize(wkInsertCnt).Insert
        End If
        
        iRowStart = iRow + wkInsertCnt
        iRowPrev = iRowStart
        ' Lets put in a Range Check, so we dont go for an infinite loop !!!
        
        If iRow > iRowMaxLimit Then
            Exit Sub
        End If
    Loop


End Sub
