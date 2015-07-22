Attribute VB_Name = "Module2"
Sub fmtOne()
Attribute fmtOne.VB_ProcData.VB_Invoke_Func = " \n14"
'************************************************************************************
'***                                                                              ***
'***  Program - fmtOne()                                                          ***
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
Dim ws2 As Worksheet

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

Set ws = Worksheets("Calendar")
Set ws2 = Worksheets("Data")

    ws.Range("Q2:Q250").Select
    Selection.Copy
    ws2.Select
    ws2.Range("D2").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 5
    ActiveWindow.ScrollRow = 8
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 12
    ActiveWindow.ScrollRow = 13
    ActiveWindow.ScrollRow = 15
    ActiveWindow.ScrollRow = 16
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 20
    ActiveWindow.ScrollRow = 23
    ActiveWindow.ScrollRow = 26
    ActiveWindow.ScrollRow = 28
    ActiveWindow.ScrollRow = 31
    ActiveWindow.ScrollRow = 33
    ActiveWindow.ScrollRow = 35
    ActiveWindow.ScrollRow = 37
    ActiveWindow.ScrollRow = 39
    ActiveWindow.ScrollRow = 40
    ActiveWindow.ScrollRow = 43
    ActiveWindow.ScrollRow = 46
    ActiveWindow.ScrollRow = 48
    ActiveWindow.ScrollRow = 49
    ActiveWindow.ScrollRow = 51
    ActiveWindow.ScrollRow = 52
    ActiveWindow.ScrollRow = 55
    ActiveWindow.ScrollRow = 57
    ActiveWindow.ScrollRow = 59
    ActiveWindow.ScrollRow = 61
    ActiveWindow.ScrollRow = 63
    ActiveWindow.ScrollRow = 65
    ActiveWindow.ScrollRow = 66
    ActiveWindow.ScrollRow = 68
    ActiveWindow.ScrollRow = 70
    ActiveWindow.ScrollRow = 72
    ActiveWindow.ScrollRow = 75
    ActiveWindow.ScrollRow = 77
    ActiveWindow.ScrollRow = 79
    ActiveWindow.ScrollRow = 80
    ActiveWindow.ScrollRow = 83
    ActiveWindow.ScrollRow = 85
    ActiveWindow.ScrollRow = 86
    ActiveWindow.ScrollRow = 87
    ActiveWindow.ScrollRow = 88
    ActiveWindow.ScrollRow = 89
    ActiveWindow.ScrollRow = 91
    ActiveWindow.ScrollRow = 92
    ActiveWindow.ScrollRow = 94
    ActiveWindow.ScrollRow = 96
    ActiveWindow.ScrollRow = 97
    ActiveWindow.ScrollRow = 98
    ActiveWindow.ScrollRow = 100
    ActiveWindow.ScrollRow = 101
    ActiveWindow.ScrollRow = 102
    ActiveWindow.ScrollRow = 103
    ActiveWindow.ScrollRow = 105
    ActiveWindow.ScrollRow = 107
    ActiveWindow.ScrollRow = 108
    ActiveWindow.ScrollRow = 109
    ActiveWindow.ScrollRow = 110
    ActiveWindow.ScrollRow = 111
    ActiveWindow.ScrollRow = 112
    ActiveWindow.ScrollRow = 113
    ActiveWindow.ScrollRow = 114
    ActiveWindow.ScrollRow = 115
    ActiveWindow.ScrollRow = 118
    ActiveWindow.ScrollRow = 120
    ActiveWindow.ScrollRow = 122
    ActiveWindow.ScrollRow = 123
    ActiveWindow.ScrollRow = 126
    ActiveWindow.ScrollRow = 127
    ActiveWindow.ScrollRow = 128
    ActiveWindow.ScrollRow = 130
    ActiveWindow.ScrollRow = 131
    ActiveWindow.ScrollRow = 132
    ActiveWindow.ScrollRow = 133
    ActiveWindow.ScrollRow = 134
    ActiveWindow.ScrollRow = 135
    ActiveWindow.ScrollRow = 138
    ActiveWindow.ScrollRow = 141
    ActiveWindow.ScrollRow = 145
    ActiveWindow.ScrollRow = 146
    ActiveWindow.ScrollRow = 149
    ActiveWindow.ScrollRow = 150
    ActiveWindow.ScrollRow = 151
    ActiveWindow.ScrollRow = 152
    ActiveWindow.ScrollRow = 153
    ActiveWindow.ScrollRow = 154
    ActiveWindow.ScrollRow = 156
    ActiveWindow.ScrollRow = 157
    ActiveWindow.ScrollRow = 158
    ActiveWindow.ScrollRow = 159
    ActiveWindow.ScrollRow = 160
    ActiveWindow.ScrollRow = 161
    ActiveWindow.ScrollRow = 162
    ActiveWindow.ScrollRow = 163
    ActiveWindow.ScrollRow = 166
    ActiveWindow.ScrollRow = 170
    ActiveWindow.ScrollRow = 172
    ActiveWindow.ScrollRow = 177
    ActiveWindow.ScrollRow = 178
    ActiveWindow.ScrollRow = 179
    ActiveWindow.ScrollRow = 182
    ActiveWindow.ScrollRow = 184
    ActiveWindow.ScrollRow = 186
    ActiveWindow.ScrollRow = 188
    ActiveWindow.ScrollRow = 190
    ActiveWindow.ScrollRow = 192
    ActiveWindow.ScrollRow = 193
    ActiveWindow.ScrollRow = 194
    ActiveWindow.ScrollRow = 195
    ActiveWindow.ScrollRow = 196
    ActiveWindow.ScrollRow = 198
    ActiveWindow.ScrollRow = 200
    ActiveWindow.ScrollRow = 201
    ActiveWindow.ScrollRow = 202
    ActiveWindow.ScrollRow = 201
    ActiveWindow.ScrollRow = 199
    ActiveWindow.ScrollRow = 198
    ActiveWindow.ScrollRow = 196
    ActiveWindow.ScrollRow = 195
    ActiveWindow.ScrollRow = 194
    ActiveWindow.ScrollRow = 193
    ActiveWindow.ScrollRow = 192
    ActiveWindow.ScrollRow = 191
    ActiveWindow.ScrollRow = 190
    ActiveWindow.ScrollRow = 189
    ActiveWindow.ScrollRow = 188
    ActiveWindow.ScrollRow = 187
    ActiveWindow.ScrollRow = 186
    ActiveWindow.ScrollRow = 185
    ActiveWindow.ScrollRow = 183
    ActiveWindow.ScrollRow = 181
    ActiveWindow.ScrollRow = 180
    ActiveWindow.ScrollRow = 179
    ActiveWindow.ScrollRow = 177
    ActiveWindow.ScrollRow = 175
    ActiveWindow.ScrollRow = 173
    ActiveWindow.ScrollRow = 172
    ActiveWindow.ScrollRow = 169
    ActiveWindow.ScrollRow = 166
    ActiveWindow.ScrollRow = 164
    ActiveWindow.ScrollRow = 162
    ActiveWindow.ScrollRow = 159
    ActiveWindow.ScrollRow = 157
    ActiveWindow.ScrollRow = 153
    ActiveWindow.ScrollRow = 152
    ActiveWindow.ScrollRow = 148
    ActiveWindow.ScrollRow = 145
    ActiveWindow.ScrollRow = 141
    ActiveWindow.ScrollRow = 139
    ActiveWindow.ScrollRow = 136
    ActiveWindow.ScrollRow = 134
    ActiveWindow.ScrollRow = 130
    ActiveWindow.ScrollRow = 126
    ActiveWindow.ScrollRow = 124
    ActiveWindow.ScrollRow = 120
    ActiveWindow.ScrollRow = 115
    ActiveWindow.ScrollRow = 112
    ActiveWindow.ScrollRow = 109
    ActiveWindow.ScrollRow = 102
    ActiveWindow.ScrollRow = 99
    ActiveWindow.ScrollRow = 98
    ActiveWindow.ScrollRow = 96
    ActiveWindow.ScrollRow = 92
    ActiveWindow.ScrollRow = 90
    ActiveWindow.ScrollRow = 89
    ActiveWindow.ScrollRow = 85
    ActiveWindow.ScrollRow = 83
    ActiveWindow.ScrollRow = 81
    ActiveWindow.ScrollRow = 77
    ActiveWindow.ScrollRow = 75
    ActiveWindow.ScrollRow = 69
    ActiveWindow.ScrollRow = 66
    ActiveWindow.ScrollRow = 63
    ActiveWindow.ScrollRow = 57
    ActiveWindow.ScrollRow = 54
    ActiveWindow.ScrollRow = 51
    ActiveWindow.ScrollRow = 48
    ActiveWindow.ScrollRow = 44
    ActiveWindow.ScrollRow = 42
    ActiveWindow.ScrollRow = 38
    ActiveWindow.ScrollRow = 35
    ActiveWindow.ScrollRow = 33
    ActiveWindow.ScrollRow = 31
    ActiveWindow.ScrollRow = 28
    ActiveWindow.ScrollRow = 26
    ActiveWindow.ScrollRow = 24
    ActiveWindow.ScrollRow = 21
    ActiveWindow.ScrollRow = 19
    ActiveWindow.ScrollRow = 16
    ActiveWindow.ScrollRow = 13
    ActiveWindow.ScrollRow = 12
    ActiveWindow.ScrollRow = 9
    ActiveWindow.ScrollRow = 8
    ActiveWindow.ScrollRow = 5
    ActiveWindow.ScrollRow = 4
    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 1
    Range("E2").Select
    Sheets("Calendar").Select
    Range("A1").Select
    Application.CutCopyMode = False
    Range("A2:A250").Select
    Selection.Copy
    Sheets("Data").Select
    Range("E2").Select
    ActiveSheet.Paste
    Range("F2").Select
    Sheets("Calendar").Select
    Range("A1").Select
    Application.CutCopyMode = False
    Range("B2:E250").Select
    Selection.Copy
    Sheets("Data").Select
    Range("F2").Select
    ActiveSheet.Paste
    Range("A1").Select
    Sheets("Calendar").Select
    Range("A3").Select
    Application.CutCopyMode = False
    Range("P2:P250").Select
    Selection.Copy
    Sheets("Data").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    Range("L2").Select
    ActiveSheet.Paste
    Range("A1").Select
    Sheets("Calendar").Select
    Range("A1").Select
    Application.CutCopyMode = False
    Range("A2").Select
    Sheets("Data").Select
    Range("B2").Select
    Selection.Copy
    Range("B3:B250").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("J2:K2").Select
    Selection.Copy
    Range("J3:K250").Select
    ActiveWindow.ScrollRow = 217
    ActiveWindow.ScrollRow = 215
    ActiveWindow.ScrollRow = 213
    ActiveWindow.ScrollRow = 209
    ActiveWindow.ScrollRow = 207
    ActiveWindow.ScrollRow = 202
    ActiveWindow.ScrollRow = 200
    ActiveWindow.ScrollRow = 196
    ActiveWindow.ScrollRow = 192
    ActiveWindow.ScrollRow = 190
    ActiveWindow.ScrollRow = 187
    ActiveWindow.ScrollRow = 179
    ActiveWindow.ScrollRow = 175
    ActiveWindow.ScrollRow = 168
    ActiveWindow.ScrollRow = 164
    ActiveWindow.ScrollRow = 161
    ActiveWindow.ScrollRow = 156
    ActiveWindow.ScrollRow = 152
    ActiveWindow.ScrollRow = 149
    ActiveWindow.ScrollRow = 141
    ActiveWindow.ScrollRow = 137
    ActiveWindow.ScrollRow = 130
    ActiveWindow.ScrollRow = 121
    ActiveWindow.ScrollRow = 107
    ActiveWindow.ScrollRow = 99
    ActiveWindow.ScrollRow = 85
    ActiveWindow.ScrollRow = 77
    ActiveWindow.ScrollRow = 70
    ActiveWindow.ScrollRow = 58
    ActiveWindow.ScrollRow = 53
    ActiveWindow.ScrollRow = 47
    ActiveWindow.ScrollRow = 43
    ActiveWindow.ScrollRow = 38
    ActiveWindow.ScrollRow = 33
    ActiveWindow.ScrollRow = 28
    ActiveWindow.ScrollRow = 25
    ActiveWindow.ScrollRow = 23
    ActiveWindow.ScrollRow = 21
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 17
    ActiveWindow.ScrollRow = 15
    ActiveWindow.ScrollRow = 13
    ActiveWindow.ScrollRow = 12
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 8
    ActiveWindow.ScrollRow = 7
    ActiveWindow.ScrollRow = 6
    ActiveWindow.ScrollRow = 5
    ActiveWindow.ScrollRow = 4
    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 1
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("B2:E2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("B2:O250").Select
    ActiveWindow.ScrollRow = 209
    ActiveWindow.ScrollRow = 210
    ActiveWindow.ScrollRow = 209
    ActiveWindow.ScrollRow = 208
    ActiveWindow.ScrollRow = 199
    ActiveWindow.ScrollRow = 194
    ActiveWindow.ScrollRow = 185
    ActiveWindow.ScrollRow = 178
    ActiveWindow.ScrollRow = 175
    ActiveWindow.ScrollRow = 171
    ActiveWindow.ScrollRow = 164
    ActiveWindow.ScrollRow = 160
    ActiveWindow.ScrollRow = 154
    ActiveWindow.ScrollRow = 148
    ActiveWindow.ScrollRow = 145
    ActiveWindow.ScrollRow = 139
    ActiveWindow.ScrollRow = 131
    ActiveWindow.ScrollRow = 122
    ActiveWindow.ScrollRow = 114
    ActiveWindow.ScrollRow = 99
    ActiveWindow.ScrollRow = 92
    ActiveWindow.ScrollRow = 84
    ActiveWindow.ScrollRow = 75
    ActiveWindow.ScrollRow = 65
    ActiveWindow.ScrollRow = 57
    ActiveWindow.ScrollRow = 52
    ActiveWindow.ScrollRow = 48
    ActiveWindow.ScrollRow = 45
    ActiveWindow.ScrollRow = 41
    ActiveWindow.ScrollRow = 38
    ActiveWindow.ScrollRow = 37
    ActiveWindow.ScrollRow = 34
    ActiveWindow.ScrollRow = 33
    ActiveWindow.ScrollRow = 32
    ActiveWindow.ScrollRow = 31
    ActiveWindow.ScrollRow = 30
    ActiveWindow.ScrollRow = 29
    ActiveWindow.ScrollRow = 27
    ActiveWindow.ScrollRow = 25
    ActiveWindow.ScrollRow = 23
    ActiveWindow.ScrollRow = 22
    ActiveWindow.ScrollRow = 20
    ActiveWindow.ScrollRow = 19
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 17
    ActiveWindow.ScrollRow = 13
    ActiveWindow.ScrollRow = 11
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 9
    ActiveWindow.ScrollRow = 8
    ActiveWindow.ScrollRow = 7
    ActiveWindow.ScrollRow = 6
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    ActiveWorkbook.Worksheets("Data").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Data").Sort.SortFields.Add Key:=Range("F2:F250"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    ActiveWorkbook.Worksheets("Data").Sort.SortFields.Add Key:=Range("G2:G250"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    ActiveWorkbook.Worksheets("Data").Sort.SortFields.Add Key:=Range("B2:B250"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Data").Sort.SortFields.Add Key:=Range("E2:E250"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Data").Sort
        .SetRange Range("B2:L250")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("A1").Select
End Sub
