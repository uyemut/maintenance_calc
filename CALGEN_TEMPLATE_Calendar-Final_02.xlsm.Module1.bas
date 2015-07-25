Attribute VB_Name = "Module1"
' *****************************************************
' **  These macros will go into the Reg Workbook    ***
' **  Macro Module area. The name is Module1.bas in ***
' **  the CALGEN_TEMPLATE_Calendar-Final_02.xlsm    ***
' **   workbook.                                    ***
' *****************************************************
Sub KopyAndPaste()
'************************************************************************************
'***                                                                              ***
'***  Program - KcopyAndPaste()                                                   ***
'***  Author - Thomas Uyemura                                                     ***
'***             Calendar Coordinator                                             ***
'***  Written - July 29th, 2015                                                   ***
'***  This script will copy the formatted eight line per day calendar to the      ***
'***    Barefoot Bay Recreation District's calendar format.                       ***
'***                                                                              ***
'************************************************************************************
    Dim iRowLastItem As Integer
    Dim vFile1 As Variant
    Dim wb1, _
        wb2  As Workbook
    Dim vFile31, vPath31 As Variant
    Dim ws1, _
        ws2, _
        ws3, _
        ws4  As Worksheet
    Dim wkNewMonth, wsNewYear _
             As String
    Dim CurrentFile, _
        NewFileType, _
        NewFileName, _
        NewFile   As String
    Dim Response As VbMsgBoxResult
         
    Set wb1 = ActiveWorkbook
    Set ws1 = wb1.Worksheets("Data")
    Set ws3 = wb1.Worksheets("Month")
    Set ws4 = wb1.Worksheets("NewCalendar")
    Application.ScreenUpdating = False
    '******************************************************************
    '  This section will do the Importing
    '******************************************************************
'    vFile1 = Application.GetOpenFilename("Excel-files,*.*", _
'            1, "Select the saved 'CALGEN_Import_TEMPLATE--01' File To Open", , False)
    'if the user didn't select a file, exit sub
'    If TypeName(vFile1) = "Boolean" Then Exit Sub
    vPath31 = ThisWorkbook.Path
    vFile31 = "CalendarCordCalGen_TemporaryFile" & Format(Date, "yyyymmdd") & ".xls"
    vFile1 = vPath31 & "\" & vFile31              '  Let's make up a tempfile on our OWN.
    Workbooks.Open Filename:=vFile1, ReadOnly:=True
    
    'Set targetworkbook
    Set wb2 = ActiveWorkbook
    Set ws2 = wb2.Worksheets("Data")
    iRowLastItem = ws2.Cells.Find(What:="*", SearchOrder:=xlRows, _
                   SearchDirection:=xlPrevious, LookIn:=xlValues).Row + 1

    ws2.Range("A2:Z" & iRowLastItem).Copy
    ws1.Range("A2").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:= _
                    False, Transpose:=False
    ws1.Range("A2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, _
                    Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'    ws1.Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, _
'                    Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'   DEBUGGING AND TESTING LINE

'   We are done with the spreadsheet we have just opened, let us close it now
    wb2.Close SaveChanges:=False
'   Now, let us delete it.
    DeleteFile (vFile1)

'     MsgBox " Testing 1 2 3 "
     wkNewMonth = InputBox("Enter the New Month for this Calendar")
     wkNewYear = InputBox("Enter the YEAR for this Calendar")
 
    ws3.Range("A1") = wkNewMonth
    ws4.Range("A1") = wkNewMonth
    ws3.Range("F1") = wkNewYear
    ws4.Range("F1") = wkNewYear
    ws3.Range("A5:G12").Copy
    ws4.Range("A5").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, _
                   Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ws3.Range("A15:G22").Copy
    ws4.Range("A15").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, _
                Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ws3.Range("A25:G32").Copy
    ws4.Range("A25").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, _
                Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ws3.Range("A35:G42").Copy
    ws4.Range("A36").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, _
                 Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ws3.Range("A45:G52").Copy
    ws4.Range("A46").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, _
                 Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ws3.Range("A1").Copy
    ws3.Activate
    ws3.Range("A1").Select
    
    ws1.Activate
    ws1.Range("A1").Select
    ws4.Activate
    ws4.Range("A1").Select
    
    CalTimeEdits
    
    CalDescriptionEdits
    
    Response = MsgBox(" Are you ready to save your NEW Calendar for the month of " & _
                      wkNewMonth & " and year of " & wkNewYear & " (Y/N)? " & vbCrLf & _
                     " If answer is No, You'll need to save it seperately" _
                      , vbQuestion + vbYesNo)
                      
    Application.ScreenUpdating = True
        
    If Response = vbNo Then
       MsgBox " The KopyAndPaste macro is successfully completed. VERY GOOD WORK ! Take a look at your FINAL " & _
                wkNewCalendar & " Calendar! .  Don't forget to SAVE AS - Excel 97-2003, Let's not change THIS TEMPLATE folks!"
        ws4.Name = wkNewMonth
       Exit Sub
    End If

    CurrentFile = ThisWorkbook.FullName
    
    MsgBox (" The next screen will ask you to save the New Calendar in a seperate excel spreadsheet.")
 
    NewFileType = "Excel Files 1997-2003 (*.xls), *.xls," & _
               "Excel Files 2007 (*.xlsx), *.xlsx," & _
               "All files (*.*), *.*"
 
    NewFile = Application.GetSaveAsFilename( _
        InitialFileName:=NewFileName, _
        fileFilter:=NewFileType)
 
    If NewFile <> "" And NewFile <> "False" Then
        Application.DisplayAlerts = False
        ws1.Delete
        ws3.Delete
        ws4.Name = wkNewMonth
        Application.DisplayAlerts = True
        ActiveWorkbook.SaveAs Filename:=NewFile, _
            FileFormat:=xlNormal, _
            Password:="", _
            WriteResPassword:="", _
            ReadOnlyRecommended:=False, _
            CreateBackup:=False
     Else
        Workbooks.Close
'        Application.Quit
     End If
    
'    MsgBox " The KopyAndPaste macro has successfully completed for the month of " & wkNewMonth & _
'            "." & vbCrLf & "VERY GOOD WORK ! " & _
'              vbCrLf & " The screen will now close. " & vbCrLf & vbCrLf & _
'              " Thank You for using the Calendar Generator Application ! :) ! "

End Sub
Sub CalTimeEdits()
'
' CalEdits Macro
'
'************************************************************************************
'***                                                                              ***
'***  Program - CalTimeEdits()                                                    ***
'***  Author - Thomas Uyemura                                                     ***
'***             Calendar Coordinator                                             ***
'***  Written - July 29th, 2015                                                   ***
'***  This Macro will formatted the time into :30 and :15 time increments         ***
'***    from the decimal internal excel formatting                                ***
'***    Barefoot Bay Recreation District's calendar format.                       ***
'***                                                                              ***
'************************************************************************************

'
    Range("A1:A2").Select
    Cells.Replace What:=".5-", Replacement:=":30-", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:=".5pm", Replacement:=":30pm", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:=".5am", Replacement:=":30am", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    ActiveCell.Replace What:="-0:30pm", Replacement:="12:30pm", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:=" 0:30pm", Replacement:=" 12:30pm", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="-0:30pm", Replacement:="-12:30pm", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:=" 0:30am", Replacement:=" 12:30am", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub


Sub CalDescriptionEdits()
'
' CalEdits Macro
'

'
    Range("A1:A2").Select
    Cells.Replace What:="E 12-4pm - Red Hats", _
    Replacement:="E 1-4pm - Red Hats", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="D 4-6pm - Gospel Singers", _
    Replacement:="D 5-6pm - Gospel Singers", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="C 9-11am - Green Thumb Club", _
    Replacement:="C 10-11am - Green Thumb Club", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="A 4-10pm - Polish Club Monthly Meeting", _
    Replacement:="A 5-10pm - Polish Club Monthly Meeting", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="A 1-4pm - Senior Singles", _
    Replacement:="A 2-4pm - Senior Singles", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="9:30-12pm - Violations Comm", _
    Replacement:="10am - Violations Comm", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="9:30-12am - Violations Comm", _
    Replacement:="10am - Violations Comm", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="12-5pm - Board of Trustee", _
    Replacement:="1pm - Board of Trustee", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="7:30-9:30pm - Yacht Club", _
    Replacement:="7:30-9:30pm-Boat & Fishing Club", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="A 6:30-11pm - HOA Board Meeting", _
    Replacement:="A 7pm - HOA Board Meeting", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="C 9:30-12pm - VETS Council", _
    Replacement:="C 10-12pm - VETS Council", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="C 9:30-12am - VETS Council", _
    Replacement:="C 10-12pm - VETS Council", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="1-2:30pm - RV CLUB", _
    Replacement:="1:30-2:30pm - RV CLUB", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="E 1-4pm - Recreation Committee", _
    Replacement:="E 2-4pm - Recreation Committee", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="D 2-4:30pm - BFB Conservative Club", _
    Replacement:="D 2:30-4:30pm - BFB Conservative Club", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="D&E 6:30-9pm - Democratic Club", _
    Replacement:="D&E 7-9pm - Democratic Club", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="D&E 5-10pm - Computer Club General Meeting", _
    Replacement:="D&E 6-10pm - Computer Club General Meeting", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="D&E 5-10pm - Computer Club Round Table Meeting", _
        Replacement:="D&E 6-10pm - Computer Club Round Table Mtg.", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="E 12:30-3pm - HOA - Orientation", _
    Replacement:="E 1-3pm - HOA - Orientation", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="7-10pm Board of Trustee", _
    Replacement:="7pm - Board of Trustee", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="R 9-11am - ARCC Meeting", _
    Replacement:="RR 9am - ARCC Meeting", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="D&E 11-3pm - Men's Golf - General Meeting", _
    Replacement:="D&E 11:30-3pm - Men's Golf - General Meeting", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

'    IF  Range("A1").Cells == "JUNE" or "JULY" or "AUGUST"
'        Cells.Replace What:="Computer Club General Meeting", _
'   Replacement:="Computer Club Round Table Mtg", L'ookAt:=xlPart, _
'        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
'        ReplaceFormat:=False
'    END-IF

End Sub

Function FileExists(ByVal FileToTest As String) As Boolean
   FileExists = (Dir(FileToTest) <> "")
End Function

Sub DeleteFile(ByVal FileToDelete As String)
   If FileExists(FileToDelete) Then 'See above
      SetAttr FileToDelete, vbNormal
      Kill FileToDelete
   Else
      MsgBox " Very bad Error, Where is file " & _
             FileToDelete & "?"
   End If
End Sub

'7  *** DRAFT *** *** DRAFT ***  FOR INTERNAL DISTRIBUTION ONLY
'D&E 11:30-3pm - Men's Golf - General Meeting

