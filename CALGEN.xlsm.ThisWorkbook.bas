Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
' *****************************************************
' **  Begining of Workbook_BeforeClose() Subroutine ***
' *****************************************************

'  Don't Let the End User notice that this spreadsheet is NOT
'   being saved.
'
    Me.Saved = True
' *****************************************************
' **  End of Workbook_BeforeClose() Subroutine      ***
' *****************************************************
End Sub

Private Sub Workbook_Open()
Dim wkName, wkFullName, _
    wkPath, _
    vFile1, vFile2 _
             As Variant

  wkName = ThisWorkbook.Name
  wkPath = ThisWorkbook.Path
  
  vFile1 = wkPath & "\" & "CALGEN_Import_TEMPLATE--01.xlsm"
  vFile2 = wkPath & "\" & "CALGEN_TEMPLATE_Calendar-Final_02.xlsm"

'   If the name is anything else, then this file is being opened by
'       the CALGEN_TEMPLATE_Calendar-Final_02 spreadsheet and WE DON't
'        want the Import_and_Sort Macro to RUN!
If UCase(Left(wkName, 6)) = "CALGEN" Then
    Application.Visible = True
   
    Workbooks.Open fileName:=vFile1, ReadOnly:=True
    Workbooks.Open fileName:=vFile2, ReadOnly:=True
    
    MsgBox " Success!  VERY GOOD WORK ! " & _
              vbCrLf & " The screen will now close. " & vbCrLf & vbCrLf & _
              " Thank You for using the Calendar Generator Application ! :) ! "
    
    Close
    Application.Quit
    
End If

End Sub
'This will go into the ThisWorkbook vb basic in the CALGEN.xlsm spreadsheet macro area
