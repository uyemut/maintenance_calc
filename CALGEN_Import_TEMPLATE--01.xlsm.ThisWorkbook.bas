Option Explicit
' *****************************************************
' **  These macros will go into the ThisWorkbook    ***
' **  Global area on close , and on open triggers   ***
' *****************************************************

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
Dim wkName As Variant

'MsgBox " debug here "
'Debug.Assert breakpoint

' Lets find the name of this workbook.  It should begin with
'     CALGEN_IMPORT_TEMPLATE*.*
' Windows("CALGEN_Import_TEMPLATE--01.xlsm").Activate
  wkName = ThisWorkbook.Name


'   If the name is anything else, then this file is being opened by
'       the CALGEN_TEMPLATE_Calendar-Final_02 spreadsheet and WE DON't
'        want the Import_and_Sort Macro to RUN!
If UCase(Left(wkName, 22)) = "CALGEN_IMPORT_TEMPLATE" Then
    Application.Visible = True
   
    Import_and_Sort
End If

' Old Code, not applicable
'If Err.Number <> 0 Then
'    Application.Visible = True
'Else
'    Application.Visible = False
'    Application.Visible = True
'
'    Import_and_Sort
'End If

End Sub
