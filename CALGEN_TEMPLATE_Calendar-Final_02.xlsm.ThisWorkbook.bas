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
'   Application.Visible = False
    
    'MsgBox " debug here "
    'Debug.Assert breakpoint

    ' Lets find the name of this workbook.  It should begin with
    '     CALGEN_TEMPLATE_CALEND*.*
  wkName = ThisWorkbook.Name


'   If the name is anything else, then this file is a Calendar File
'       excel spreadsheet and WE DON't
'        want the automatically run any MACROS !!!!
  If UCase(Left(wkName, 22)) = "CALGEN_TEMPLATE_CALEND" Then
     Application.Visible = True
   
     KopyAndPaste
     
     Workbooks.Close
   End If

    
End Sub
