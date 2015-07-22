Attribute VB_Name = "Office2010-Mod2"
Option Explicit
'**************************************************
' Title: PST2MSG
'
' Description:
' This VB application will export selected
' Outlook folders to file system as MSG files.
' The intent is to allow quick reference when
' burned to CD due to Outlook not opening
' Read Only PST files.
'
' Use: Paste the code into a VB5/6 module
' There is an optional Form explained in code
'
' Notes:
' This code is offered 'As Is'.
' No support will be provided by me.
'
' Author: Steven Harvey
' Free to use for all
'
'**************************************************
Public Const BIF_RETURNONLYFSDIRS = 1
Public Const MAX_PATH = 260
 
Public Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
 
'APIs for the folder selection
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
 
Private objNS As NameSpace
Private objFolder As Outlook.MAPIFolder
Private strDestination As String
Private strTopFolder As String
Private strLogFile As String
Private strFolderName As String
Private intErrors As Boolean
Public intUserAbort As Integer
 
Sub Main()
  Set objNS = Application.GetNamespace("MAPI")
  Set objFolder = objNS.PickFolder
   
  If Not objFolder Is Nothing Then
    strTopFolder = CleanString(objFolder.Name)
    strDestination = GetFileDir
    If strDestination <> "" Or strDestination <> Null Then
      strFolderName = CleanString(objFolder.Name)
      strLogFile = strDestination & "\" & strFolderName & "\ExportLog.txt"
      strDestination = strDestination & "\" & strFolderName
      If FolderExist(strDestination) Then
        MsgBox "This folder has already been exported here. Please clear the destination or choose another."
        Exit Sub
      Else
        '****** frmProcessing displays while processing messages.
        '****** It has a message asking user to wait while processing.
        '****** It also has a cancel button to set intUserAbort to 1.
        '****** Form's button code is below
        '*** Private Sub cmdCancel_Click()
        '***   intUserAbort = 1
        '***   Unload Me
        '*** End Sub
        'frmProcessing.Show
         
        intUserAbort = 0
        ProcessFolder objFolder, strDestination
         
        'Unload frmProcessing
         
        If intUserAbort = 0 Then
            MsgBox "Export Complete!" & vbCrLf & "Export log file location:" & vbCrLf & strLogFile
        Else
            MsgBox "Processing cancelled." & vbCrLf & "Export log file location:" & vbCrLf & strLogFile
        End If
      End If
    Else
      MsgBox "Destination folder selection cancelled!"
    End If
  Else
    MsgBox "MAPI folder selection cancelled!"
  End If
 
Set objNS = Nothing
Set objFolder = Nothing
End Sub
 
Function FolderExist(sFileName As String) As Boolean
  FolderExist = IIf(Dir(sFileName, vbDirectory) <> "", True, False)
End Function
 
Public Function StripNulls(ByVal OriginalStr As String) As String
    If (InStr(OriginalStr, Chr$(0)) > 0) Then
        OriginalStr = Left$(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    StripNulls = OriginalStr
End Function
 
Public Function GetFileDir() As String
Dim ret As String
    Dim lpIDList As Long
    Dim sPath As String, udtBI As BrowseInfo
    Dim RdStrings() As String, nNewFiles As Long
 
    'Show a browse-for-folder form:
    With udtBI
        .lpszTitle = lstrcat("Please select a folder to export to:", "")
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
 
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList = 0 Then Exit Function
         
    'Get the selected folder.
    sPath = String$(MAX_PATH, 0)
    SHGetPathFromIDList lpIDList, sPath
    CoTaskMemFree lpIDList
    sPath = StripNulls(sPath)
    GetFileDir = sPath
End Function
 
Sub ProcessFolder(StartFolder As Outlook.MAPIFolder, strPath As String)
    Dim objItem As Object
         
'      frmProcessing.Label2.Caption = "Processing " & StartFolder
    MkDir strPath
    ' process all the items in this folder
    For Each objItem In StartFolder.Items
      DoEvents
      SaveAsMsg objItem, strPath
      Set objItem = Nothing
    Next
     
    ' process all the subfolders of this folder
    For Each objFolder In StartFolder.Folders
        Dim strSubFolder As String
        strSubFolder = strPath & "\" & CleanString(StartFolder.Name)
        Call ProcessFolder(objFolder, strSubFolder)
    Next
     
    Set objFolder = Nothing
    Set objItem = Nothing
End Sub
 
Private Function CleanString(strData As String) As String
    'Replace invalid strings.
    strData = ReplaceChar(strData, "_", "")
    strData = ReplaceChar(strData, "´", "'")
    strData = ReplaceChar(strData, "`", "'")
    strData = ReplaceChar(strData, "{", "(")
    strData = ReplaceChar(strData, "[", "(")
    strData = ReplaceChar(strData, "]", ")")
    strData = ReplaceChar(strData, "}", ")")
    strData = ReplaceChar(strData, "/", "-")
    strData = ReplaceChar(strData, "\", "-")
    strData = ReplaceChar(strData, ":", "")
    strData = ReplaceChar(strData, ",", "")
    'Cut out invalid signs.
    strData = ReplaceChar(strData, "*", "'")
    strData = ReplaceChar(strData, "?", "")
    strData = ReplaceChar(strData, """", "'")
    strData = ReplaceChar(strData, "<", "")
    strData = ReplaceChar(strData, ">", "")
    strData = ReplaceChar(strData, "|", "")
    CleanString = Trim(strData)
End Function
 
Private Function ReplaceChar(strData As String, strBadChar As String, strGoodChar As String) As String
Dim i As Long
Dim tmpChar, tmpString As String
    For i = 1 To Len(strData)
      tmpChar = Mid(strData, i, 1)
      If tmpChar = strBadChar Then
        tmpString = tmpString & strGoodChar
      Else
        tmpString = tmpString & tmpChar
      End If
    Next i
    ReplaceChar = Trim(tmpString)
End Function
 
Private Sub SaveAsMsg(objItem As Object, strFolderPath As String)
On Error Resume Next
Dim strSubject As String
Dim strSaveName As String
 
    Err.Clear
    If Not objItem Is Nothing Then
      Select Case TypeName(objItem)
        Case "AppointmentItem"
          strSaveName = Format(objItem.Start, "mm-dd-yyyy hh.nn.ss") & " " & IIf(Len(strFolderPath & objItem.Subject) > 255, Left(objItem.Subject, 255 - Len(strFolderPath)), objItem.Subject) & ".msg"
        Case "MailItem"
          strSaveName = Format(objItem.ReceivedTime, "mm-dd-yyyy hh.nn.ss") & " " & IIf(Len(strFolderPath & objItem.Subject) > 255, Left(objItem.Subject, 255 - Len(strFolderPath)), objItem.Subject) & ".msg"
          If Err Then
              WriteToLog "Error #" & Err.Number & ": " & Err.Description & " Unable to process message '" & strFolderPath & "\" & objItem.Subject & "'."
              strSaveName = strFolderPath & "\" & objItem.Subject & ".msg"
          End If
        Case "NoteItem"
          strSaveName = objItem.Subject & ".msg"
        Case "TaskItem"
          strSaveName = objItem.Subject & ".msg"
        Case "ContactItem"
          strSaveName = objItem.FileAs & ".msg"
        Case Else
          strSaveName = ""
      End Select
        Err.Clear
        objItem.SaveAs strFolderPath & "\" & CleanString(strSaveName), olMSG
        If Err Then
            WriteToLog "Error #" & Err.Number & ": " & Err.Description & " Unable to process message '" & strFolderPath & "\" & objItem.Subject & "'."
        Else
          WriteToLog "Success: " & strFolderPath & "\" & CleanString(strSaveName)
        End If
    End If
End Sub
 
Private Sub WriteToLog(strMessage As String)
  intErrors = True
  Open strLogFile For Append As #1
  Write #1, strMessage
  Close #1
End Sub
Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
   Dim myRecipient As Items
   If Item.MessageClass = "IPM.Note" Then
       For Each myRecipient In Item.Recipients
'           If myRecipient.Address = "<EMAIL ADDRESS TO FIND>" Then
'           Test Test
'           End If
       Next
   End If
End Sub
