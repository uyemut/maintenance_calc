Attribute VB_Name = "Office2010-Mod3"
Option Explicit
Dim objPane As NavigationPane
Private Sub Application_Startup()
    Set objPane = Application.ActiveExplorer.NavigationPane
   
End Sub

Private Sub objPane_ModuleSwitch(ByVal CurrentModule As NavigationModule)
 
    Dim objModule As CalendarModule
    Dim objGroup As NavigationGroup
    Dim objNavFolder As NavigationFolder
    Dim objCalendar As Folder
    Dim objFolder As Folder
      
    Dim i As Integer
      
    If CurrentModule.NavigationModuleType = olModuleCalendar Then
    Set Application.ActiveExplorer.CurrentFolder = Session.GetDefaultFolder(olFolderCalendar)
    DoEvents
      
    Set objCalendar = Session.GetDefaultFolder(olFolderCalendar)
    Set objPane = Application.ActiveExplorer.NavigationPane
    Set objModule = objPane.Modules.GetNavigationModule(olModuleCalendar)
      
    With objModule.NavigationGroups
        Set objGroup = .GetDefaultNavigationGroup(olMyFoldersGroup)
  
    ' To use a different group
       ' Set objGroup = .Item("group name")
    End With
  
  
    For i = 1 To objGroup.NavigationFolders.Count
        Set objNavFolder = objGroup.NavigationFolders.Item(i)
        Select Case i
  
        ' Enter the calendar index numbers you want to open
            Case 1, 3, 4
                objNavFolder.IsSelected = True
               
        ' Set to True to open side by side
                objNavFolder.IsSideBySide = False
            Case Else
                objNavFolder.IsSelected = False
        End Select
    Next
    End If
 
    Set objPane = Nothing
    Set objModule = Nothing
    Set objGroup = Nothing
    Set objNavFolder = Nothing
    Set objCalendar = Nothing
    Set objFolder = Nothing
 
End Sub

Sub DeleteOldAppointments()

    ' See http://support.microsoft.com/kb/285202 for Outlook constants.

    ' Declare all variables.
    Dim objOutlook As Outlook.Application
    Dim objNamespace As Outlook.NameSpace
    Dim objFolder As Outlook.MAPIFolder
    Dim objAppointment As Outlook.AppointmentItem
    Dim objAttachment As Outlook.Attachment
    Dim objVariant As Variant
    Dim lngDeletedAppointments As Long
    Dim lngCleanedAppointments As Long
    Dim lngCleanedAttachments As Long
    Dim intCount As Integer
    Dim intDateDiff As Integer

    ' Create an object for the Outlook application.
    Set objOutlook = Application
    ' Retrieve an object for the MAPI namespace.
    Set objNamespace = objOutlook.GetNamespace("MAPI")
    ' Retrieve a folder object for the default calendar folder.
    Set objFolder = objNamespace.GetDefaultFolder(olFolderCalendar)

    ' Loop through the items in the folder. NOTE: This has to
    ' be done backwards; if you process forwards you have to
    ' re-run the macro an inverese exponential number of times.
    For intCount = objFolder.Items.Count To 1 Step -1
        ' Retrieve an object from the folder.
        Set objVariant = objFolder.Items.Item(intCount)
        
        ' Allow the system to process. (Helps you to cancel the
        ' macro, or continue to use Outlook in the background.)
        DoEvents
        
        ' Filter objects for appointments/meetings.
        If objVariant.Class = olAppointment Then
            ' Create an appointment object from the current object.
            Set objAppointment = objVariant
        
            ' This is optional, but it helps me to see in the
            ' debug window where the macro is currently at.
            Debug.Print objAppointment.Start
            
            ' Calculate the difference in days between
            ' now and the date of the calendar object.
            intDateDiff = DateDiff("d", objAppointment.Start, Now)
            
            ' Look for year-old non-recurring appointments.
            If intDateDiff > 365 And objAppointment.RecurrenceState = olApptNotRecurring Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
                
            ' Delete attachments from 6-month-old non-recurring appointments.
            ElseIf intDateDiff > 180 And objAppointment.RecurrenceState = olApptNotRecurring Then
                ' Test if the calendar object has attachments.
                If objAppointment.Attachments.Count > 0 Then
                    ' Loop through the attachments collection.
                    While objAppointment.Attachments.Count > 0
                        ' Delete the current attachment.
                        objAppointment.Attachments.Remove 1
                        ' Increment the count of deleted attachments.
                        lngCleanedAttachments = lngCleanedAttachments + 1
                    Wend
                    ' Increment the count of cleaned appointments.
                    lngCleanedAppointments = lngCleanedAppointments + 1
                End If
            
            ' Delete large attachments from 60-day-old appointments.
            ElseIf intDateDiff > 60 Then
                ' Test if the calendar object has attachments.
                If objAppointment.Attachments.Count > 0 Then
                    ' Loop through the attachments collection.
                    For Each objAttachment In objAppointment.Attachments
                        ' Test if the attachment is too large.
                        If objAttachment.Size > 500000 Then
                            ' Delete the current attachment.
                            objAttachment.Delete
                            ' Increment the count of deleted attachments.
                            lngCleanedAttachments = lngCleanedAttachments + 1
                        End If
                    Next
                    ' Increment the count of cleaned appointments.
                    lngCleanedAppointments = lngCleanedAppointments + 1
                End If
            End If
        
        End If
        
   Next

    ' Display the number of calendar objects that were cleaned or deleted.
   MsgBox "Deleted " & lngDeletedAppointments & " appointment(s)." & vbCrLf & _
      "Cleaned " & lngCleanedAppointments & " appointment(s)." & vbCrLf & _
      "Deleted " & lngCleanedAttachments & " attachment(s)."

End Sub

Sub ResetCalendarsFolders()

    '******************************************************************
    '  Author - Tom Uyemura
    '           Barefoot Bay Recreation District
    '  Language - Visual Basic for Applications - VBA
    '  Created Oct. 6th, 2014
    '
    '  ResetCalendarsFolders()
    '  This program will delete the building calendars and recreate
    '       them with the same names. Calendars to be reset will be:
    '       Bldg A
    '       Bldg C
    '       Bldg D/E
    '
    '******************************************************************
    
    Dim objCalendar  As Folder
    Dim objFolder    As Folder
    Dim printCal     As Folder
         
    On Error Resume Next
     
    Set objCalendar = Session.GetDefaultFolder(olFolderCalendar)
    
    Set printCal = objCalendar.Folders("Bldg A")
    printCal.Delete
    Set printCal = objCalendar.Folders.Add("Bldg A")

    Set printCal = objCalendar.Folders("Bldg C")
    printCal.Delete
    Set printCal = objCalendar.Folders.Add("Bldg C")

    Set printCal = objCalendar.Folders("Bldg D/E")
    printCal.Delete
    Set printCal = objCalendar.Folders.Add("Bldg D/E")

End Sub
