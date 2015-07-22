Attribute VB_Name = "Office2010-Mod1"
Option Explicit
Dim CalFolder As Outlook.Folder
Dim printCal As Outlook.Folder
Dim k As Integer

Sub PrintCalendarsAsOne()
    '******************************************************************
    '  Author - Tom Uyemura
    '           Barefoot Bay Recreation District
    '  Language - Visual Basic for Applications - VBA
    '  Created Aug. 22nd , 2014
    
    '  Lines usually changed are lines that contain variable "sFilter"
    '        and that the correct folders are SELECTED
    '
    '  PrintCalendarsAsOne()
    '  This program will consolidate all selected ( checked ) calendars
    '       in My Calendars
    '       into one folder (calendar) named Print.  Any existing folder
    '       named "Print" will be DELETED at the begining of this program
    '
    '  CopyAppttoPrint() Subroutine will just choose appointments within
    '     the date range specified by the sFilter variable
    '
    '
    '******************************************************************

    Dim objPane      As Outlook.NavigationPane
    Dim objModule    As Outlook.CalendarModule
    Dim objGroup     As Outlook.NavigationGroup
    Dim objNavFolder As Outlook.NavigationFolder
    Dim objCalendar  As Folder
    Dim objFolder    As Folder
       
    Dim i            As Integer
    Dim g            As Integer
    Dim j            As Integer
    
    Dim theFolder    As String
    theFolder = "Print"
      
    On Error Resume Next
     
    Set objCalendar = Session.GetDefaultFolder(olFolderCalendar)
    Set printCal = objCalendar.Folders(theFolder)
    printCal.Delete
    Set printCal = objCalendar.Folders.Add(theFolder)
       
    Set Application.ActiveExplorer.CurrentFolder = objCalendar
    DoEvents
       
    Set objPane = Application.ActiveExplorer.NavigationPane
    Set objModule = objPane.Modules.GetNavigationModule(olModuleCalendar)
       
  With objModule.NavigationGroups
     
    For g = 1 To .Count
 
    Set objGroup = .Item(g)
    
    For i = 1 To objGroup.NavigationFolders.Count
        Set objNavFolder = objGroup.NavigationFolders.Item(i)
        If objNavFolder.IsSelected = True Then
      
           'run macro to copy appt
           Set CalFolder = objNavFolder.Folder
           CopyAppttoPrint
           j = j + 1
        End If
    
    Next i
    Next g
    
    End With
   
    MsgBox ("PrintCalendarsAsOne Macro Completed")
    MsgBox (g & " Folders processed, " & vbCrLf & _
            j & " total items processed, " & vbCrLf & _
            k & " total appointments processed")
    
    Set objPane = Nothing
    Set objModule = Nothing
    Set objGroup = Nothing
    Set objNavFolder = Nothing
    Set objCalendar = Nothing
    Set objFolder = Nothing
    i = 0
    g = 0
    j = 0
    k = 0
End Sub
  
    '******************************************************************
    '  Author - Tom Uyemura
    '           Barefoot Bay Recreation District
    '  Language - Visual Basic for Applications - VBA
    '  Created Jan. 23rd , 2015
    '
    '  This program works with PrintCalendarsAsOne.  It is call only once but within a recursive statement.
    '    Located around line 62!!
    '
    '******************************************************************
  
Sub CopyAppttoPrint()
      
   Dim calItems As Outlook.Items
   Dim ResItems As Outlook.Items
   Dim calName  As Variant
   Dim sFilter As String
   Dim iNumRestricted As Integer
   Dim itm, newAppt As Object
  
   Set calItems = CalFolder.Items
     
   If CalFolder = printCal Then
     Exit Sub
   End If
     
' Sort all of the appointments based on the start time
   calItems.Sort "[Start]"
   calItems.IncludeRecurrences = True
  
  calName = CalFolder.Parent.Name
' to use category named for account & calendar name
'  calName = CalFolder.Parent.Name & "-" & CalFolder.Name
      
'create the filter - this copies appointments today to 3 days from now
   sFilter = "[Start] >= '" & Date & "'" & " And [Start] < '" & Date + 400 & "'"
   
   ' Apply the filter
   Set ResItems = calItems.Restrict(sFilter)
   
   iNumRestricted = 0
   
   'Loop through the items in the collection.
   For Each itm In ResItems
      iNumRestricted = iNumRestricted + 1
      k = k + 1
      Set newAppt = printCal.Items.Add(olAppointmentItem)
   
 With newAppt
 ' delete any lines you don't need to include
    .Start = itm.Start
    .End = itm.End
    .Subject = itm.Subject
    .Body = itm.Body
    .Location = itm.Location
    .AllDayEvent = itm.AllDayEvent
    .Categories = calName '& ";" & itm.Categories
    .ReminderSet = False
End With
            
  newAppt.Save
   
   Next
   ' Display the actual number of appointments created
    Debug.Print calName & " " & (iNumRestricted & " appointments were created")
   
   Set itm = Nothing
   Set newAppt = Nothing
   Set ResItems = Nothing
   Set calItems = Nothing
   Set CalFolder = Nothing
     
End Sub

Sub DeleteBOCAppointments()
    '******************************************************************
    '  Author - Tom Uyemura
    '           Barefoot Bay Recreation District
    '  Language - Visual Basic for Applications - VBA
    '  Created Aug. 22nd , 2014
    '
    '  This program will remove all appointments that are listed in the
    '    very long IF statement below for the Calendar listed in the variable
    '    "theMonth".  Usually set for the "Print" calendar.
    '
    '   ** NOTE *** This program will work ONLY for one folder(calendar) at a time
    '     and verify variable theMonth has the correct folder name
    '
    '
    '
    '   TO-DO ( later )
    '     consolidate that if statement to an array or table or something
    '      Less complicated.
    '
    '******************************************************************

    ' See http://support.microsoft.com/kb/285202 for Outlook constants.

    ' Declare all variables.
    Dim objOutlook             As Outlook.Application
    Dim objNamespace           As Outlook.NameSpace
    Dim objFolder              As Outlook.MAPIFolder
    Dim objAppointment         As Outlook.AppointmentItem
    Dim objAttachment          As Outlook.Attachment
    Dim objVariant             As Variant
    Dim lngDeletedAppointments As Long
    Dim lngCleanedAppointments As Long
    Dim lngCleanedAttachments  As Long
    Dim intCount               As Integer
    Dim intDateDiff            As Integer
    
    Dim theString              As String
    Dim theMonth               As String
    theMonth = "Print"

    ' Create an object for the Outlook application.
    Set objOutlook = Application
    ' Retrieve an object for the MAPI namespace.
    Set objNamespace = objOutlook.GetNamespace("MAPI")
    ' Retrieve a folder object for the default calendar folder.
    ' Set objFolder = objNamespace.GetDefaultFolder(olFolderCalendar)
    Set objFolder = Session.GetDefaultFolder(olFolderCalendar).Folders(theMonth)
    
    

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
            lngCleanedAppointments = lngCleanedAppointments + 1
            theString = objAppointment
                        
            ' Look for year-old non-recurring appointments.
            If (objAppointment = "Needlepoint ") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "River of Life") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "River of Life Church Service for Children") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "Grace Christian Fellowship Church") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "MJ's  Exercisers ") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "MJ's  Exercisers") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "Chair Exercisers") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "Exercisers") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "Canasta") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "Cribbage ") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "Cribbage - Sunday Afternoon") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "Cribbage - Monday Nights") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "Gentle Yoga") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "Art Group") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "Zumba Class") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "Zumba Class - Outside building A around Pool 1") Then
                ' Delete the appointment.  Added March 30th, 2015
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "Billiards") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "Bridge Club") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "Tuesday - Duplicate Bridge Club") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "Ladies Poker") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "Men's Poker") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "Mens Poker ") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "Aqua Zumba") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "AA Meeting") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "Ladies Billiards") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "Crafters/Scrapbooking") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "Scrapbooking/Digital - Just for the High Season Starting in October") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "Scrapbooking/Digital - Just for the High Season") Then
                ' Delete the appointment.  Added 3/30/2015
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
'            ElseIf (objAppointment = "Chess Club") Then
'                  REMOVED BECAUSE OF CHANGE TO A VARIABLE DAY SCHEDULE
'                ' Delete the appointment.
'                objAppointment.Delete
'                ' Increment the count of deleted appointments.
'                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "Scrabble") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "Scrabble ") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "Saturday Bridge Club") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "Mah-Jong - 5/7/14 Disbanded ") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "Mah-Jong - 5/7/14 Disbanded") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "All You Can Eat Pasta") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "Euchre") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "Maintenance") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "Tops # 456") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "Pinochle") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "Pinochle ") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "The Golden Oldies Band - Suspended until Fall") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "The Golden Oldies Band") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "All that Jazz Dance Group") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "Monday Shuffleboard Active Oct-Apr") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "Monday Shuffleboard Club") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "BFB Orchestra ") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "Maintenance Crew") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "Line Dancing") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "Trinity Peace Church") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "(NO LONGER Meeting)Trinity Peace Church") Then
                ' Delete the appointment.  Added 3/30/2015
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "Mah Jongg") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "MJ's Chair Exercisers") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
             ElseIf (objAppointment = "Custodian Cleanup Time") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
             ElseIf (objAppointment = "Custodian Cleanup Period") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "Custodian Cleanup") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            ElseIf (objAppointment = "Custodian Setup Time") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
             ElseIf (objAppointment = "Smoothies Dance Club") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
             ElseIf (objAppointment = "Marine Corp Setup") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
             ElseIf (objAppointment = "(ARE THEY STILL Meeting??)The Golden Oldies Band - Suspended until Fall") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
             ElseIf (objAppointment = "Zumba Class - Outside building A around Pool 1") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
             ElseIf (objAppointment = "") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
             ElseIf (objAppointment = "") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
             ElseIf (objAppointment = "") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
             ElseIf (objAppointment = "BBRD Tentative Tentatively RESERVED") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
             ElseIf (objAppointment = "Maintenance Crew set up for BFB TRustees Meeting") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
             ElseIf (objAppointment = "Determined Youth") Then
                ' Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
             ElseIf Left(theString, 10) = "(CANCELLED" Then
'                MsgBox " Tested positive for (CANCELLED equals: " & objAppointment
                ' Delete the appointment.
                 objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
             ElseIf Left(theString, 10) = "(Tentative" Then
'                MsgBox " Tested positive for (Tentative equals: " & objAppointment
                ' Delete the appointment.
                 objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
             ElseIf Left(theString, 11) = "(NO MEETING" Then
'                MsgBox " Tested positive for (NO MEETING: " & objAppointment
                ' Delete the appointment.
                  objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
             ElseIf Left(theString, 1) = "(" Then
 '               MsgBox " Tested positive for (: " & objAppointment
                ' Delete the appointment.
                 objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
             End If

'            If Left("(CANCELLED") Then
'               objAppointment.Delete
'               lngDeletedAppointments = lngDeletedAppointments + 1
'            End If

        End If
    
   Next

    ' Display the number of calendar objects that were cleaned or deleted.
   MsgBox "Deleted " & lngDeletedAppointments & " appointment(s)." & vbCrLf & _
      "Cleaned " & lngCleanedAppointments & " appointment(s)." & vbCrLf & _
      " "
'      "Deleted " & lngCleanedAttachments & " attachment(s)."

End Sub

Sub KeepBOCAppointments()
    '******************************************************************
    '  Author - Tom Uyemura
    '           Barefoot Bay Recreation District
    '  Language - Visual Basic for Applications - VBA
    '  Created Jan. 23rd , 2015
    '
    '  This program will Keep all appointments that are listed in the
    '    very long IF statement below for the Calendar listed in the variable
    '    "theMonth" and Delete all other.
    '    This list should be KEPT in Syncronization with DeleteBOCAppointments!!
    '
    '   ** NOTE *** This program will work ONLY for one folder(calendar) at a time
    '     and verify variable theMonth has the correct folder name
    '
    '
    '
    '   TO-DO ( later )
    '     consolidate that if statement to an array or table or something
    '      Less complicated.
    '
    '******************************************************************

    ' See http://support.microsoft.com/kb/285202 for Outlook constants.

    ' Declare all variables.
    Dim objOutlook             As Outlook.Application
    Dim objNamespace           As Outlook.NameSpace
    Dim objFolder              As Outlook.MAPIFolder
    Dim objAppointment         As Outlook.AppointmentItem
    Dim objAttachment          As Outlook.Attachment
    Dim objVariant             As Variant
    Dim lngDeletedAppointments As Long
    Dim lngCleanedAppointments As Long
    Dim lngCleanedAttachments  As Long
    Dim intCount               As Integer
    Dim intDateDiff            As Integer
    Dim intFoundFlag           As Integer
    
    Dim theMonth               As String
    theMonth = "Print"
    intFoundFlag = 0

    ' Create an object for the Outlook application.
    Set objOutlook = Application
    ' Retrieve an object for the MAPI namespace.
    Set objNamespace = objOutlook.GetNamespace("MAPI")
    ' Retrieve a folder object for the default calendar folder.
    ' Set objFolder = objNamespace.GetDefaultFolder(olFolderCalendar)
    Set objFolder = Session.GetDefaultFolder(olFolderCalendar).Folders(theMonth)
    
    

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
            lngCleanedAppointments = lngCleanedAppointments + 1
            intFoundFlag = 0
                        
            ' Look for year-old non-recurring appointments.
            If (objAppointment = "Needlepoint ") Or _
               (objAppointment = "River of Life") Or _
                (objAppointment = "River of Life Church Service for Children") Or _
                (objAppointment = "Grace Christian Fellowship Church") Or _
                (objAppointment = "MJ's  Exercisers ") Or _
                (objAppointment = "MJ's  Exercisers") Or _
                (objAppointment = "Chair Exercisers") Or _
                (objAppointment = "Exercisers") Or _
                (objAppointment = "Canasta") Or _
                (objAppointment = "Cribbage ") Or _
                (objAppointment = "Cribbage - Sunday Afternoon") Or _
                (objAppointment = "Cribbage - Monday Nights") Or _
                (objAppointment = "Gentle Yoga") Or _
                (objAppointment = "Art Group") Or _
                (objAppointment = "Zumba Class") Or _
                (objAppointment = "Billiards") Or _
                (objAppointment = "Bridge Club") Or _
                (objAppointment = "Tuesday - Duplicate Bridge Club") Or _
                (objAppointment = "Ladies Poker") Or _
                (objAppointment = "Men's Poker") Or _
                (objAppointment = "Mens Poker ") Or _
                (objAppointment = "Aqua Zumba") Or _
                (objAppointment = "AA Meeting") Then
           '    set flag
                intFoundFlag = 1
            End If
                
            If (objAppointment = "Ladies Billiards") Or _
                (objAppointment = "Crafters/Scrapbooking") Or _
                (objAppointment = "Scrapbooking/Digital - Just for the High Season Starting in October") Or _
                (objAppointment = "Scrapbooking/Digital - Just for the High Season") Or _
                (objAppointment = "Scrabble") Or _
                (objAppointment = "Scrabble ") Or _
                (objAppointment = "Saturday Bridge Club") Or _
                (objAppointment = "Mah-Jong - 5/7/14 Disbanded ") Or _
                (objAppointment = "Mah-Jong - 5/7/14 Disbanded") Or _
                (objAppointment = "All You Can Eat Pasta") Or _
                (objAppointment = "Euchre") Or _
                (objAppointment = "Pinochle") Or _
                (objAppointment = "Maintenance") Or _
                (objAppointment = "Tops # 456") Or _
                (objAppointment = "Pinochle") Or _
                (objAppointment = "Pinochle ") Or _
                (objAppointment = "The Golden Oldies Band - Suspended until Fall") Or _
                (objAppointment = "The Golden Oldies Band") Or _
                (objAppointment = "All that Jazz Dance Group") Or _
                (objAppointment = "Monday Shuffleboard Active Oct-Apr") Or _
                (objAppointment = "Monday Shuffleboard Club") Or _
                (objAppointment = "BFB Orchestra ") Or _
                (objAppointment = "Maintenance Crew") Or _
                (objAppointment = "Mah Jongg") Then
            '   set flag
                intFoundFlag = 1
            End If
            
            If (objAppointment = "MJ's Chair Exercisers") Or _
                (objAppointment = "Custodian Cleanup Time") Or _
                (objAppointment = "Line Dancing") Or _
                (objAppointment = "Trinity Peace Church") Or _
                (objAppointment = "NO LONGER Meeting)Trinity Peace Church") Or _
                (objAppointment = "Custodian Cleanup Period") Or _
                (objAppointment = "Custodian Cleanup") Or _
                (objAppointment = "Custodian Setup Time") Or _
                (objAppointment = "Zumba Class - Outside building A around Pool 1") Or _
                (objAppointment = "Smoothies Dance Club") Then
            '   set Flag
                intFoundFlag = 1
            End If
            
            If intFoundFlag = 0 Then
                ' If NOT on the above List , then Delete the appointment.
                objAppointment.Delete
                ' Increment the count of deleted appointments.
                lngDeletedAppointments = lngDeletedAppointments + 1
            End If
            
        End If
    
   Next

    ' Display the number of calendar objects that were cleaned or deleted.
   MsgBox "Deleted " & lngDeletedAppointments & " appointment(s)." & vbCrLf & _
      "Cleaned " & lngCleanedAppointments & " appointment(s)." & vbCrLf & _
      " "
'      "Deleted " & lngCleanedAttachments & " attachment(s)."

End Sub
  
Sub InsSalutationShort()
    Application.ActiveWindow.Insert ("Test")
End Sub
