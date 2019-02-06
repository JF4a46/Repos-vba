Private Type hourSettings
    heure As String
    plage As String
End Type

Sub createGcdocsFolder()
inac = Module1.IEgcdocs()


End Sub

Sub createMeetingInvitation()

    Dim app As Outlook.Application
    Dim meeting As Outlook.AppointmentItem
    
    'Instance des Objets
    Set app = Outlook.Application    'Instance de l'application
    Set meeting = app.CreateItem(olAppointmentItem)  'Instance de la nouvelle entrée du calendrier
    meeting.MeetingStatus = olMeeting
    Dim template As Outlook.MailItem
    Set template = app.CreateItemFromTemplate("")
    

    'Affichage de l'entrée du calendrier
    meeting.Display
    sName = InputBox(Prompt:="Press a button to let IE set the meeting", XPos:=15000, YPos:=8000)
   
    
    
        Dim LArray() As String
        LArray = Split(meeting.start)
        dateDebut = Module1.inverse(LArray(0))
        'Debug.Print dateDebut
        temp = Split(LArray(1), ":")
        heure = temp(0)
        Dim plage As hourSettings
        plage = determinePlage(heure)
        minutes = temp(1)
        
        Debug.Print dateDebut
        
        Link = IEWebex(meeting.Subject, dateDebut, plage.heure, plage.plage, minutes)
        
        meeting.Body = template.Body
        meeting.Location = Link
        
   
    
    Set app = Nothing
    Set meeting = Nothing
    

End Sub



Function IEWebex(ByVal titre As String, ByVal dateDebut As String, ByVal heure As String, ByVal plage As String, ByVal minutes As String) As String
'This will load a webpage in IE
    Dim i As Long
    Dim url As String
    Dim IE As Object
    Dim frame As Object
    Dim usernameInput As Object
    Dim passwordInput As Object
    Dim logonBtn As Object
    Dim username As String
    Dim password As String
    username = ""
    password = ""
 
    'Create InternetExplorer Object
    Set IE = CreateObject("InternetExplorer.Application")
 
    'Set IE.Visible = True to make IE visible, or False for IE to run in the background
    IE.Visible = True
 
    'Define URL
    url = ""
    Call Module1.loadPage(url, IE)

    Set frame = IE.Document.getelementsbyname("mainFrame")(0)

    On Error GoTo alreadyLogged
    
    Set usernameInput = frame.contentDocument.getelementsbyname("userName")(0)
    Set passwordInput = frame.contentDocument.getelementsbyname("password")(0)
    usernameInput.Value = username
    passwordInput.Value = password
    Set logonBtn = frame.contentDocument.getelementsbyname("btnLogon")(0)
    logonBtn.Click
    
alreadyLogged:
    Application.Wait Now + #12:00:02 AM#

    Set frame = IE.Document.getelementsbyname("header")(0)
    frame.contentDocument.getelementbyid("wcc-lnk-MC").Click
    Application.Wait Now + #12:00:05 AM#
    Set frame = IE.Document.getelementsbyname("mainFrame")(0).contentDocument.getelementsbyname("menu")(0).contentDocument.getelementsbyname("treemenu")(0)
    frame.contentDocument.getelementbyid("wcc-lnk-scheduleaMeeting").Click
    
    Application.Wait Now + #12:00:02 AM#
    
    Set frame = IE.Document.getelementsbyname("mainFrame")(0).contentDocument.getelementsbyname("main")(0).contentDocument
    
    frame.getelementsbyname("ConfName")(0).Value = titre
    
    'frame.getelementbyid("datePickerContainer").getelementsbyclassname("sr-only")(0).innerText = dateDebut
    'frame.getelementbyid("wcc-ipt-startDatePicker").Value = dateDebut
    Dim dayPicker As Object
    Dim monthPicker As Object
    Dim yearPicker As Object
    Dim LArray() As String
    LArray = Split(dateDebut, "/")
    Set dayPicker = frame.getelementsbyname("startDateOfDay")(0)
    Set monthPicker = frame.getelementsbyname("startDateOfMonth")(0)
    Set yearPicker = frame.getelementsbyname("startDateOfYear")(0)
    
    monthPicker.selectedindex = LArray(0) - 1
    dayPicker.selectedindex = LArray(1) - 1
    'yearPicker.selectedindex = LArray(2)
    
    frame.getelementsbyname("startTimeOfHour")(0).Value = heure
    
    Application.Wait Now + #12:00:01 AM#
    
    If plage = "am" Then
    Debug.Print "I shall press AM"
    frame.getelementbyid("wcc-rd-startTimeOfAmPm-am").Click
    'frame.getelementbyid("input-radio-1").Value = "1"
    'frame.getelementbyid("input-radio-2").Value = "0"
    End If
    
    If plage = "pm" Then
    Debug.Print "I shall press PM"
    frame.getelementbyid("wcc-rd-startTimeOfAmPm-pm").Click
    'frame.getelementbyid("input-radio-1").Value = "0"
    'frame.getelementbyid("input-radio-2").Value = "1"
    End If
    
    frame.getelementsbyname("startTimeOfMinute")(0).Value = minutes
    frame.getelementsbyname("DurationHour")(0).Value = "0"
    frame.getelementsbyname("DurationMinute")(0).Value = "30"
    
    'frame.getelementbyid("wcc-ipt-startDatePicker").Value = dateDebut
    frame.getelementbyid("wcc-btn-schedule").Click
    'frame.getelementsbyname("WizardForm")(0).submit
    
    Application.Wait Now + #12:00:02 AM#
    
    Set frame = IE.Document.getelementsbyname("mainFrame")(0).contentDocument.getelementsbyname("main")(0).contentDocument
    IEWebex = frame.getelementbyid("mc-ipt-meetingLink").Value
    'Unload IE
    IE.Quit
    Set IE = Nothing
    Set objElement = Nothing
    Set objCollection = Nothing
    
End Function

Function IEgcdocs()
    
    Dim username As String
    Dim password As String
    username = ""
    password = ""
    Dim url As String
    Dim IE As InternetExplorer
    
    url = ""
    Set IE = New InternetExplorerMedium
    
    'Set IE.Visible = True to make IE visible, or False for IE to run in the background
    IE.Visible = True
 
    Call loadPage(url, IE)
    

    IE.Document.getelementbyid("Username").Value = username
    IE.Document.getelementbyid("Password").Value = password
    IE.Document.getelementbyid("loginbutton").Click
    url = "https://gcdocs.gc.ca/ssc-spc/llisapi.dll?func=ll&objType=0&objAction=create&parentId=17657835&nextURL=%2Fssc%2Dspc%2Fllisapi%2Edll%3Ffunc%3Dll%26objId%3D17657835%26objAction%3Dbrowse%26viewType%3D1"
    Call loadPage(url, IE)

    IE.Document.getelementbyid("name").Value = "test6"
    IE.Document.getelementbyid("addButton").Click

    Application.Wait Now + #12:00:02 AM#
   
    
    Set IE = Nothing
End Function




Function loadPage(url As String, IE As Object)
    IE.Navigate url
    Debug.Print url
    'IE ReadyState = 4 signifies the webpage has loaded (the first loop is set to avoid inadvertently skipping over the second loop)
    Do While IE.readyState = 4: DoEvents: Loop   'Do While
    Do Until IE.readyState = 4: DoEvents: Loop   'Do Until
    'Webpage Loaded
End Function

Function waitLoading(IE As Object)

    Do While IE.readyState = 4: DoEvents: Loop   'Do While
    Do Until IE.readyState = 4: DoEvents: Loop   'Do Until

End Function

Function determinePlage(ByVal heure As String) As hourSettings
Dim ret As hourSettings
    With ret
    .heure = ""
    .plage = ""
    End With

    intHeure = CInt(heure)
    
    If intHeure > 12 Then
        ret.heure = CStr(intHeure Mod 12)
        ret.plage = "pm"
    End If
    If intHeure = 12 Then
        ret.heure = "12"
        ret.plage = "pm"
    End If
    If intHeure < 12 Then
        ret.heure = CStr(intHeure)
        ret.plage = "am"
    End If
    
    determinePlage = ret
End Function


Function inverse(start As String) As String
    Dim LArray() As String
    LArray = Split(start, "-")
    Length = UBound(LArray, 1) - LBound(LArray, 1) + 1
    
    inverse = LArray(1) & "/" & LArray(2) & "/" & LArray(0)
'    Dim ret As String
'    For i = (Length - 1) To 0 Step -1
'    ret = ret & "/" & LArray(i)
'
'
'    Next i
'
'    inverse = Right(ret, Len(ret) - 1)


End Function

