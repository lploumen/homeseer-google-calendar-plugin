Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports System.IO
Imports System.Text
'Imports HomeSeer2.application


Module Module1
	
    Public hs As Object = Nothing 'Scheduler.hsapplication 'Object

    Friend Users As New Collections.SortedList

    Friend DaysAhead As Integer = 7
    Friend DaysBehind As Integer = 1

    Friend ConfigPage As clsConfigPage ' Class for our config web page.
	
    Public Const IFACE_NAME As String = "HSGCal"
    Public Const DateSort As String = "yyyyMMddHHmmss.fffff"

    Public gIOEnabled As Boolean
	Public callback As Object ' callback to HS object
    Public gBaseCode As String ' base housecode
    Public colBaseCodes As New Collections.SortedList
	Public Const MAX_IO_CODES As Short = 16
	Public Const MAX_DEVICE_CODES As Short = 64
	Public Const MAX_HOUSE_CODES As Short = 26 + 2 ' [ | for I/O devices
	
    Public InterfaceVersion As Short

    ' X10 commands
    Public Const UON As Short = 2
	Public Const UOFF As Short = 3
    Public Const NO_X10 As Short = 17 ' not x10 (unknown state)
    Public Const VALUE_SET As Short = 19 ' for hspi setio call
	
	' I/O types for device "iotype" property
	Public Const IOTYPE_INPUT As Short = 0
	Public Const IOTYPE_OUTPUT As Short = 1
	Public Const IOTYPE_ANALOG_INPUT As Short = 2
	Public Const IOTYPE_VARIABLE As Short = 3
	Public Const IOTYPE_CONTROL As Short = 4 ' new for ver 2 interface
	
	' interface status
	' for InterfaceStatus function call
	Public Const ERR_NONE As Short = 0
	Public Const ERR_SEND As Short = 1
	Public Const ERR_INIT As Short = 2
	
	' new for version 2 of interface
	Public Const UI_DROP_LIST As Short = 1 ' present user with drop down list of trigger options
	Public Const UI_TEXT_BOX As Short = 2 ' present user with a text box
	Public Const UI_CHECK_BOX As Short = 3 ' present user with a check box
	Public Const UI_BUTTON As Short = 4 ' button for remote dialogs
	Public Const UI_LABEL As Short = 5 ' label for display purposes only
	
	' attributes for controls
	Public Const CATTR_NO_EDIT As Short = &H1s ' bit=1 control cannot be edited by user
	Public Const CATTR_ALLOW_TEXT_ENTRY As Short = &H2s ' bit=2 for combo drop list contols, user can enter text
	

	' create some virtual devices from HS for use in displaying information
	' about your hardware. Devices can represent I/O points, temperature points, etc.
    Friend Sub InitDevices()
        Dim ft As Object
        Dim r As Object
        On Error Resume Next

        ' First, build a collection of house codes already in use by this plug-in.
        EnumerateDevices()

        If gBaseCode Is Nothing Then gBaseCode = ""
        If gBaseCode = "" Then
            If colBaseCodes.Count > 0 Then
                gBaseCode = colBaseCodes.GetByIndex(0)
            End If
        End If

        'CreateDevices()

    End Sub

    Friend UpdateThread As Threading.Thread
    Friend UpdateThread_Interval As New TimeSpan(0, 30, 0)
    Friend UpdateThread_Trigger As Boolean = False
    Friend Sub UpdateThread_Start()
        Dim Restart As Boolean = False
        If UpdateThread Is Nothing Then
            Restart = True
        ElseIf Not UpdateThread.IsAlive Then
            Restart = True
        End If
        If Restart Then
            UpdateThread = New Threading.Thread(AddressOf UpdateThread_Proc)
            UpdateThread.Name = "Calendar Update Thread"
            UpdateThread.Start()
        End If
    End Sub

    Friend Sub UpdateThread_Proc()
        Dim g As GoogleCalendar
        Dim dte As Date
        Dim DoneWaiting As Boolean = False
        Try
            Do
                If Users IsNot Nothing Then
                    If Users.Count > 0 Then
                        For i As Integer = 0 To Users.Count - 1
                            g = Users.GetByIndex(i)
                            If g IsNot Nothing Then
                                g.InitRecords(DaysAhead, DaysBehind)
                            End If
                        Next
                    End If
                End If
                dte = Now
                DoneWaiting = False
                Do
                    If UpdateThread_Trigger Then
                        UpdateThread_Trigger = False
                        DoneWaiting = True
                    Else
                        Threading.Thread.Sleep(1000)
                    End If
                    If Now.Subtract(dte) >= UpdateThread_Interval Then DoneWaiting = True
                Loop Until DoneWaiting
                DoneWaiting = False
            Loop
        Catch exa As Threading.ThreadAbortException
        Catch ex As Exception
            hs.WriteLog(IFACE_NAME & " Error", "Exception in Google Calendar Update Thread: " & ex.Message)
        End Try

    End Sub

    Public Sub EnumerateDevices()
        Dim DE As Object = hs.GetDeviceEnumerator
        Dim dv As Object = Nothing
        Dim s As String

        colBaseCodes.Clear()

        Try
            Do While Not DE.Finished
                dv = DE.GetNext
                If dv IsNot Nothing Then
                    If Trim(dv.interface) = IFACE_NAME Then
                        'hs.WriteLog(IFACE_NAME, "Found a plug-in device at " & dv.hc & dv.dc)
                        s = dv.hc
                        If Not colBaseCodes.ContainsKey(s.Trim.ToUpper) Then
                            colBaseCodes.Add(s.Trim.ToUpper, s.Trim.ToUpper)
                        End If
                    End If
                End If
            Loop
        Catch ex As Exception
            hs.WriteLog(IFACE_NAME & " Error", "Exception in EnumerateDevices: " & ex.Message)
        End Try


    End Sub

    Public Function GetDeviceCode(ByRef sCheck As String) As String
        Dim i As Short = 0
        Dim j As Short = 0
        Dim iHC As Integer = -1
        Dim s As String = ""
        Dim iRet As Short = -1
        Dim bRetry As Boolean = False

        ' EnumerateDevices went through all HomeSeer devices and if the device's Interface property was
        '   set to the name of this plug-in, then the house code for that device was stored in colBaseCodes.
        ' This procedure takes as a starting point a house code or optional house code with unit code as 
        '   a starting point, and returnes the next available address.  If there is an address in the same
        '   house code available, then that is returned (house code and device code).  If one is not available,
        '   then GetNextFreeIOCode is called to set up a new house code.
        ' You should NOT save your house code in an INI file for this reason:  You call GetNextFreeIOCode and 
        '   are vended the house code "(" for example, but you do not create a device - you just save that house code.
        '   Now along comes another plug-in that does the same thing, and since there are no devices with the "("
        '   house code, it is also given the "(" house code, only that plug-in creates a device at the address "(1".
        '   If you saved the house code and assumed you could now use it, you would be wrong - the only way to save
        '   a house code for a plug-in is to actually create a device using that house code.

        If Not sCheck Is Nothing Then
            s = sCheck.Trim
        End If
Retry:
        If s = "" Then
            If colBaseCodes.Count > 0 Then
                s = colBaseCodes.GetByIndex(0)
            End If
        End If
        If s Is Nothing Then GoTo GetNewHC
        If Len(s.Trim) = 0 Then GoTo GetNewHC

        If Len(s.Trim) > 1 Then
            j = Val(Mid(s.Trim, 2))
        Else
            j = 1
        End If
        s = Left(s, 1)
        For i = j To 99
            iRet = hs.DeviceExists(s.ToUpper & i.ToString)
            If iRet = -1 Then
                Return s.ToUpper & i.ToString
            End If
        Next

        'If we fell through to here, then there was nothing found with that starting house code (sCheck).
        If colBaseCodes.Count = 0 Then GoTo GetNewHC
        If bRetry Then
            Try
                colBaseCodes.RemoveAt(0)
            Catch ex As Exception
            End Try
        Else
            bRetry = True
        End If
        If colBaseCodes.Count = 0 Then GoTo GetNewHC
        s = ""
        GoTo Retry

GetNewHC:
        'hs.WriteLog(IFACE_NAME, "Calling GetNextFreeIOCode")
        iHC = callback.GetNextFreeIOCode
        'hs.WriteLog(IFACE_NAME, "GetNextFreeIOCode returned " & iHC.ToString & " (" & Chr(iHC) & ")")

        ' exit if all codes are used
        If iHC = -1 Then
            hs.WriteLog(IFACE_NAME & " Error", "Sorry, all device codes used.")
            Return ""
        End If

        ' use the ascii version of the basecode
        gBaseCode = Chr(iHC)
        Return gBaseCode & "1"

    End Function


	' create some virtual devices to represent our hardware
	' called from setup dialog
	' each housecode allocated gives you 64 devices
	' get more housecodes for more devices
	Sub CreateDevices()
		Dim i As Short
		Dim j As Short
		Dim h As Short
		Dim d As Short
		Dim index As Short
		Dim lIndex As Integer
		Dim dv As Object
		Dim unit As Short
        Dim dev_code As Short
		Dim DE As Object
		On Error Resume Next

        Return ' lpl simply don't process this function

		
        ' First, check to see if our devices are created already...
        '   EnumerateDevices was called earlier, so colBaseCodes would be populated if there are devices.
        If colBaseCodes.Count > 0 Then
            ' Devices exist.
            Exit Sub
        End If

        ' call back to HS to get a free housecode if necessary
        gBaseCode = GetDeviceCode(gBaseCode)
        If gBaseCode = "" Then Exit Sub

        ' create our devices. We will create one device for each zone of our security panel
		' this allows us to display the status of each zone. Each zone device will also have
		' a config button for specific configuration if necessary
		'
		' for other hardware, only one device may be needed to display status
        ' note the special "iotype" property that must be set
        If Len(gBaseCode) > 1 Then
            dev_code = Val(Mid(gBaseCode, 2))
            gBaseCode = Left(gBaseCode, 1)
        Else
            dev_code = 1
        End If
        dv = hs.GetDevice(index)
        ' lpl removed acme devices...

        ' create one device to represent the global panel status, such as armed, etc.
        'If InterfaceVersion < 3 Then
        ' index = hs.NewDevice("Acme Panel Status")
        'dv = hs.GetDevice(index)
        'Else
        'lIndex = hs.NewDeviceRef("Acme Panel Status")
        'dv = hs.GetDeviceByRef(lIndex)
        'End If
        'dv.location = IFACE_NAME
        'dv.hc = gBaseCode
        'dv.dc = dev_code.ToString
        'dv.interface = IFACE_NAME
        'dv.misc = 0 ' On/Off only, no dim
        'dv.dev_type_string = "Panel Status"
        'dv.iotype = IOTYPE_CONTROL
        ' add 2 buttons to this device
        'dv.buttons = IFACE_NAME & Chr(2) & "Arm" & Chr(1) & IFACE_NAME & Chr(2) & "Disarm"
        ' set a default status for this device to a string saying we are not connected to the panel
        ' hs.SetDeviceString(dv.hc & dv.dc, "No Connection")
        dev_code += 1
        For i = 1 To 5
            ' this is a good place to actually assign the real zone names
            If InterfaceVersion < 3 Then
                index = hs.NewDevice("Zone " & i.ToString)
                dv = hs.GetDevice(index)
            Else
                lIndex = hs.NewDeviceRef("Zone " & i.ToString)
                dv = hs.GetDeviceByRef(lIndex)
            End If
            dv.location = IFACE_NAME
            dv.hc = gBaseCode
            dv.dc = dev_code.ToString
            dv.interface = IFACE_NAME
            dv.misc = 0 ' On/Off only, no dim
            ' The following simply shows up in the device properties but has no other use
            dv.dev_type_string = "Security Zone"
            ' specify the type of device
            ' if your device is an I/O input point use: IOTYPE_INPUT
            ' if your device is an I/O output point use: IOTYPE_OUPUT
            ' if your device is a variable use: IOTYPE_VARIABLE
            ' if your device is a controllable piece of hardware such as an MP3 player or security panel zone use: IOTYPE_CONTROL
            ' we will use the CONTROL type since we are representing security zones
            dv.iotype = IOTYPE_CONTROL
            ' add possible values for this device
            dv.values = "Bypassed" & Chr(2) & "1" & Chr(1) & "Not Bypassed" & Chr(2) & "2"
            dev_code += 1
        Next
        ' add more devices here if needed

        ' Set the house code to the last device used so that a call to GetDeviceCode
        '   will return the next one available.
        gBaseCode &= dev_code.ToString

    End Sub

    Friend Sub LoadSettings()

        DaysAhead = CInt(Val(hs.GetINISetting("Settings", "DaysAhead", "7", "HSGCal.ini").Trim))
        DaysBehind = CInt(Val(hs.GetINISetting("Settings", "DaysBehind", "1", "HSGCal.ini").Trim))
        UpdateThread_Interval = New TimeSpan
        Dim m As Double = Val(hs.GetINISetting("Settings", "Refresh", "30", "HSGCal.ini").trim)
        UpdateThread_Interval = TimeSpan.FromMinutes(m)

    End Sub
    Friend Sub SaveSettings()

        hs.SaveINISetting("Settings", "DaysAhead", DaysAhead.ToString, "HSGCal.ini")
        hs.SaveINISetting("Settings", "DaysBehind", DaysBehind.ToString, "HSGCal.ini")
        hs.SaveINISetting("Settings", "Refresh", UpdateThread_Interval.TotalMinutes.ToString, "HSGCal.ini")

    End Sub

    Friend Function LoadUsers() As String

        'gHSServerPort = Trim(hs.GetINISetting("Settings", "svrport", "0"))

        Try

            Dim g As GoogleCalendar
            Dim u As GoogleCalendar.UserPair
            Dim i As Integer = 0
            Dim uName As String = ""
            Dim uPass As String = ""
            Dim uFriend As String = ""

            If Users Is Nothing Then Users = New Collections.SortedList

            Do
TryAnother:
                i += 1
                uName = hs.GetINISetting("Users", "uName" & i.ToString, "none", "HSGCal.ini").Trim
                If uName.ToLower = "none" Then Exit Do
                If Not uName.Contains("@") Then
                    hs.WriteLog(IFACE_NAME & " Warning", uName & " does not appear to be a valid eMail address for a Google Calendar account.")
                    GoTo TryAnother
                End If
                uPass = hs.GetINISetting("Users", "uPass" & i.ToString, "none", "HSGCal.ini").Trim
                If uName.ToLower = "none" Then
                    hs.WriteLog(IFACE_NAME & " Warning", uName & " not loaded, password is missing from the INI file.")
                    GoTo TryAnother
                End If

                uFriend = hs.GetINISetting("Users", "uFriend" & i.ToString, "none", "HSGCal.ini").Trim
                If uFriend.ToLower = "none" Then
                    uFriend = uName
                End If
                u = New GoogleCalendar.UserPair
                u.UserName = uName.Trim
                u.Password = uPass.Trim
                u.FriendlyName = uFriend.Trim
                g = New GoogleCalendar(u)
                Try
                    Users.Add(u.UserName.ToLower, g)
                Catch ex As Exception
                    If Users.ContainsKey(u.UserName.ToLower) Then
                        Users.Remove(u.UserName.ToLower)
                        Try
                            Users.Add(u.UserName.ToLower, g)
                        Catch ex2 As Exception
                            hs.WriteLog(IFACE_NAME & " Error", "Failed to add new user record for " & u.UserName & ": " & ex2.Message)
                        End Try
                    End If
                End Try
            Loop
            hs.WriteLog(IFACE_NAME, "Loaded " & Users.Count.ToString & " users.")
        Catch ex As Exception
            hs.WriteLog(IFACE_NAME & " Error", "Exception loading users: " & ex.Message)
            Return "Exception loading users: " & ex.Message
        End Try

        'Dim evs() As Google.GData.Calendar.EventEntry
        'For x As Integer = 0 To Users.Count - 1
        '    g = Users.GetByIndex(x)
        '    If g IsNot Nothing Then
        '        evs = g.GetAll
        '        If evs IsNot Nothing Then
        '            If evs.Length > 0 Then
        '                For i As Integer = 0 To IIf(evs.Length > 4, 5, evs.Length - 1)
        '                    hs.WriteLog(IFACE_NAME, evs(i).Title.Text & " by " & evs(i).Authors(0).Email & " starting on " & evs(i).Times(0).StartTime.TimeOfDay.ToString)
        '                Next
        '            End If
        '        End If
        '    End If
        'Next

        Return ""

    End Function

    Friend Sub SaveUsers()

        Try

            Dim g As GoogleCalendar
            Dim u As GoogleCalendar.UserPair
            Dim i As Integer = 0

            If Users Is Nothing Then Exit Sub
            If Users.Count < 1 Then Exit Sub

            hs.ClearINISection("Users", "HSGCal.ini")
            For i = 0 To Users.Count - 1
                g = Users.GetByIndex(i)
                If g IsNot Nothing Then
                    u = g.User
                    If u IsNot Nothing Then
                        hs.SaveINISetting("Users", "uName" & (i + 1).ToString, u.UserName, "HSGCal.ini")
                        hs.SaveINISetting("Users", "uPass" & (i + 1).ToString, u.Password, "HSGCal.ini")
                        hs.SaveINISetting("Users", "uFriend" & (i + 1).ToString, u.FriendlyName, "HSGCal.ini")
                    End If
                End If
            Next
        Catch ex As Exception
            hs.WriteLog(IFACE_NAME & " Error", "Exception saving users: " & ex.Message)
        End Try

    End Sub

	Sub wait(ByRef secs As Short)
        Dim s As Long
		
        s = VB.Timer()
		Do 
			System.Windows.Forms.Application.DoEvents()
        Loop While VB.Timer() < s + secs

    End Sub

    Friend Sub SetupEnvironment(ByVal FILE_NAME As String)
        Try
            Dim sver As String = ""
            Dim ourver As String = ""
            Dim sdate As New Date
            Dim ourdate As New Date
            Dim EXEPath As String = My.Application.Info.DirectoryPath

            ' FILE_NAME is the name of your compiled file, usually hspi_(appname).dll
            ' This procedure copies the application file to the HomeSeer BIN directory so that ASPX pages that
            '   reference this plug-in will have the reference available.

            If FILE_NAME Is Nothing Then
                hs.WriteLog(IFACE_NAME, "Error, SetupEnvironment was called with a null FILE_NAME parameter.")
                Exit Sub
            End If
            If FILE_NAME = "" Then
                hs.WriteLog(IFACE_NAME, "Error, SetupEnvironment was called with an empty FILE_NAME parameter.")
                Exit Sub
            End If

            If System.IO.File.Exists(EXEPath & "\html\bin\" & FILE_NAME) Then
                Try
                    sver = GetVersionInfo(EXEPath & "\html\bin\" & FILE_NAME)
                    ourver = GetVersionInfo(EXEPath & "\" & FILE_NAME)
                    sdate = My.Computer.FileSystem.GetFileInfo(EXEPath & "\html\bin\" & FILE_NAME).LastWriteTime
                    ourdate = My.Computer.FileSystem.GetFileInfo(EXEPath & "\" & FILE_NAME).LastWriteTime
                Catch ex As Exception
                End Try
                If (sver <> ourver) Or (sdate <> ourdate) Then
                    Try
                        System.IO.File.Copy(EXEPath & "\" & FILE_NAME, EXEPath & "\html\bin\" & FILE_NAME, True)
                    Catch ex As Exception
                        hs.WriteLog(IFACE_NAME, "Error copying " & FILE_NAME & " to html\bin folder, resource file may not be accessible. " & ex.Message)
                    End Try
                End If
            Else
                Try
                    System.IO.File.Copy(EXEPath & "\" & FILE_NAME, EXEPath & "\html\bin\" & FILE_NAME, True)
                Catch ex As Exception
                    hs.WriteLog(IFACE_NAME, "Error copying " & FILE_NAME & " to html\bin folder, resource file may not be accessible. " & ex.Message)
                End Try
            End If

        Catch ex As Exception
            hs.WriteLog(IFACE_NAME, "Error setting up environment: " & ex.Message)
        End Try
    End Sub

    Public Function GetResourceToFile(ByVal resname As String, ByVal filename As String) As String
        ' Filename can be a path to be added to the root directory and the file will be created using resname 
        '   as the filename, or filename can be a fully qualified path and filename relative to the root path.
        ' EXAMPLES:
        '               \HTML\
        '               \HTML\MyApplication\
        '               \HTML\MyApp\DogEatDog.htm
        '
        Dim thisExe As System.Reflection.Assembly
        Dim bteArray As Byte()
        Dim i As Integer
        Dim fs As FileStream
        Dim EXEPath As String = System.IO.Path.GetDirectoryName(System.AppDomain.CurrentDomain.BaseDirectory)
        Dim s As String = ""
        Dim p As String = ""

        Try
            If filename.Trim.EndsWith("\") Or filename.Trim.EndsWith("/") Then
                s = EXEPath & Left(filename.Trim, Len(filename.Trim) - 1) & "\" & resname
            Else
                s = EXEPath & filename.Trim
            End If
            p = Path.GetFullPath(s)
            If Not IO.Directory.Exists(p) Then
                IO.Directory.CreateDirectory(p)
            End If
        Catch ex As Exception
            Return "Error, unable to save " & resname & " configuration page to the " & p & " folder: " & ex.Message
        End Try

        If (s = "") Or (p = "") Then
            Return "Error, cooked path or file are invalid. Path=" & p & ", file=" & s
        End If

        If File.Exists(s) Then
            Try
                File.Delete(s)
            Catch ex As Exception
                Return "Deleting file " & s & ", Error (1) returned is " & ex.Message
            End Try
        End If

        Try
            fs = New FileStream(s, FileMode.CreateNew)
        Catch ex As DirectoryNotFoundException
            Return "Cannot create file, the containing folder does not exist: " & s
        Catch ex As Exception
            Return "Creating file stream for " & resname & ", Error (2) returned is " & ex.Message
        End Try

        Dim w As New BinaryWriter(fs)
        thisExe = System.Reflection.Assembly.GetExecutingAssembly()
        Dim filest As System.IO.Stream

        Try
            filest = thisExe.GetManifestResourceStream(resname)
        Catch ex As Exception
            Return "Creating resource file stream for " & resname & ", Error (3) returned is " & ex.Message
        End Try

        If filest Is Nothing Then
            Return "Embedded resource " & resname & " was not found in the manifest assembly."
        End If

        ReDim bteArray(filest.Length)
        Try
            i = filest.Read(bteArray, 0, filest.Length)
        Catch ex As Exception
            Return "Reading resource stream from " & resname & ", Error (4) returned is " & ex.Message
        End Try

        Try
            fs.Write(bteArray, 0, bteArray.Length)
        Catch ex As Exception
            Return "Writing file from stream resource " & resname & ", Error (5) returned is " & ex.Message
        End Try

        w.Close()
        fs.Close()

        Return ""
    End Function

    Public Function GetVersionInfo(ByVal sFile As String) As String

        GetVersionInfo = "N/A"

        Dim myFileVersionInfo As System.Diagnostics.FileVersionInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(sFile)

        Try
            GetVersionInfo = myFileVersionInfo.FileMajorPart.ToString & "." & _
              myFileVersionInfo.FileMinorPart.ToString & "." & _
              myFileVersionInfo.FileBuildPart.ToString & "." & _
              myFileVersionInfo.FilePrivatePart.ToString
        Catch ex As Exception
            hs.WriteLog(IFACE_NAME, "Error in GetVersionInfo, " & ex.Message)
        End Try

    End Function

    Public Function stripHTML(ByVal strHTML As String) As String
        'Strips the HTML tags from strHTML

        Dim strOutput As String = ""
        Dim objRegExp As RegularExpressions.Regex = New RegularExpressions.Regex("<(.|\n)+?>", RegularExpressions.RegexOptions.IgnoreCase)

        'Replace all HTML tag matches with the empty string
        strOutput = objRegExp.Replace(strHTML, "")

        stripHTML = strOutput         'Return the value of strOutput

        objRegExp = Nothing
    End Function


End Module