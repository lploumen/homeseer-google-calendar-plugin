Imports System
Imports System.Collections
Imports Google.GData.Client
Imports Google.GData.Extensions
Imports Google.GData.Calendar
Imports System.Web


<Serializable()> _
Public Class GoogleCalendar

    <Serializable()> _
    Public Class UserPair
        Public Sub New()
            MyBase.New()
        End Sub
        Public UserName As String = ""
        Public Password As String = ""
        Public FriendlyName As String = ""
        Public ShowPass As Boolean = False
    End Class


    Public Sub New(ByRef User As UserPair)
        MyBase.new()
        If String.IsNullOrEmpty(User.UserName) Then
            Throw New Exception("CalData.New called with an empty name value.")
            Exit Sub
        End If
        If String.IsNullOrEmpty(User.Password) Then
            Throw New Exception("CalData.New called with an empty user password.")
            Exit Sub
        End If
        mvarUser = User
        mvarUser.UserName = User.UserName.Trim
        mvarUser.Password = User.Password.Trim
        mvarUser.FriendlyName = User.FriendlyName.Trim
        ComStatus = True
        mvarEntries = New Collections.Generic.List(Of EventEntry)
        service = New CalendarService("HSGCalApp_" & mvarUser.UserName)
    End Sub

    Private mvarUser As UserPair
    Private mvarEntries As Collections.Generic.List(Of EventEntry)
    Private WithEvents service As CalendarService 'New CalendarService("HSGCalApp")
    Private AuthToken As String = ""
    Private LastAuth As Date

    Public ComStatus As Boolean = False

    Public ReadOnly Property User() As UserPair
        Get
            Return mvarUser
        End Get
    End Property
    Public ReadOnly Property UserName() As String
        Get
            If mvarUser IsNot Nothing Then
                Return mvarUser.UserName
            Else
                Return ""
            End If
        End Get
    End Property
    Public ReadOnly Property FriendlyName() As String
        Get
            If mvarUser IsNot Nothing Then
                Return mvarUser.FriendlyName
            Else
                Return ""
            End If
        End Get
    End Property
    'Public WriteOnly Property Set_Event_List() As EventEntry()
    '    Set(ByVal value As EventEntry())
    '        If value Is Nothing Then Exit Property
    '        If value.Length < 1 Then Exit Property
    '        If mvarEntries Is Nothing Then mvarEntries = New Collections.Generic.List(Of EventEntry)
    '        If mvarEntries.Count > 0 Then mvarEntries.Clear()
    '        For x As Integer = 0 To value.Length - 1
    '            mvarEntries.Add(value(x))
    '        Next
    '    End Set
    'End Property
    Public ReadOnly Property Count() As Integer
        Get
            If mvarEntries IsNot Nothing Then
                Return mvarEntries.Count
            Else
                Return 0
            End If
        End Get
    End Property
    'Public WriteOnly Property Add() As EventEntry
    '    Set(ByVal value As EventEntry)
    '        If value Is Nothing Then Exit Property
    '        If mvarEntries Is Nothing Then mvarEntries = New Collections.Generic.List(Of EventEntry)
    '        mvarEntries.Add(value)
    '    End Set
    'End Property
    Public Sub Clear()
        If mvarEntries Is Nothing Then
            mvarEntries = New Collections.Generic.List(Of EventEntry)
            Exit Sub
        End If
        mvarEntries.Clear()
    End Sub
    Public ReadOnly Property GetAll() As EventEntry()
        Get
            If mvarEntries IsNot Nothing Then
                Return mvarEntries.ToArray
            Else
                Return Nothing
            End If
        End Get
    End Property
    Public ReadOnly Property GetAll_Collection() As Collections.Generic.List(Of EventEntry)
        Get
            Return mvarEntries
        End Get
    End Property

    Public ReadOnly Property Get_Exact_Date(ByVal GetDate As Date) As EventEntry()
        Get
            If mvarEntries Is Nothing Then Return Nothing
            If mvarEntries.Count < 1 Then Return Nothing
            Dim arr() As EventEntry = mvarEntries.ToArray
            Dim ret As New Collections.Generic.List(Of EventEntry)
            Dim AddIt As Boolean = False

            For Each entry As EventEntry In arr
                ' let's find the entries for that date
                If entry.Times.Count > 0 Then
                    For Each w As [When] In entry.Times
                        AddIt = False
                        If w.StartTime.Date = GetDate.Date Then
                            AddIt = True
                        ElseIf w.EndTime.Date = GetDate.Date Then
                            AddIt = True
                        ElseIf GetDate.Date < w.StartTime.Date AndAlso w.EndTime.Date > GetDate.Date Then
                            AddIt = True
                        End If
                        If AddIt Then
                            ret.Add(entry)
                        End If
                    Next
                End If
            Next
            Return ret.ToArray
        End Get
    End Property
    Public ReadOnly Property Get_Exact_Date_Range(ByVal StartDate As Date, ByVal EndDate As Date) As EventEntry()
        Get
            If mvarEntries Is Nothing Then Return Nothing
            If mvarEntries.Count < 1 Then Return Nothing
            Dim arr() As EventEntry = mvarEntries.ToArray
            Dim ret As New Collections.Generic.List(Of EventEntry)
            Dim AddIt As Boolean = False

            For Each entry As EventEntry In arr
                AddIt = False
                If entry.Times.Count > 0 Then
                    For Each w As [When] In entry.Times
                        'hs.writelog(IFACE_NAME, "Entry: " & entry.Title.Text & " Start:" & w.StartTime.ToString & ", End:" & w.EndTime.ToString)
                        If w.StartTime.Date >= StartDate.Date AndAlso w.EndTime.Date <= EndDate.Date Then
                            'hs.writelog(IFACE_NAME, "Entry: " & entry.Title.Text & " Both dates in range of " & StartDate.Date.ToString & "/" & EndDate.Date.ToString)
                            ' Both dates are in the range.
                            AddIt = True
                        ElseIf w.StartTime.Date >= StartDate.Date AndAlso w.StartTime.Date <= EndDate.Date Then
                            'hs.writelog(IFACE_NAME, "Entry: " & entry.Title.Text & " START is in range of " & StartDate.Date.ToString & "/" & EndDate.Date.ToString)
                            ' Appointments that START within the date range.
                            AddIt = True
                        ElseIf w.EndTime.Date >= StartDate.Date AndAlso w.EndTime.Date <= EndDate.Date Then
                            'hs.writelog(IFACE_NAME, "Entry: " & entry.Title.Text & " END is in range of " & StartDate.Date.ToString & "/" & EndDate.Date.ToString)
                            ' Appointments that END within the date range.
                            AddIt = True
                        End If
                        If AddIt Then
                            ret.Add(entry)
                        End If
                    Next
                End If
            Next
            Return ret.ToArray
        End Get
    End Property
    Public ReadOnly Property Get_DateTime_Range(ByVal FindStart As Date, ByVal FindEnd As Date) As EventEntry()
        Get
            If mvarEntries Is Nothing Then Return Nothing
            If mvarEntries.Count < 1 Then Return Nothing
            Dim arr() As EventEntry = mvarEntries.ToArray
            Dim ret As New Collections.Generic.List(Of EventEntry)
            Dim AddIt As Boolean = False

            For Each entry As EventEntry In arr
                AddIt = False
                If entry.Times.Count > 0 Then
                    For Each w As [When] In entry.Times
                        If w.StartTime >= FindStart AndAlso w.EndTime <= FindEnd Then
                            ' Both dates are in the range.
                            AddIt = True
                        ElseIf w.StartTime >= FindStart AndAlso w.StartTime <= FindEnd Then
                            ' Appointments that START within the date range.
                            AddIt = True
                        ElseIf w.EndTime >= FindStart AndAlso w.EndTime <= FindEnd Then
                            ' Appointments that END within the date range.
                            AddIt = True
                        End If
                        If AddIt Then
                            ret.Add(entry)
                        End If
                    Next
                End If
            Next
            Return ret.ToArray
        End Get
    End Property

    Public ReadOnly Property Get_StartingOn_Exact_Date(ByVal GetDate As Date) As EventEntry()
        Get
            If mvarEntries Is Nothing Then Return Nothing
            If mvarEntries.Count < 1 Then Return Nothing
            Dim arr() As EventEntry = mvarEntries.ToArray
            Dim ret As New Collections.Generic.List(Of EventEntry)

            For Each entry As EventEntry In arr
                ' let's find the entries for that date
                If entry.Times.Count > 0 Then
                    For Each w As [When] In entry.Times
                        If w.StartTime.Date = GetDate.Date Then
                            ret.Add(entry)
                        End If
                    Next
                End If
            Next
            Return ret.ToArray
        End Get
    End Property


    Public ReadOnly Property Get_by_Author(ByVal Author As String) As EventEntry()
        Get
            If mvarEntries Is Nothing Then Return Nothing
            If mvarEntries.Count < 1 Then Return Nothing
            If String.IsNullOrEmpty(Author) Then Return Nothing

            Dim arr() As EventEntry = mvarEntries.ToArray
            Dim ret As New Collections.Generic.List(Of EventEntry)

            For Each entry As EventEntry In arr
                If entry.Authors IsNot Nothing Then
                    If entry.Authors.Count > 0 Then
                        For x As Integer = 0 To entry.Authors.Count - 1
                            If entry.Authors(x).Name.Trim.ToLower = Author.Trim.ToLower Then
                                ret.Add(entry)
                            ElseIf entry.Authors(x).Email.Trim.ToLower = Author.Trim.ToLower Then
                                ret.Add(entry)
                            End If
                        Next
                    End If
                End If
            Next
            Return ret.ToArray
        End Get
    End Property
    Public ReadOnly Property Get_by_Participant(ByVal ParticipantEmail As String) As EventEntry()
        Get
            If mvarEntries Is Nothing Then Return Nothing
            If mvarEntries.Count < 1 Then Return Nothing
            If String.IsNullOrEmpty(ParticipantEmail) Then Return Nothing

            Dim arr() As EventEntry = mvarEntries.ToArray
            Dim ret As New Collections.Generic.List(Of EventEntry)
            'Dim Author As String = ""

            For Each entry As EventEntry In arr
                'Try
                '    Author = entry.Authors(0).Email.Trim.ToLower
                'Catch ex As Exception
                'End Try
                If entry.Participants IsNot Nothing Then
                    If entry.Participants.Count > 0 Then
                        For x As Integer = 0 To entry.Participants.Count - 1
                            If entry.Participants(x).Email.Trim.ToLower = ParticipantEmail.Trim.ToLower Then ' AndAlso _
                                'entry.Participants(x).Email.Trim.ToLower <> Author Then
                                ret.Add(entry)
                            End If
                        Next
                    End If
                End If
            Next
            Return ret.ToArray
        End Get
    End Property
    Public Sub InitRecords(ByVal DaysAhead As Integer, Optional ByVal DaysBack As Integer = 0)
        If (DaysAhead < 1) And (DaysBack < 0) Then
            Me.Clear()
            Exit Sub
        End If
        If DaysBack > 0 Then
            DaysBack = DaysBack * -1
        End If
        Dim s As String = ""
        Dim StartDate As Date
        Dim EndDate As Date
        Try
            StartDate = Now.AddDays(DaysBack).Date
        Catch ex As Exception
            Throw New Exception("Error calculating days back: " & ex.Message)
            Exit Sub
        End Try
        Try
            EndDate = Now.AddDays(DaysAhead).Date
        Catch ex As Exception
            Throw New Exception("Error calculating days ahead: " & ex.Message)
            Exit Sub
        End Try
        'hs.WriteLog(IFACE_NAME & " Debug", "Calling Refresh with " & StartDate.ToString & " and " & EndDate.ToString)
        s = RefreshFeed(StartDate, EndDate)
        If Not String.IsNullOrEmpty(s) Then
            Throw New Exception(s)
        End If
    End Sub
    Public Sub InitRecords(ByVal ToDate As Date)
        If ToDate.Date < Now.Date Then
            Me.Clear()
            Exit Sub
        End If
        Dim s As String = ""
        'hs.WriteLog(IFACE_NAME & " Debug", "Calling Refresh with " & ToDate.ToString)
        s = RefreshFeed(Now.Date, ToDate.Date)
        If Not String.IsNullOrEmpty(s) Then
            Throw New Exception(s)
        End If
    End Sub


    Private Function RefreshFeed(ByRef qStartTime As Date, ByRef qEndTime As Date) As String
        Const calendarURI As String = "http://www.google.com/calendar/feeds/default/private/full?singleevents=true&orderby=starttime&showhidden=true"
        'Const calendarURI As String = "http://www.google.com/calendar/feeds/default/private/full?orderby=starttime&showhidden=true"
        Dim query As New EventQuery()
        Dim AddIt As Boolean = False

        If service Is Nothing Then Return "Service not initialized."

        If Now.Subtract(LastAuth).TotalHours > 12 Then
            If mvarUser.UserName IsNot Nothing AndAlso mvarUser.UserName.Length > 0 Then
                Try
                    If String.IsNullOrEmpty(AuthToken) Then
                        service.setUserCredentials(mvarUser.UserName, mvarUser.Password)
                        AuthToken = service.QueryAuthenticationToken
                        LastAuth = Now
                    Else
                        service.SetAuthenticationToken(AuthToken)
                        service.QueryAuthenticationToken()
                        LastAuth = Now
                    End If
                Catch ex As Exception
                    ComStatus = False
                    Return "Exception setting user credentials for " & mvarUser.UserName & " : " & ex.Message
                End Try
            End If
        End If


        query.Uri = New Uri(calendarURI)

        query.StartTime = qStartTime '#1/1/1980#
        query.EndTime = qEndTime

        Dim calFeed As EventFeed
        Try
            calFeed = TryCast(service.Query(query), EventFeed)
        Catch ex As Exception
            ComStatus = False
            Return "Exception retrieving calendar entries: " & ex.Message & " for " & mvarUser.UserName
        End Try

        ' now populate the entries
        Try
            Me.Clear()
            While calFeed IsNot Nothing AndAlso calFeed.Entries.Count > 0
                'hs.WriteLog(IFACE_NAME & " Debug", "Getting " & calFeed.Entries.Count.ToString & " for " & mvarUser.FriendlyName)
                For Each entry As EventEntry In calFeed.Entries
                    ' Make sure we only get our entries.
                    'If entry.Authors IsNot Nothing Then
                    'If entry.Authors.Count >= 1 Then
                    'If entry.Authors(0).Email.Trim.ToLower = mvarUser.UserName.Trim.ToLower Then
                    AddIt = False
                    If entry.Participants IsNot Nothing Then
                        If entry.Participants.Count > 0 Then
                            For it As Integer = 0 To entry.Participants.Count - 1
                                If entry.Participants(it).Email IsNot Nothing Then
                                    If entry.Participants(it).Email.Trim.ToLower = mvarUser.UserName.Trim.ToLower Then
                                        AddIt = True
                                        Exit For
                                    End If
                                End If
                            Next
                            If AddIt Then
                                mvarEntries.Add(entry)
                            End If
                        End If
                    End If
                Next
                ' just query the same query again.
                If calFeed.NextChunk IsNot Nothing Then
                    query.Uri = New Uri(calFeed.NextChunk)
                    calFeed = TryCast(service.Query(query), EventFeed)
                Else
                    calFeed = Nothing
                End If
            End While
        Catch ex As Exception
            ComStatus = False
            Return "Exception retrieving calendar entries for " & mvarUser.UserName & ": " & ex.Message
        End Try

        'hs.WriteLog(IFACE_NAME & " Debug", "Retrieved " & mvarEntries.Count.ToString & " entries for " & mvarUser.FriendlyName)

        ComStatus = True
        Return ""

    End Function

    Private Function GetEntry(ByVal FeedURL As String) As Google.GData.Calendar.EventEntry
        If String.IsNullOrEmpty(FeedURL) Then Return Nothing

        Dim query As New EventQuery()
        Dim AddIt As Boolean = False

        If service Is Nothing Then Return Nothing

        If Now.Subtract(LastAuth).TotalHours > 12 Then
            If mvarUser.UserName IsNot Nothing AndAlso mvarUser.UserName.Length > 0 Then
                Try
                    If String.IsNullOrEmpty(AuthToken) Then
                        service.setUserCredentials(mvarUser.UserName, mvarUser.Password)
                        AuthToken = service.QueryAuthenticationToken
                        LastAuth = Now
                    Else
                        service.SetAuthenticationToken(AuthToken)
                        service.QueryAuthenticationToken()
                        LastAuth = Now
                    End If
                Catch ex As Exception
                    Return Nothing
                End Try
            End If
        End If

        'If mvarUser.UserName IsNot Nothing AndAlso mvarUser.UserName.Length > 0 Then
        '    Try
        '        service.setUserCredentials(mvarUser.UserName, mvarUser.Password)
        '    Catch ex As Exception
        '        Return Nothing
        '    End Try
        'End If


        query.Uri = New Uri(FeedURL)

        Dim calFeed As EventFeed
        Try
            calFeed = TryCast(service.Query(query), EventFeed)
        Catch ex As Exception
            Return Nothing
        End Try

        ' now populate the entries
        Try
            If calFeed IsNot Nothing AndAlso calFeed.Entries.Count = 1 Then
                Return calFeed.Entries(0)
            End If
        Catch ex As Exception
            Return Nothing
        End Try

        Return Nothing

    End Function

    'Private Sub NewFeed_Handler(ByVal Sender As Object, ByVal e As Google.GData.Client.ServiceEventArgs) Handles service.NewFeed
    '    Console.WriteLine("NewFeed_Handler")
    'End Sub

    <Serializable()> _
    Public Class HSGCalEntry

        Private mvarEntry As Google.GData.Calendar.EventEntry
        Private mvarUser As GoogleCalendar.UserPair

        Public Sub New(ByVal GEntry As Google.GData.Calendar.EventEntry, _
                       ByVal User As GoogleCalendar.UserPair)
            MyBase.new()
            If GEntry Is Nothing Then
                Throw New Exception("A valid Google Calendar Entry Object must be provided when initializing this object.")
                Exit Sub
            End If
            If User Is Nothing Then
                Throw New Exception("A valid Google Calendar User Object must be provided when initializing this object.")
                Exit Sub
            End If
            mvarEntry = GEntry
            mvarUser = User
        End Sub

        Public ReadOnly Property GoogleEntry() As Google.GData.Calendar.EventEntry
            Get
                Return mvarEntry
            End Get
        End Property

        Public ReadOnly Property User() As GoogleCalendar.UserPair
            Get
                Return mvarUser
            End Get
        End Property

        Public ReadOnly Property Title() As String
            Get

                Try
                    If mvarEntry.Title.Type <> AtomTextConstructType.text Then
                        Return stripHTML(mvarEntry.Title.Text)
                    Else
                        Return mvarEntry.Title.Text
                    End If
                Catch ex As Exception
                    Return ""
                End Try
            End Get
        End Property
        Public ReadOnly Property TitleHTML() As String
            Get

                Try
                    If mvarEntry.Title.Type <> AtomTextConstructType.html Then
                        Return stripHTML(mvarEntry.Title.Text)
                    Else
                        Return mvarEntry.Title.Text
                    End If
                Catch ex As Exception
                    Return ""
                End Try
            End Get
        End Property
        Public ReadOnly Property Author() As String
            Get
                If mvarEntry.Authors IsNot Nothing Then
                    Return mvarEntry.Authors(0).Email
                Else
                    Return ""
                End If
            End Get
        End Property
        Public ReadOnly Property Category() As String()
            Get
                Dim cat As New Collections.Generic.List(Of String)
                If mvarEntry.Categories IsNot Nothing Then
                    For Each c As Google.GData.Client.AtomCategory In mvarEntry.Categories
                        If c IsNot Nothing Then
                            If Not String.IsNullOrEmpty(c.Label) Then
                                cat.Add(c.Label)
                            End If
                        End If
                    Next
                    Return cat.ToArray
                Else
                    Return Nothing
                End If
            End Get
        End Property
        Public ReadOnly Property CategoryList() As String
            Get
                Dim cat As New Collections.Generic.List(Of String)
                If mvarEntry.Categories IsNot Nothing Then
                    For Each c As Google.GData.Client.AtomCategory In mvarEntry.Categories
                        If c IsNot Nothing Then
                            If Not String.IsNullOrEmpty(c.Label) Then
                                cat.Add(c.Label)
                            End If
                        End If
                    Next
                    Return Join(cat.ToArray, ",")
                Else
                    Return ""
                End If
            End Get
        End Property
        Public ReadOnly Property Description() As String
            Get
                If mvarEntry.Content IsNot Nothing Then
                    If Not String.IsNullOrEmpty(mvarEntry.Content.Content) Then
                        Return mvarEntry.Content.Content
                    Else
                        Return ""
                    End If
                Else
                    Return ""
                End If
            End Get
        End Property
        Public ReadOnly Property EventLink() As String
            Get
                Return mvarEntry.Links(0).HRef.Content
            End Get
        End Property
        Public ReadOnly Property Status() As Google.GData.Calendar.EventEntry.EventStatus
            Get
                Return mvarEntry.Status
            End Get
        End Property
        Public ReadOnly Property StatusText() As String
            Get
                Select Case mvarEntry.Status.Value
                    Case Google.GData.Calendar.EventEntry.EventStatus.CANCELED_VALUE
                        Return "Canceled"
                    Case Google.GData.Calendar.EventEntry.EventStatus.CONFIRMED_VALUE
                        Return "Confirmed"
                    Case Google.GData.Calendar.EventEntry.EventStatus.TENTATIVE_VALUE
                        Return "Tentative"
                    Case Else
                        Return "(Unknown)"
                End Select
            End Get
        End Property
        'Public ReadOnly Property Recurrence() As String
        '    Get
        '        If mvarEntry.Recurrence IsNot Nothing Then
        '            If Not String.IsNullOrEmpty(mvarEntry.Recurrence.Value) Then
        '                Return mvarEntry.Recurrence.Value
        '            Else
        '                Return "None"
        '            End If
        '        Else
        '            Return "None"
        '        End If
        '    End Get
        'End Property
        <Serializable()> _
        Public Enum enumFreeBusy
            Free = 1
            Busy = 2
            Undefined = 3
        End Enum
        Public ReadOnly Property FreeBusy() As enumFreeBusy
            Get
                Select Case mvarEntry.EventTransparency.Value
                    Case Google.GData.Calendar.EventEntry.Transparency.OPAQUE_VALUE
                        Return enumFreeBusy.Busy
                    Case Google.GData.Calendar.EventEntry.Transparency.TRANSPARENT_VALUE
                        Return enumFreeBusy.Free
                    Case Else
                        Return enumFreeBusy.Undefined
                End Select
            End Get
        End Property
        <Serializable()> _
        Public Enum enumVisibility
            vDefault = 1
            vConfidential = 2
            vPrivate = 3
            vPublic = 4
            vUnknown = 5
        End Enum
        Public ReadOnly Property Visibility() As enumVisibility
            Get
                Select Case mvarEntry.EventVisibility.Value
                    Case Google.GData.Calendar.EventEntry.Visibility.CONFIDENTIAL_VALUE
                        Return enumVisibility.vConfidential
                    Case Google.GData.Calendar.EventEntry.Visibility.DEFAULT_VALUE
                        Return enumVisibility.vDefault
                    Case Google.GData.Calendar.EventEntry.Visibility.PRIVATE_VALUE
                        Return enumVisibility.vPrivate
                    Case Google.GData.Calendar.EventEntry.Visibility.PUBLIC_VALUE
                        Return enumVisibility.vPublic
                    Case Else
                        Return enumVisibility.vUnknown
                End Select
            End Get
        End Property
        Public ReadOnly Property StartDT() As Date
            Get
                Return mvarEntry.Times(0).StartTime
            End Get
        End Property
        Public ReadOnly Property EndDT() As Date
            Get
                Return mvarEntry.Times(0).EndTime
            End Get
        End Property
        <Serializable()> _
        Public Enum enumRemindMethod
            All = 1
            Alert = 2
            eMail = 3
            SMS = 4
            None = 5
            Unspecified = 6
        End Enum
        <Serializable()> _
        Public Class Remind
            Public Sub New()
                MyBase.New()
            End Sub
            Public RemindWhen As TimeSpan
            Public RemindMethod As enumRemindMethod = enumRemindMethod.Unspecified
        End Class
        Public ReadOnly Property Reminders(ByVal GC As GoogleCalendar) As Remind()
            Get
                'Try
                '    hs.WriteLog(IFACE_NAME & " Debug", "Recurrence: " & mvarEntry.Recurrence.Value)
                'Catch ex2 As Exception
                '    hs.WriteLog(IFACE_NAME & " Debug", "Recurrence is NOTHING")
                'End Try
                'Try
                '    hs.WriteLog(IFACE_NAME & " Debug", "Alt:" & mvarEntry.AlternateUri.Content.ToString)
                'Catch ex As Exception
                'End Try

                'Try
                '    Dim EXL As Google.GData.Client.ExtensionList = mvarEntry.ExtensionElements
                '    hs.WriteLog(IFACE_NAME & " Debug", "EXL:" & EXL.Count.ToString)
                '    Dim ixl As Google.GData.Client.IExtensionElementFactory
                '    For i As Integer = 0 To EXL.Count - 1
                '        ixl = EXL.Item(i)
                '        If ixl IsNot Nothing Then
                '            hs.WriteLog(IFACE_NAME & " Debug", "Ext " & i.ToString & ":" & ixl.XmlName)
                '        End If
                '    Next
                'Catch ex As Exception
                'End Try

                'Return Nothing


                Dim colR As New Collections.Generic.List(Of Remind)
                Dim exCol As Google.GData.Extensions.ExtensionCollection(Of Reminder)
                Dim RM As Remind
                Dim ReminderOK As Boolean = False
                Dim RemindersOK As Boolean = False

                If mvarEntry.Recurrence IsNot Nothing Then
                    If Not String.IsNullOrEmpty(mvarEntry.Recurrence.Value) Then
                        If mvarEntry.OriginalEvent Is Nothing Then
                            If mvarEntry.Reminder IsNot Nothing Then
                                GoTo SingleRemind
                            Else
                                Return Nothing
                            End If
                        End If
                        If Not String.IsNullOrEmpty(mvarEntry.OriginalEvent.Href) Then
                            If GC Is Nothing Then Return Nothing
                            Dim REV As Google.GData.Calendar.EventEntry
                            REV = GC.GetEntry(mvarEntry.OriginalEvent.Href)
                            If REV IsNot Nothing Then
                                Dim eREV As New GoogleCalendar.HSGCalEntry(REV, mvarUser)
                                Return eREV.Reminders(Nothing)
                            Else
                                Return Nothing
                            End If
                        Else
                            If mvarEntry.Reminder IsNot Nothing Then
                                GoTo SingleRemind
                            Else
                                Return Nothing
                            End If
                        End If
                        'Return Nothing
                    End If
                End If

                Try
                    exCol = mvarEntry.Reminders
                    If exCol IsNot Nothing Then
                        If exCol.Count > 0 Then
                            RemindersOK = True
                        Else
                            RemindersOK = False
                        End If
                    Else
                        RemindersOK = False
                    End If
                Catch ex As Exception
                    RemindersOK = False
                End Try
                Try
                    If RemindersOK Then
                        Dim r As Google.GData.Extensions.Reminder
                        If mvarEntry.Reminders IsNot Nothing Then
                            exCol = mvarEntry.Reminders
                            For ir As Integer = 0 To exCol.Count - 1
                                'For Each r As Google.GData.Extensions.Reminder In mvarEntry.Reminders
                                r = exCol.Item(ir)
                                If r IsNot Nothing Then
                                    RM = New Remind
                                    RM.RemindWhen = New TimeSpan(r.Days, r.Hours, r.Minutes, 0)
                                    Select Case r.Method
                                        Case Reminder.ReminderMethod.alert
                                            RM.RemindMethod = enumRemindMethod.Alert
                                        Case Reminder.ReminderMethod.all
                                            RM.RemindMethod = enumRemindMethod.All
                                        Case Reminder.ReminderMethod.email
                                            RM.RemindMethod = enumRemindMethod.eMail
                                        Case Reminder.ReminderMethod.none
                                            RM.RemindMethod = enumRemindMethod.None
                                        Case Reminder.ReminderMethod.sms
                                            RM.RemindMethod = enumRemindMethod.SMS
                                        Case Else
                                            RM.RemindMethod = enumRemindMethod.Unspecified
                                    End Select
                                    colR.Add(RM)
                                End If
                            Next
                            'Return colR.ToArray
                        Else
                            'Return Nothing
                        End If
                    End If
                Catch ex As Exception
                    hs.WriteLog(IFACE_NAME & " Error", "Exception (" & Erl.ToString & ") accessing Reminders property for calendar: " & ex.Message)
                End Try
                GoTo AddEmUp
SingleRemind:
                Try
                    If mvarEntry.Reminder IsNot Nothing Then
                        ReminderOK = True
                    Else
                        ReminderOK = False
                    End If
                Catch ex As Exception
                    ReminderOK = False
                End Try
                Try
                    If ReminderOK Then
                        If mvarEntry.Reminder IsNot Nothing Then
                            RM = New Remind
                            RM.RemindWhen = New TimeSpan(mvarEntry.Reminder.Days, mvarEntry.Reminder.Hours, mvarEntry.Reminder.Minutes, 0)
                            Select Case mvarEntry.Reminder.Method
                                Case Reminder.ReminderMethod.alert
                                    RM.RemindMethod = enumRemindMethod.Alert
                                Case Reminder.ReminderMethod.all
                                    RM.RemindMethod = enumRemindMethod.All
                                Case Reminder.ReminderMethod.email
                                    RM.RemindMethod = enumRemindMethod.eMail
                                Case Reminder.ReminderMethod.none
                                    RM.RemindMethod = enumRemindMethod.None
                                Case Reminder.ReminderMethod.sms
                                    RM.RemindMethod = enumRemindMethod.SMS
                                Case Else
                                    RM.RemindMethod = enumRemindMethod.Unspecified
                            End Select
                            colR.Add(RM)
                            'Return colR.ToArray
                        Else
                            'Return Nothing
                        End If
                    End If
                Catch ex As Exception
                    hs.WriteLog(IFACE_NAME & " Error", "Exception (" & Erl.ToString & ") accessing REMINDER property for calendar: " & ex.Message)
                End Try
AddEmUp:

                If colR IsNot Nothing Then
                    If colR.Count > 0 Then
                        Return colR.ToArray
                    Else
                        Return Nothing
                    End If
                Else
                    Return Nothing
                End If
            End Get
        End Property
        Public ReadOnly Property Location() As String
            Get
                Dim colLoc As New Collections.Generic.List(Of String)
                If mvarEntry.Locations IsNot Nothing Then
                    For Each w As Google.GData.Extensions.Where In mvarEntry.Locations
                        If Not String.IsNullOrEmpty(w.ValueString) Then
                            colLoc.Add(w.ValueString)
                        End If
                    Next
                    Return Join(colLoc.ToArray, ", ")
                Else
                    Return ""
                End If
            End Get
        End Property
        <Serializable()> _
        Public Enum enumAttendeeStatus
            Accepted = 1
            Declined = 2
            Invited = 3
            Tentative = 4
            Undefined = 9
        End Enum
        <Serializable()> _
        Public Enum enumAttendeeType
            AttendanceOptional = 1
            AttendanceRequired = 2
            Undefined = 9
        End Enum
        <Serializable()> _
        Public Class GCalAttendee
            Public Sub New()
                MyBase.New()
            End Sub
            Public eMail As String
            Public Type As enumAttendeeType
            Public Status As enumAttendeeStatus
        End Class

        Public ReadOnly Property AttendeesFilt(ByVal Filter As String) As GCalAttendee()
            Get
                Dim colATT As New Collections.Generic.List(Of GCalAttendee)
                Dim ATT() As GCalAttendee
                ATT = Me.Attendees
                If ATT Is Nothing Then Return Nothing
                If ATT.Length = 0 Then Return Nothing
                For x As Integer = 0 To ATT.Length - 1
                    If ATT(x).eMail.Trim.ToLower = Filter.Trim.ToLower Then
                        ' Do Nothing
                    Else
                        colATT.Add(ATT(x))
                    End If
                Next
                If colATT Is Nothing Then Return Nothing
                If colATT.Count < 1 Then Return Nothing
                Return colATT.ToArray
            End Get
        End Property

        Public ReadOnly Property Attendees() As GCalAttendee()
            Get
                Dim colAtt As New Collections.Generic.List(Of GCalAttendee)
                Dim ATT As GCalAttendee
                'if mvarentry.Contributors
                Try
                    If mvarEntry.Participants IsNot Nothing Then
                        For Each w As Google.GData.Extensions.Who In mvarEntry.Participants
                            If w IsNot Nothing Then
                                ATT = New GCalAttendee
                                ATT.eMail = w.Email
                                If w.Attendee_Type IsNot Nothing AndAlso Not String.IsNullOrEmpty(w.Attendee_Type.Value) Then
                                    Select Case w.Attendee_Type.Value
                                        Case Who.AttendeeType.EVENT_REQUIRED
                                            ATT.Type = enumAttendeeType.AttendanceRequired
                                        Case Who.AttendeeType.EVENT_OPTIONAL
                                            ATT.Type = enumAttendeeType.AttendanceOptional
                                        Case Else
                                            ATT.Type = enumAttendeeType.Undefined
                                    End Select
                                Else
                                    ATT.Type = enumAttendeeType.Undefined
                                End If
                                If w.Attendee_Status IsNot Nothing AndAlso Not String.IsNullOrEmpty(w.Attendee_Status.Value) Then
                                    Select Case w.Attendee_Status.Value
                                        Case Who.AttendeeStatus.EVENT_ACCEPTED
                                            ATT.Status = enumAttendeeStatus.Accepted
                                        Case Who.AttendeeStatus.EVENT_DECLINED
                                            ATT.Status = enumAttendeeStatus.Declined
                                        Case Who.AttendeeStatus.EVENT_INVITED
                                            ATT.Status = enumAttendeeStatus.Invited
                                        Case Who.AttendeeStatus.EVENT_TENTATIVE
                                            ATT.Status = enumAttendeeStatus.Tentative
                                        Case Else
                                            ATT.Status = enumAttendeeStatus.Undefined
                                    End Select
                                Else
                                    ATT.Status = enumAttendeeStatus.Undefined
                                End If
                                colAtt.Add(ATT)
                            End If
                        Next
                        Return colAtt.ToArray
                    Else
                        Return Nothing
                    End If
                Catch ex As Exception
                    hs.WriteLog(IFACE_NAME & " Error", "Exception (" & Erl.ToString & ") in HSGCal Attendees property: " & ex.Message)
                End Try
                Return Nothing
            End Get
        End Property



    End Class

End Class
