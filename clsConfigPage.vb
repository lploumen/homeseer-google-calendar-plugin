Imports System.Text
Imports VB = Microsoft.VisualBasic


Public Class clsConfigPage

    ' ------------------------------------------------------------------------------------------------------------------------------
    '
    '   Provides a configuration web page for the plug-in.
    '
    ' ------------------------------------------------------------------------------------------------------------------------------


    Public link As String       ' Actual link string, such as:    my_page
    Public linktext As String   ' Display text for link, such as: "Security Status"  <-- Text of the HS generated link button.
    Public page_title As String ' Title of the web page, such as: "Acme Security Panel Status"
    Public lPairs As Long       ' Number of name/value pairs after GetFormData
    Public tPair() As Pair      ' Name/Value pairs array, populated by GetFormData

    Public Function GenPage(ByRef lnk As String) As String
        Dim p As New StringBuilder
        Dim Row As Short = 0
        Dim RowClass(1) As String

        RowClass(0) = "tableroweven"
        RowClass(1) = "tablerowodd"

        p.Append(HTML_NewLine)

        p.Append(HTML_StartForm)
        p.Append(HTML_StartTable(0, 0, 60))

        p.Append(HTML_StartRow)
        p.Append(HTML_StartCell("tableheader", 8, , True))
        p.Append("Google Calendar Plug-In Configuration")
        p.Append(HTML_EndCell)
        p.Append(HTML_EndRow)

        p.Append(HTML_StartRow)
        p.Append(HTML_StartCell("tablecolumn", 2, , True))
        p.Append("User Name (eMail Address)")
        p.Append(HTML_EndCell)
        p.Append(HTML_StartCell("tablecolumn", 2, , True))
        p.Append("Friendly Name")
        p.Append(HTML_EndCell)
        p.Append(HTML_StartCell("tablecolumn", 2, , True))
        p.Append("Password")
        p.Append(HTML_EndCell)
        p.Append(HTML_StartCell("tablecolumn", 2, , True))
        p.Append("Actions")
        p.Append(HTML_EndCell)
        p.Append(HTML_EndRow)

        ' An HTML form field having a name and a value all makes sense until you get to a form button.
        ' In the case of a button, the VALUE is what is displayed on the button, so keep that in mind 
        '   that a button is backwards from what logic would dictate!

        Dim g As GoogleCalendar
        'Dim u As GoogleCalendar.UserPair

        If Users IsNot Nothing Then
            If Users.Count > 0 Then
                For x As Integer = 0 To Users.Count - 1
                    g = Users.GetByIndex(x)
                    If g Is Nothing Then Continue For
                    p.Append(HTML_StartRow)
                    p.Append(HTML_StartCell(RowClass(Row), 2, ALIGN_LEFT))
                    p.Append(g.UserName)
                    p.Append(AddHidden("user_id_" & x.ToString, g.UserName))
                    p.Append(HTML_EndCell)

                    p.Append(HTML_StartCell(RowClass(Row), 2, ALIGN_LEFT))
                    p.Append(FormTextBox("", "friendly_" & x.ToString, g.User.FriendlyName, 12))
                    p.Append(HTML_EndCell)

                    p.Append(HTML_StartCell(RowClass(Row), 2, ALIGN_LEFT))
                    If g.User.ShowPass Then
                        p.Append(FormTextBox("", "password_" & x.ToString, g.User.Password, 12))
                    Else
                        p.Append(FormPasswordTextBox("", "password_" & x.ToString, "xxxxxxxxxx", 12))
                    End If
                    p.Append(HTML_EndCell)

                    p.Append(HTML_StartCell(RowClass(Row), 1, ALIGN_LEFT))
                    p.Append(FormCheckBox("Show Password", "show_pass_" & x.ToString, IIf(g.User.ShowPass, "True", "False"), IIf(g.User.ShowPass, True, False), True))
                    p.Append(HTML_EndCell)

                    p.Append(HTML_StartCell(RowClass(Row), 1, ALIGN_LEFT))
                    p.Append(FormButton("save_" & x.ToString, "Save", "Save changes made to user " & g.UserName, , , True))
                    p.Append("&nbsp;")
                    p.Append(FormButton("delete_" & x.ToString, "Delete", "Delete user " & g.UserName, , , True))
                    p.Append(HTML_EndCell)
                    p.Append(HTML_EndRow)

                    g.User.ShowPass = False

                    Row = Row Xor 1
                Next
            End If
        End If

        p.Append(HTML_StartRow)
        p.Append(HTML_StartCell(RowClass(Row), 8, ALIGN_LEFT))
        p.Append("&nbsp;")
        p.Append(HTML_EndCell)
        p.Append(HTML_EndRow)
        p.Append(HTML_StartRow)
        p.Append(HTML_StartCell(RowClass(Row), 8, ALIGN_LEFT))
        p.Append("&nbsp;New User:")
        p.Append(HTML_EndCell)
        p.Append(HTML_EndRow)
        Row = Row Xor 1

        p.Append(HTML_StartRow)
        p.Append(HTML_StartCell(RowClass(Row), 2, ALIGN_LEFT))
        p.Append(FormTextBox("", "user_new", "Joe.Sample@GMail.com", 30))
        p.Append(HTML_EndCell)

        p.Append(HTML_StartCell(RowClass(Row), 2, ALIGN_LEFT))
        p.Append(FormTextBox("", "new_friendly", "Joe", 12))
        p.Append(HTML_EndCell)

        p.Append(HTML_StartCell(RowClass(Row), 2, ALIGN_LEFT))
        p.Append(FormTextBox("", "new_password", "password", 12))
        p.Append(HTML_EndCell)

        p.Append(HTML_StartCell(RowClass(Row), 2, ALIGN_LEFT))
        p.Append(FormButton("add_new", "Add", "Add New User", , , True))
        p.Append(HTML_EndCell)
        p.Append(HTML_EndRow)

        Row = Row Xor 1

        p.Append(HTML_EndTable)

        p.Append(HTML_NewLine)

        p.Append(HTML_StartTable(0, 0, 60, ALIGN_LEFT))

        p.Append(HTML_StartRow)
        p.Append(HTML_StartCell("tablecolumn", 8, , True))
        p.Append("Additional Options")
        p.Append(HTML_EndCell)
        p.Append(HTML_EndRow)

        Row = 0
        p.Append(HTML_StartRow)
        p.Append(HTML_StartCell(RowClass(Row), 2, ALIGN_RIGHT))
        p.Append("Number of days AHEAD to retrieve appointments when refreshing.")
        p.Append(HTML_EndCell)
        p.Append(HTML_StartCell(RowClass(Row), 1, ALIGN_LEFT))
        p.Append(FormTextBox("", "DaysAhead", DaysAhead.ToString, 1))
        p.Append(HTML_EndCell)
        p.Append(HTML_StartCell(RowClass(Row), 5, ALIGN_LEFT))
        p.Append("&nbsp;")
        p.Append(HTML_EndCell)
        p.Append(HTML_EndRow)
        Row = Row Xor 1
        p.Append(HTML_StartRow)
        p.Append(HTML_StartCell(RowClass(Row), 2, ALIGN_RIGHT))
        p.Append("Number of days BACK to retrieve appointments when refreshing.")
        p.Append(HTML_EndCell)
        p.Append(HTML_StartCell(RowClass(Row), 1, ALIGN_LEFT))
        p.Append(FormTextBox("", "DaysBehind", DaysBehind.ToString, 1))
        p.Append(HTML_EndCell)
        p.Append(HTML_StartCell(RowClass(Row), 5, ALIGN_LEFT))
        p.Append("&nbsp;")
        p.Append(HTML_EndCell)
        p.Append(HTML_EndRow)
        Row = Row Xor 1
        p.Append(HTML_StartRow)
        p.Append(HTML_StartCell(RowClass(Row), 2, ALIGN_RIGHT))
        p.Append("Refresh Interval to Retrieve Calendar Items from Google")
        p.Append(HTML_EndCell)
        p.Append(HTML_StartCell(RowClass(Row), 1, ALIGN_LEFT))
        p.Append(FormTextBox("", "Refresh", UpdateThread_Interval.TotalMinutes.ToString, 2) & " Minutes")
        p.Append(HTML_EndCell)
        p.Append(HTML_StartCell(RowClass(Row), 5, ALIGN_LEFT))
        p.Append("&nbsp;")
        p.Append(HTML_EndCell)
        p.Append(HTML_EndRow)
        Row = Row Xor 1

        p.Append(HTML_StartRow)
        p.Append(HTML_StartCell("", 8, ALIGN_LEFT))
        p.Append(FormButton("dummysave", "Save", "Save changes made to the settings", , , True))
        p.Append(HTML_EndCell)
        p.Append(HTML_EndRow)


        p.Append(HTML_EndTable)

        'UpdateThread_Interval

        p.Append(AddHidden("ref_page", Me.link))

        p.Append(HTML_EndForm)

        p.Append(HTML_NewLine & HTML_NewLine)

        Return p.ToString

    End Function

    ' put requests call here
    Public Function PagePut(ByRef data As String) As String
        Dim UserMod As New Collections.SortedList
        Dim g As GoogleCalendar
        Dim u As GoogleCalendar.UserPair
        Dim uID As Integer
        Dim sPass As String = ""
        Dim sUser As String = ""
        Dim sFriend As String = ""

        Dim ChangedUsers As Boolean = False
        Dim ChangedSetting As Boolean = False

        GetFormData(data, Me.lPairs, Me.tPair)

        'p.Append(AddHidden("user_id_" & x.ToString, g.UserName))
        'p.Append(FormTextBox("", "password_" & x.ToString, g.User.Password, 12))
        'p.Append(FormPasswordTextBox("", "password_" & x.ToString, "xxxxxxxxxx", 12))
        'p.Append(FormCheckBox("Show Password", "show_pass_" & x.ToString, IIf(g.User.ShowPass, "True", "False"), IIf(g.User.ShowPass, True, False), True))
        'p.Append(FormButton("delete_" & x.ToString, "Delete", "Delete user " & g.UserName, , , True))
        If lPairs > 0 Then
            For i As Integer = 0 To lPairs - 1
                If tPair(i).Name.Trim.ToLower.StartsWith("user_id_") Then
                    uID = CInt(Val(Mid(tPair(i).Name.Trim, 9)))
                    For x As Integer = 0 To Users.Count - 1
                        g = Users.GetByIndex(x)
                        If g IsNot Nothing Then
                            If g.UserName.Trim.ToLower = tPair(i).Value.Trim.ToLower Then
                                Try
                                    UserMod.Add("K" & uID.ToString, g.User)
                                Catch ex As Exception
                                End Try
                            End If
                        End If
                    Next
                End If
            Next
            ' UserMod collection should have all of our users by ID number...
            For i As Integer = 0 To lPairs - 1

                If tPair(i).Name.Trim.ToLower.StartsWith("password_") Then
                    uID = CInt(Val(Mid(tPair(i).Name.Trim, 10)))
                    u = Nothing
                    Try
                        u = UserMod.Item("K" & uID.ToString)
                    Catch ex As Exception
                        u = Nothing
                    End Try
                    If u IsNot Nothing Then
                        sPass = tPair(i).Value.Trim.ToLower
                        If sPass <> "xxxxxxxxxx" Then
                            If sPass.Trim <> u.Password.Trim Then
                                ChangedUsers = True
                                u.Password = sPass.Trim
                            End If
                        End If
                    End If

                ElseIf tPair(i).Name.Trim.ToLower.StartsWith("friendly_") Then
                    uID = CInt(Val(Mid(tPair(i).Name.Trim, 10)))
                    u = Nothing
                    Try
                        u = UserMod.Item("K" & uID.ToString)
                    Catch ex As Exception
                        u = Nothing
                    End Try
                    If u IsNot Nothing Then
                        sFriend = tPair(i).Value.Trim
                        If sFriend.Trim <> u.FriendlyName.Trim Then
                            ChangedUsers = True
                            u.FriendlyName = sFriend.Trim
                        End If
                    End If

                ElseIf tPair(i).Name.Trim.ToLower.StartsWith("show_pass_") Then
                    uID = CInt(Val(Mid(tPair(i).Name.Trim, 11)))
                    u = Nothing
                    Try
                        u = UserMod.Item("K" & uID.ToString)
                    Catch ex As Exception
                        u = Nothing
                    End Try
                    If u IsNot Nothing Then
                        u.ShowPass = True
                    End If

                ElseIf tPair(i).Name.Trim.ToLower.StartsWith("delete_") Then
                    uID = CInt(Val(Mid(tPair(i).Name.Trim, 8)))
                    u = Nothing
                    Try
                        u = UserMod.Item("K" & uID.ToString)
                    Catch ex As Exception
                        u = Nothing
                    End Try
                    If u IsNot Nothing Then
                        ChangedUsers = True
                        sUser = u.UserName
                        u = Nothing
                        Try
                            UserMod.Remove("K" & uID.ToString)
                        Catch ex As Exception
                        End Try
                        For t As Integer = 0 To Users.Count - 1
                            u = Users.GetByIndex(t)
                            If u IsNot Nothing Then
                                If u.UserName.Trim.ToLower = sUser.Trim.ToLower Then
                                    u = Nothing
                                    Users.RemoveAt(t)
                                    Exit For
                                End If
                            End If
                        Next
                    End If


                    'p.Append(FormTextBox("", "user_new", "Joe.Sample@GMail.com", 20))
                    'p.Append(FormTextBox("", "password_new", "password", 12))
                    'p.Append(FormButton("add_new", "Add", "Add New User", , , True))
                ElseIf tPair(i).Name.Trim.ToLower = "add_new" Then
                    sUser = ""
                    sPass = ""
                    sFriend = ""
                    For a As Integer = 0 To lPairs - 1
                        If tPair(a).Name.Trim.ToLower = "user_new" Then
                            sUser = tPair(a).Value.Trim
                        ElseIf tPair(a).Name.Trim.ToLower = "new_password" Then
                            sPass = tPair(a).Value.Trim
                        ElseIf tPair(a).Name.Trim.ToLower = "new_friendly" Then
                            sFriend = tPair(a).Value.Trim
                        End If
                        If Not (String.IsNullOrEmpty(sUser) Or String.IsNullOrEmpty(sPass) Or String.IsNullOrEmpty(sFriend)) Then
                            Exit For
                        End If
                    Next
                    If Not (String.IsNullOrEmpty(sUser) Or String.IsNullOrEmpty(sPass)) Then
                        u = New GoogleCalendar.UserPair
                        u.UserName = sUser.Trim
                        u.Password = sPass.Trim
                        u.FriendlyName = sFriend.Trim
                        If String.IsNullOrEmpty(u.FriendlyName) Then
                            u.FriendlyName = u.UserName
                        End If
                        u.ShowPass = False
                        g = New GoogleCalendar(u)
                        Users.Add(u.UserName.ToLower, g)
                        ChangedUsers = True
                    End If

                    'p.Append(FormTextBox("", "DaysAhead", DaysAhead.ToString, 2))
                    'p.Append(FormTextBox("", "DaysBehind", DaysBehind.ToString, 2))
                    'p.Append(FormTextBox("", "Refresh", UpdateThread_Interval.TotalMinutes.ToString, 4) & " Minutes")
                ElseIf tPair(i).Name.Trim.ToLower = "daysahead" Then
                    Dim d As Integer
                    Try
                        d = CInt(Val(tPair(i).Value.Trim))
                    Catch ex As Exception
                        d = 0
                    End Try
                    If d < 1 Then d = 7
                    If DaysAhead <> d Then
                        DaysAhead = d
                        ChangedSetting = True
                    End If
                ElseIf tPair(i).Name.Trim.ToLower = "daysbehind" Then
                    Dim d As Integer
                    Try
                        d = CInt(Val(tPair(i).Value.Trim))
                    Catch ex As Exception
                        d = 0
                    End Try
                    If d < 0 Then d = Math.Abs(d)
                    If DaysBehind <> d Then
                        DaysBehind = d
                        ChangedSetting = True
                    End If
                ElseIf tPair(i).Name.Trim.ToLower = "refresh" Then
                    Dim m As Double
                    Try
                        m = Val(tPair(i).Value.Trim)
                    Catch ex As Exception
                        m = -1
                    End Try
                    If m < 10 Then m = 10
                    'UpdateThread_Interval()
                    Dim TS As New TimeSpan
                    TS = TimeSpan.FromMinutes(m)
                    If TS.TotalSeconds <> UpdateThread_Interval.TotalSeconds Then
                        UpdateThread_Interval = TimeSpan.FromMinutes(m)
                        ChangedSetting = True
                    End If

                End If

            Next

        End If

        If ChangedUsers Then SaveUsers()
        If ChangedSetting Then SaveSettings()
        If ChangedSetting Then UpdateThread_Trigger = True


        'LaunchApp("http://HomeSeer.com")

        Return ""

    End Function

End Class
