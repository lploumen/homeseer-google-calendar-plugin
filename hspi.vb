Imports VB = Microsoft.VisualBasic
Imports System.Text

<Serializable()> _
Public Class HSPI
    Inherits MarshalByRefObject

    ' sample plug-in for HomeSeer
    ' this control will interface between HomeSeer and a custom
    ' device supporting X10, IR, or I/O
    ' after this control is registered, the new device will appear
    ' in the interfaces tab in HomeSeer and the user may select it

    Dim AlreadyDisplayedError As Boolean

    ' capabilites of device (this OCX) (bits)
    Const CA_IO As Short = &H4 '(4) supports I/O

    ' callback constants
    ' see the HS docs and Appendix A for information on HS callbacks
    Const EV_TYPE_X10 As Short = 1
    Const EV_TYPE_LOG As Short = 2
    Const EV_TYPE_STATUS_CHANGE As Short = 4
    Const EV_TYPE_AUDIO As Short = 8
    Const EV_TYPE_X10_TRANSMIT As Short = &H10S
    Const EV_TYPE_CONFIG_CHANGE As Short = &H20S
    Const EV_TYPE_STRING_CHANGE As Short = &H40S


    ' for web page generation
    Dim lPairs As Integer ' number of name=value pairs sent by form

    Dim tPair() As Pair ' array of name=value pairs

    ' call hs.RegisterLink and passes this object
    Public link As String ' actual link string
    Public linktext As String ' display text for link
    Public page_title As String ' title of web page

    ' end web support


#Region "    Common Plug-In Interface Procedures    "

    ' ********************************** Common Interface ****************************
    '
    ' All plug-ins must support these methods and properties

    ' return capabilities of this device
    Public Function Capabilities() As Integer
        ' tell HS we support all interfaces
        Capabilities = CA_IO
    End Function

    ' return the name of your X10 device
    ' this name will appear in the HS interfaces tab
    Public ReadOnly Property Name() As String
        Get
            Return IFACE_NAME
        End Get
    End Property

    ' HomeSeer will call this function to register a callback object
    ' If events are needed (for things like IR matching, X10 events, etc.), call back
    ' to HomeSeer using this function.

    Public Sub RegisterCallback(ByRef frm As Object)
        ' call back into HS and get a reference to the HomeSeer ActiveX interface
        ' this can be used make calls back into HS like hs.SetDeviceStatus, etc.
        ' The callback object is a different interface reserved for plug-ins.
        callback = frm
        hs = frm.GetHSIface
        If hs Is Nothing Then
            MsgBox("Unable to access HS interface", MsgBoxStyle.Critical)
        Else
            hs.WriteLog(IFACE_NAME, "Communicating with HomeSeer...")
            InterfaceVersion = hs.InterfaceVersion
        End If

    End Sub

    ' return status of interface
    ' see Module1.bas for constants
    Public Function InterfaceStatus() As Short
        If Users Is Nothing Then Return ERR_NONE
        If Users.Count < 1 Then Return ERR_NONE
        Dim u As GoogleCalendar
        Dim Problem As Boolean = False
        For x As Integer = 0 To Users.Count - 1
            u = Users.GetByIndex(x)
            If u IsNot Nothing Then
                If u.ComStatus = False Then Problem = True
            End If
        Next
        If Problem Then Return ERR_SEND
        Return ERR_NONE
    End Function

    ' return the access level for this plugin
    ' 1=everyone can access, no protection
    ' 2=level 2 plugin. Level 2 license required to run this plugin
    Public Function AccessLevel() As Short
        AccessLevel = 1
    End Function


    ' This indicates to HS that we are ready for HomeSeer 2.0 in areas such as
    '   SetIOEx instead of SetIO
    Public ReadOnly Property SupportsHS2() As Boolean
        Get
            Return True
        End Get
    End Property

#End Region

#Region "    SetIOEx - Device Command Handler    "

    ' This function is called when an action on a device is to be executed.
    ' If a device's status is set to ON or OFF, or its value changes, this function is called
    ' note that this will be called for ALL I/O devices. You must check and make sure that
    ' the housecode value belongs to you. When your control is initialized, allocate a free
    ' housecode and save it in the registry. Restore it whenever you are initialized
    Public Sub SetIOEx(ByVal dv As Object, ByVal housecode As String, ByVal devicecode As String, ByVal command As Short, _
                       ByVal brightness As Short, ByVal data1 As Short, ByVal data2 As Short, ByVal voice_command As String, ByVal host As String)

        Dim bOurs As Boolean = False

        If Not dv Is Nothing Then
            If dv.interface.trim = IFACE_NAME Then
                bOurs = True
                hs.WriteLog(IFACE_NAME, "SetIOEx called for " & dv.location & " " & dv.location2 & " " & dv.name)
            Else
                Exit Sub    ' Not ours.
            End If
        End If

        If Not bOurs Then
            If gBaseCode.ToUpper <> housecode.Trim.ToUpper Then
                Exit Sub    ' Not ours
            End If
        End If

        hs.WriteLog(IFACE_NAME, "SetIOEx for " & housecode & devicecode & _
                                ", Cmd=" & command.ToString & _
                                ", Brt=" & brightness.ToString & _
                                ", D1/D2=" & data1.ToString & "/" & data2.ToString & _
                                ", Vcmd=" & voice_command & _
                                ", Host=" & host)

    End Sub

#End Region

#Region "    HSEvent - HomeSeer Event Callback Handler    "

    Public Sub HSEvent(ByRef parms() As Object)

        Select Case CInt(parms(0))
            Case EV_TYPE_STATUS_CHANGE
                hs.WriteLog(IFACE_NAME, "Status change " & parms(2) & parms(1) & " Command: " & parms(3).ToString)

            Case EV_TYPE_AUDIO
                ' parms(1) holds TRUE=audio start FALSE=audio stopped
                ' parms(2) holds ID of audio device being controlled
                If parms(1) Then
                    hs.WriteLog(IFACE_NAME, "Audio is started")
                Else
                    hs.WriteLog(IFACE_NAME, "Audio has stopped")
                End If

            Case EV_TYPE_STRING_CHANGE
                hs.WriteLog(IFACE_NAME, "String change event for " & parms(1).ToString & " to string:" & parms(2).ToString)

        End Select
    End Sub

#End Region

#Region "    I/O - Other Plug-In Related Procedures    "

    ' ******************************** I/O Interface *************************
    '
    ' this interface is for any generic device. It will appear in the device list under
    ' the I/O section on the interfaces tab in HS
    ' this can be used for any type of input/output device including RF, temperature, or
    ' relay input output. To represent the device, assign HomeSeer virtual devices to
    ' your individual device points. See the .bas file for sample function that will assign
    ' devices for you
    Public Function InitIO(ByRef port As Integer) As String
        '  In the case of this sample plug-in, we have to call InitIO since that is where
        '       the web page links are done, so we are not going to exit if ANY of the Init
        '       procedures have been called as this next line would normally do.
        '   This is a sample plug-in so it is not a big deal, but normally this is not good
        '       since somebody could enable this plug-in ONLY for IR for example,
        '       and that would mean that InitIO never gets called.
        'If gIREnabled Or gIOEnabled Or gX10Enabled Then Exit Function
        If gIOEnabled Then Return ""

        AlreadyDisplayedError = False

        ' If there are any ASPX web pages that will reference this plug-in other than through scripting,
        '   then uncomment the next line - SetupEnvironment will copy this plug-in's compiled code to the
        '   HomeSeer BIN directory so that it can be referenced by ASPX pages.
        SetupEnvironment("hspi_HSGCal.dll")

        ' This is a great time to use GetResourceToFile to extract any embedded resources that you may need.
        GetResourceToFile("HSGCalPluginDoc.htm", "\HTML\HSGCal\HSGCalPluginDoc.htm")

        ' Using a feature of HomeSeer HS2 versions after 2.2.0.0, let's add our help file to the user's help page:
        Try
            hs.RegisterHelpLink("/HSGCal/HSGCalPluginDoc.htm", "HSGCal Plug-In Help", IFACE_NAME)
        Catch ex As Exception
            ' Do nothing.  This prevents the plug-in from encountering an error on versions of HS that do not support RegisterHelpLink.
        End Try

        Dim s As String = ""
        s = LoadUsers()
        If String.IsNullOrEmpty(s) Then
            ' This is good!
        Else
            Return s
        End If

        ' Get the rest of our configuration settings.
        LoadSettings()

        ' get out base virtual housecode to represent our device
        'InitDevices()

        InitCommon()

        ' Start the calendar update thread
        UpdateThread_Start()

        gIOEnabled = True

        Return ""

    End Function

    Private Sub InitCommon()

        link = "HSGCal"
        linktext = "Calendar Agenda"
        page_title = "Google Calendar Agenda List"
        hs.RegisterLinkEx(Me, IFACE_NAME)


        ConfigPage = New clsConfigPage
        ConfigPage.link = "HSGCalConfig"
        ConfigPage.linktext = "HS GCal Config"
        ConfigPage.page_title = "HomeSeer/Google Calendar Plug-In Configuration"
        hs.RegisterConfigLink(ConfigPage, IFACE_NAME)


        '' Register a callback for notifications of audio changes in HS.
        ''   This will tell us when HS needs to speak and when the speaker client is speaking/done speaking.
        'hs.RegisterEventCB(EV_TYPE_AUDIO, Me)
        '' Register a callback for device string changes.
        'hs.RegisterEventCB(EV_TYPE_STRING_CHANGE, Me)
        '' Register a callback for device status changes as well - there are many other types of callbacks you can use.
        'hs.RegisterEventCB(EV_TYPE_STATUS_CHANGE, Me)

    End Sub

    ' shutdown the I/O interface
    ' called when HS exits
    Public Sub ShutdownIO()
        ' Do something...

        Try
            ConfigPage = Nothing
        Catch ex As Exception
        End Try

        ' Unregister the help resource - normally not needed, but one of our registered links is to a page provided by this 
        '   plug-in, and so since we are shutting down, we should remove our help page links so the user does not run across one
        '   that no longer works.
        Try
            ' First, remove all of them.
            hs.UnRegisterHelpLinks(IFACE_NAME)
            ' Now add back in the help link which still works even after this plug-in is shut down.
            hs.RegisterHelpLink("/HSGCal/HSGCalPluginDoc.htm", "HSGCal Plug-In Help", IFACE_NAME)
        Catch ex As Exception
            ' Do nothing - the calls are Try/Catch bound so that ShutdownIO continues even if there is a problem with these calls.
        End Try

        'then...
        gIOEnabled = False
    End Sub

#End Region

#Region "    Device Config Procedures    "

    ' return TRUE if we support the ability to configure individual devices
    Public ReadOnly Property SupportsConfigDevice2() As Boolean
        Get
            Return False
        End Get
    End Property

#End Region

#Region "    Plug-In Supported Conditions Procedures    "

    Public ReadOnly Property SupportsConditionUI() As Boolean
        ' See also SupportsConditionHTML
        Get
            Return False
        End Get
    End Property

    Public ReadOnly Property SupportsConditionHTML() As Boolean
        '
        ' If SupportsConditionUI is TRUE, then return TRUE here also to use
        '   the newer HTML based ConditionUI procedures, which allow you to 
        '   have more control in the ConditionUI - you can have more than 2 condition
        '   controls and can generate your own HTML to improve the appearance.
        ' Set this to FALSE or do not include it at all to have the plug-in 
        '   use the original ConditionUI related procedures.
        '
        ' If this return TRUE, then the following procedures are required/used:
        '   - ConditionUIHTML
        '   - ConditionUIHTMLProc
        '   - ConditionUIFormatHTML
        '   - ConditionCheckHTML
        ' ... and the following procedures are IGNORED:
        '   -x- ConditionUI
        '   -x- ConditionUIFormat
        '   -x- ConditionCheck
        '
        Get
            Return True
        End Get
    End Property

    ' Condition line
    ' Each condition is a collection of controls that will be displayed in the gui as condition choice
    ' this function is called repeatedly for each condition that the plugin supports
    ' return an empty string if no more conditions
    ' only two controls are allowed for each condition, normally a item/value pair such as:
    '    "zone kitchen" -> "is open"
    ' Only UI_DROP_LIST is supported
    ' the conditions will appear on the condition frame in the HS event properties
    ' The first parameter in the string is very important and cannot be an empty string
    '
    ' Note: This is used only if SupportsConditionHTML returns FALSE or is not present.
    '
    Public Function ConditionUI(ByRef cond_num As Short) As String
        Dim attributes As String = "0"
        Dim s As New StringBuilder
        Dim i As Short

        Select Case cond_num

            Case 1
                ' our zone conditions
                s.Append("Acme Zone Condition")
                s.Append(vbTab & UI_DROP_LIST & vbTab & attributes & vbTab)
                s.Append("Zone: ")
                For i = 0 To 5
                    s.Append(vbTab & "") 'gZoneNames(i))
                Next
                s.Append(vbCr)
                s.Append(vbTab)
                s.Append(vbTab & UI_DROP_LIST & vbTab & attributes & vbTab)
                s.Append("Acme Zone Status is:")
                s.Append(vbTab & "Open" & vbTab & "Closed")
                Return s.ToString

            Case 2
                ' our status conditions
                s.Append("Acme Panel Status")
                s.Append(vbTab & UI_DROP_LIST & vbTab & attributes & vbTab)
                s.Append("Acme Panel Status:")
                s.Append(vbTab & "Is Armed" & vbTab & "Is Disarmed")
                Return s.ToString

            Case Else
                Return ""

        End Select

    End Function

    '
    ' This (and other procedures with HTML in the name) are used only if SupportsConditionHTML is True.
    '   This set of procedures allows more control over the condition UI by allowing HTML that the plug-in 
    '   provides to be passed to HomeSeer for display.
    '
    Public Function ConditionUIHTML(ByVal cond As Integer, ByVal sData As String) As String
        If sData Is Nothing Then sData = ""
        Dim s As New System.Text.StringBuilder

        '
        ' All of the 'HTML related Condition procedures make use of a data string that HomeSeer stores
        '   and subsequently passes to these procedures.  ConditionUIHTMLProc is where the HTML page 
        '   that the user entered data into is returned as item=value pairs, which is what is used as
        '   the data string.
        '
        Dim sTextBox As String = "Can you see me now?"
        Dim sText2 As String = "The OTHER Test"

        ' HomeSeer calls this procedure with a value of Cond starting at 1 and incrementing
        '   each time it is called until a null string ("") is returned.  Each return that is 
        '   not null is a new condition.  When this procedure is called with a -1 value for 
        '   the condition number, then that means the condition is being edited, and the
        '   sData parameter will contain the data string so that the HTML returned can include
        '   the existing values.
        If cond = -1 Then
            ' This is being called as part of an EDIT, so we
            '   need to show the existing data in the returned HTML.
            Dim p() As Pair = Nothing
            Dim paircount As Integer
            ' 
            ' GetFormData converts a string in the form: name=value&name=value&name=value
            '   into pairs for easier processing.
            '
            GetFormData(sData, paircount, p)
            If paircount < 1 Then Return ""
            For i As Integer = 0 To paircount - 1
                If p(i).Name = "MyPlugTextBox" Then
                    sTextBox = p(i).Value
                ElseIf p(i).Name.ToLower = "plug_cond_id" Then
                    ' You can use any method you want for detecting which condition it is
                    '   when there are multiple conditions.  In this example, a hidden field
                    '   is included in the HTML, which is then included in the saved data, 
                    '   that indicates which condition it is.  Since the HTML is then built
                    '   based upon the condition number in the select case below, we will
                    '   set "cond" to the value that corresponds to the appropriate condition.
                    '   If you do not use a select case statement to know which conditions to 
                    '   generate, you can adjust this to use your method to specify the condition.
                    Select Case p(i).Value
                        Case "Test Condition"
                            cond = 1
                        Case "Another Condition"
                            cond = 2
                    End Select
                End If
            Next
        End If


        '
        ' Generate the HTML to be returned to HomeSeer and displayed.  Note that if this condition 
        '   is being edited, then the condition number passed to this procedure was -1, and the data
        '   string must contain some sort of indicator as to which condition is being edited.  If that
        '   was done properly, then the cond parameter is changed to be the real condition number.
        '
        ' IMPORTANT:
        '   You do NOT have unlimited boundaries on what HTML can be returned.  You must understand that this
        '       HTML returned here will be a part of a table and inside a form already.  Do not create any
        '       forms in this HTML unless you can process data entirely within the form using JavaScript as 
        '       the contents of the form will not be returned!  You may create tables, but make sure the tables
        '       are closed properly - do not leave any tables opened or include extra </table> HTML tags or the
        '       output to the user will be corrupted and may prevent your conditions from working.
        '
        Select Case cond
            Case 1
                s.Append("Test HTML Condition" & vbTab)     'Always start with the condition name and vbTab (Chr(2)).
                s.Append("<b>This is a test of the emergency broadcast network.</b><br>")
                s.Append(FormTextBox("Please enter something below:", "MyPlugTextBox", sTextBox, 30))
                s.Append(HTML_NewLine)
                s.Append(FormCheckBox("Option A", "Opt_A", "A", True, False))
                s.Append(HTML_StartFont(COLOR_RED))
                s.Append(FormCheckBox("Option B", "Opt_B", "B", True, False))
                s.Append(HTML_EndFont)
                s.Append(FormCheckBox("Option C", "Opt_A", "A", True, False))
                ' Add my own indicator field so that I know which condition this is.
                s.Append(AddHidden("plug_cond_ID", "Test Condition"))
            Case 2
                s.Append("Another Test Condition" & vbTab)     'Always start with the condition name and vbTab (Chr(2)).
                s.Append("<b>HomeSeer <i>ROCKS</i>!</b><br>")
                s.Append(FormTextBox("", "2nd Condition Text", sText2, 50))
                ' Add my own indicator field so that I know which condition this is.
                s.Append(AddHidden("plug_cond_ID", "Another Condition"))
            Case Else
                Return ""
        End Select

        Return s.ToString

    End Function

    Public Function ConditionUIHTMLProc(ByVal sData As String) As String
        If sData Is Nothing Then Return ""

        If sData.Trim = "" Then Return ""

        '
        ' When the ConditionUIHTML form is submitted to HomeSeer, HomeSeer will process the item=value pairs
        '   and use the ones that it needs to process the page.  The remaining item=value pairs will be
        '   collected and are then passed to this procedure.  In this procedure, you can modify this string
        '   as you see fit, and then the modified string is what HomeSeer will save/store for use in the
        '   other Condition____HTML procedures.
        '
        ' In this example, we are going to take the text that the user entered in the first condition
        '   and append the word " (Modified)" to it if it does not already contain it.
        '
        ' Note that all other item=value pairs are maintained and appended to the string being returned
        '   to HomeSeer.
        '
        Dim p() As Pair = Nothing
        Dim paircount As Integer
        ' 
        ' GetFormData converts a string in the form: name=value&name=value&name=value
        '   into pairs for easier processing.
        '
        GetFormData(sData, paircount, p)
        If paircount < 1 Then Return ""
        Dim sReturn As String = ""
        For i As Integer = 0 To paircount - 1
            If p(i).Name = "MyPlugTextBox" Then
                If Not p(i).Value.ToLower.Contains("modified") Then
                    p(i).Value &= " (Modified)"
                End If
            End If
            If sReturn.Trim.Length > 0 Then sReturn &= "&"
            sReturn &= p(i).Name & "=" & p(i).Value
        Next

        Return sReturn
    End Function


    ' Return a formatted condition for display purposes
    ' String is in the format: index,value
    ' each item is seperated by the character chr(2)
    ' 0 = plugin name (can ignore, used mainly by HS
    ' 1 = condition name, this is field number 1 from ConditionUI() ("Zone Condition" is one example)
    '     use this to identify which condition we are formatting
    ' 2 = first condition value
    ' 3 = second condition value
    '
    ' Note: This is used only if SupportsConditionHTML returns FALSE or is not present.
    '
    Public Function ConditionUIFormat(ByRef cond_str As String) As String
        Dim items() As String

        If cond_str Is Nothing Then Return "Error, bad condition string passed to ConditionUIFormat"
        If InStr(cond_str, Chr(2)) = 0 Then Return "Error, bad condition string passed to ConditionUIFormat"

        items = Split(cond_str, Chr(2))

        Select Case items(1)
            Case "Acme Zone Condition"
                Return "Zone status for " & items(2) & " is " & items(3)

            Case "Acme Panel Status"
                Return "Panel status " & items(2)
        End Select

        Return "Invalid/Unknown condition string passed to ConditionUIFormat."

    End Function

    Public Function ConditionUIFormatHTML(ByVal sInput As String) As String
        If sInput Is Nothing Then Return "The condition is in Error"

        If sInput.Trim = "" Then Return "(Nothing Selected)"
        Dim p() As Pair = Nothing
        Dim paircount As Integer
        ' 
        ' GetFormData converts a string in the form: name=value&name=value&name=value
        '   into pairs for easier processing.
        '
        GetFormData(sInput, paircount, p)
        If paircount < 1 Then Return "(Nothing Selected)"

        For i As Integer = 0 To paircount - 1
            If p(i).Name = "MyPlugTextBox" Then
                Return "You wrote: " & p(i).Value
            ElseIf p(i).Name = "2nd Condition Text" Then
                Return "You wrote: " & p(i).Value
            End If
        Next

        Return "No Sample Condition"

    End Function

    ' Check a given condition and return TRUE if condition is true else false
    ' The cond_str parameter is identical to the string passed to ConditionUIFormat()
    '
    ' Note: This is used only if SupportsConditionHTML returns FALSE or is not present.
    '
    Public Function ConditionCheck(ByRef cond_str As String) As Boolean
        Dim items() As String

        ' assume false condition
        ConditionCheck = False

        If cond_str Is Nothing Then Return False
        If InStr(cond_str, Chr(2)) = 0 Then Return False

        ' get all values
        items = Split(cond_str, Chr(2))

        Select Case items(1)
            Case "Acme Zone Condition"
                ' is this zone status true?
                ' for testing, we just test for the zone being kitchen and the status being Open
                If items(2) = "Kitchen" And items(3) = "Open" Then
                    Return True
                Else
                    Return False
                End If
            Case "Acme Panel Status"
                ' is this panel status true?
                ' for testing, we check for the panel status being armed
                If items(2) = "Is Armed" Then
                    Return True
                Else
                    Return False
                End If
        End Select

    End Function

    Public Function ConditionCheckHTML(ByVal sInput As String) As Boolean
        If sInput Is Nothing Then Return False

        If sInput.Trim = "" Then Return False
        Dim p() As Pair = Nothing
        Dim paircount As Integer
        ' 
        ' GetFormData converts a string in the form: name=value&name=value&name=value
        '   into pairs for easier processing.
        '
        GetFormData(sInput, paircount, p)

        ' 
        ' In this simple sample, if the word TRUE appears in the text
        '   that is entered, then the condition is true.
        '
        For i As Integer = 0 To paircount - 1
            If p(i).Name = "MyPlugTextBox" Then
                If p(i).Value.ToLower.Contains("true") Then Return True
            ElseIf p(i).Name = "2nd Condition Text" Then
                If p(i).Value.ToLower.Contains("true") Then Return True
            End If
        Next
        Return False

    End Function

#End Region

#Region "    Plug-In Supported Triggers Procedures    "

    ' called when HS builds trigger options in event properties
    ' if this plugin supports custom triggers, return TRUE
    Public ReadOnly Property SupportsTriggerUI() As Boolean
        Get
            Return False
        End Get
    End Property

    ' called when HS builds trigger options in event properties
    ' if this plugin supports custom triggers, return the trigger options
    ' the returned string is a list of controls that will be displayed on the trigger tab
    '
    ' Note that when HS saves a trigger from the event properties, it saves the index value for
    ' drop list controls, not the string itself. This allows you to rename trigger entries and
    ' not corrupt configured triggers
    ' trig_str = currently configured trigger when displaying trigger in event properties
    Public Function TriggerUI(ByRef trig_str As String) As String
        Dim attributes As String = "0"
        Dim i As Short
        Dim sbTUI As New StringBuilder

        sbTUI.Append("Panel Trigger")
        sbTUI.Append(vbTab & UI_DROP_LIST & vbTab & attributes & vbTab)
        ' Demonstrating Chr(5) that separates the Name from the Value in the UI drop list.
        '   The trigger string will contain the value after Chr(5).
        sbTUI.Append("Select Zone:")
        For i = 0 To 5
            sbTUI.Append(vbTab & Chr(5) & i.ToString) 'gZoneNames(i) & Chr(5) & i.ToString)
        Next
        sbTUI.Append(vbCr)

        sbTUI.Append(vbTab)
        sbTUI.Append(vbTab & UI_DROP_LIST & vbTab & attributes & vbTab)
        sbTUI.Append("Select Trigger:")
        sbTUI.Append(vbTab & vbTab) 'gZoneActions(0) & vbTab & gZoneActions(1))
        sbTUI.Append(vbCr)

        sbTUI.Append(vbTab)
        sbTUI.Append(vbTab & UI_CHECK_BOX & vbTab & attributes & vbTab)
        sbTUI.Append("Enabled")
        sbTUI.Append(vbTab & "1")
        sbTUI.Append(vbCr)

        sbTUI.Append(vbTab)
        sbTUI.Append(vbTab & UI_LABEL & vbTab & attributes & vbTab)
        sbTUI.Append("All Notes")
        sbTUI.Append(vbTab)
        sbTUI.Append(vbCr)

        sbTUI.Append(vbTab)
        sbTUI.Append(vbTab & UI_TEXT_BOX & vbTab & attributes & vbTab)
        sbTUI.Append("Notes:")
        sbTUI.Append(vbTab)

        ' add another trigger type

        sbTUI.Append(Chr(1) & IFACE_NAME & vbTab)
        sbTUI.Append("Panel Trigger2")
        sbTUI.Append(vbTab & UI_DROP_LIST & vbTab & attributes & vbTab)
        sbTUI.Append("Select Trigger:")
        sbTUI.Append(vbTab & vbTab) 'gZoneActions(0) & vbTab & gZoneActions(1))
        sbTUI.Append(vbCr)

        Return sbTUI.ToString

    End Function


    ' when HS needs to display a trigger to a user, this function is called if the event
    ' is set to trigger using this plugin
    ' the parameter passed here is a string of seperated values, seperated with a chr(2)
    ' each setting is from each control on the trigger tab as it was entered from the
    ' TriggerUI function. The string returned here will be displayed in the event list in HS
    ' and on the web page in the event list
    '
    ' A sample trigger may look like:
    ' ACME Sample Plug-In Chr(2) Panel Trigger Chr(2) kitchen Chr(2) Open Chr(2) notes here
    Public Function TriggerUIFormat(ByRef trig As String) As String
        Dim values() As String
        Dim i As Integer

        If trig Is Nothing Then Return "Bad trigger string passed to TriggerUIFormat."
        If InStr(trig, Chr(2)) = 0 Then Return "Bad trigger string passed to TriggerUIFormat."

        values = Split(trig, Chr(2))
        ' index
        ' 0 = plugin name (not used here, but used in HS to identify owner of trigger)
        ' 1 = trigger name (not needed here)
        ' 2 = value of first control in trigger tab
        ' 3 = value of second, etc.

        If values(1) = "Panel Trigger2" Then
            Return "Motion status: " & values(2)
        Else
            i = Val(values(2).Trim)
            Return "Zone " & i.ToString & " detects " & values(3) & " Notes: " & values(5) 'gZoneNames(i) & " detects " & values(3) & " Notes: " & values(5)
        End If

    End Function


    ' when a user clicks OK after editing a trigger, this function is called so
    ' the data entry can be validated
    ' the trigstr parameter is the actual trigger string that will be saved
    ' return an empty string if no error, else return a string describing
    ' which entry is bad. HS will pop up a dialog with this information
    ' On the web interface, the page will be redisplayed with the error text
    ' displayed in RED across the top
    Public Function ValidateTriggerUI(ByRef trigstr As String) As String
        Dim items() As String
        Dim i As Integer
        Dim s As String

        If trigstr Is Nothing Then Return "Invalid trigger string passed to ValidateTriggerUI"
        If trigstr.Trim = "" Then Return "Invalid trigger string passed to ValidateTriggerUI"
        If InStr(trigstr, Chr(2)) = 0 Then Return "Invalid trigger string passed to ValidateTriggerUI"

        items = Split(trigstr, Chr(2))
        If UBound(items) < 2 Then
            ' The above test will vary depending upon how many controls are in YOUR trigger string.
            Return "Invalid trigger string passed to ValidateTriggerUI"
        End If

        s = Replace(trigstr, Chr(2), "-/-") ' Chr(2) is unprintable - make it show something.
        hs.WriteLog(IFACE_NAME, "DEBUG ValidateTriggerUI trigstr: " & s)

        ' simulate and error if the user selects "Kitchen" as the zone
        i = Val(items(2).Trim)
        If "" = "Kitchen" Then 'gZoneNames(i) = "Kitchen" Then
            Return "Invalid entry, Kitchen"
        Else
            Return ""
        End If

    End Function


    ' when an event happens in this device, call back and pass a formatted string
    ' that HS should match with a trigger. The string must match exactly to the trigger
    ' as passed to the TriggerUIFormat function
    ' HS will attempt to match the entire string taking into account other trigger
    ' parameters such as conditions, dates, days, etc.
    ' If there is a parameter that you do not want HS to match on, set its value to *
    ' For example, in this sample the user may have created a trigger that is formatted
    ' like:
    ' ACME Sample Plug-In Chr(2) Panel Trigger Chr(2) kitchen Chr(2) Motion Chr(2) notes here
    ' The trigger values are index numbers 2(kitchen), 3(Motion), and 4(notes here)
    ' In this case, the last trigger parameter (4) is just a note, so we change that
    ' to an * so HS does not attempt a match on it
    '
    ' THIS IS NOT A PLUG-IN REQUIRED PROCEDURE - IT IS FOR SAMPLE PLUG-IN FUNCTIONALITY ONLY
    '
    Private Sub DeviceTrigger(ByRef zone As String, ByRef action As String)
        Dim trigstr As String = ""
        Dim chk_string() As String
        Dim i As Integer

        ' Do a little bit of verification first...
        If zone Is Nothing Then Exit Sub
        If action Is Nothing Then Exit Sub
        If zone.Trim = "" Then Exit Sub
        If action.Trim = "" Then Exit Sub

        'For i = 0 To UBound(gZoneNames)
        '    ' Set our index (i) to the value corresponding to the zone, since the trigger string
        '    '   contains the index, not the zone name due to our use of Chr(5) in TriggerUI
        '    If zone = gZoneNames(i) Then Exit For
        'Next

        ' build the trigger string based on the event and callback to HS
        trigstr = IFACE_NAME & Chr(2) & "Panel Trigger" & Chr(2) & i.ToString & Chr(2) & action & Chr(2) & "*"

        ' This next step is optional - it uses PreCheckTrigger to find matching triggers.
        ' This is only needed when you have triggers with analog values and you want to see what
        '   values the user entered when they created the trigger.  See the SDK reference for more detail.
        chk_string = callback.PreCheckTrigger(trigstr)
        If chk_string Is Nothing Then Exit Sub
        Try
            If UBound(chk_string) < 0 Then
                Exit Sub
            End If
        Catch ex As Exception
            Exit Sub
        End Try

        If UBound(chk_string) = 0 Then
            If chk_string(0) Is Nothing Then Exit Sub
            If chk_string(0).Trim = "" Then Exit Sub
        End If

        hs.WriteLog(IFACE_NAME, "PreCheckTrigger found " & (UBound(chk_string) + 1).ToString & " triggers that match - triggering events!")

        ' Since we ARE doing a pre-test with PreCheckTrigger and we have gotten to this point, then
        '   we know there is at least one match, so let's call CheckTrigger.
        ' Normally we would go directly to CheckTrigger since we have no triggers with values in a range.
        callback.CheckTrigger(trigstr)

    End Sub

#End Region

#Region "    Plug-In Provided Actions Procedures    "

    ' called when HS builds action options in event properties
    ' if this plugin supports custom actions, return TRUE
    Public ReadOnly Property SupportsActionUI() As Boolean
        Get
            Return False
        End Get
    End Property


    ' list of controls to display in event properties. Same format as the triggerUI function
    ' first entry for drop lists must be "No Action" or an empty string. Index 0 into this control
    ' is always a no action
    ' For text controls, an empty string indicates no action
    ' action_str = currently configured actions when displaying actions in event properties
    Public Function ActionUI(ByRef action_str As String) As String
        Dim Attributes As String = "0"
        Dim sbAUI As New StringBuilder
        Dim i As Short

        sbAUI.Append("Panel Actions")
        ' Zone selection for Bypass action.
        sbAUI.Append(vbTab & UI_DROP_LIST & vbTab & Attributes & vbTab)
        sbAUI.Append("Bypass Zone:")
        sbAUI.Append(vbTab & "No Action")
        For i = 0 To 5
            sbAUI.Append(vbTab & "Zone " & i.ToString) 'gZoneNames(i))
        Next
        sbAUI.Append(vbCr)

        sbAUI.Append(vbTab)
        ' Selection for arm/disarm action.
        sbAUI.Append(vbTab & UI_DROP_LIST & vbTab & CATTR_ALLOW_TEXT_ENTRY & vbTab)
        sbAUI.Append("Panel Actions:")
        sbAUI.Append(vbTab & "No Action")
        sbAUI.Append(vbTab & "Arm Panel" & vbTab & "Disarm Panel")
        sbAUI.Append(vbCr)

        sbAUI.Append(vbTab)
        ' Text box to enter something to be displayed on a security keypad.
        sbAUI.Append(vbTab & UI_TEXT_BOX & vbTab & Attributes & vbTab)
        sbAUI.Append("Display message on keypad:")
        sbAUI.Append(vbTab & "")
        sbAUI.Append(vbCr)

        sbAUI.Append(vbTab)
        ' Checkbox to close Relay 1.
        sbAUI.Append(vbTab & UI_CHECK_BOX & vbTab & Attributes & vbTab)
        sbAUI.Append("Close Relay 1:")
        sbAUI.Append(vbTab & "1")
        sbAUI.Append(vbCr)

        sbAUI.Append(vbTab)
        ' Notes - saved with the event action, but not used when the action is executed.
        'sbAUI.Append(vbTab & UI_LABEL & vbTab & Attributes & vbTab)
        sbAUI.Append(vbTab & UI_TEXT_BOX & vbTab & Attributes & "\M=4\W=20" & vbTab)
        sbAUI.Append("Notes")
        sbAUI.Append(vbTab)
        sbAUI.Append(vbCr)

        sbAUI.Append(vbTab)
        ' A button to press to make the user think they are doing something.  ;-)
        sbAUI.Append(vbTab & UI_BUTTON & vbTab & Attributes & vbTab)
        sbAUI.Append("More ...")
        sbAUI.Append(vbTab)

#If 0 Then  ' Remove this directive to see the 2nd action set.
        sbAUI.Append(Chr(1) & IFACE_NAME & Chr(4))
        sbAUI.Append("Panel Actions 2" & vbTab)

        sbAUI.Append("Panel Actions 2")
        sbAUI.Append(vbTab & UI_DROP_LIST & vbTab & Attributes & vbTab)
        sbAUI.Append("Bypass Zone:")
        sbAUI.Append(vbTab & "No Action")
        For i = 0 To 5
            sbAUI.Append(vbTab & gZoneNames(i))
        Next
        sbAUI.Append(vbCr)

        sbAUI.Append(vbTab)
        sbAUI.Append(vbTab & UI_DROP_LIST & vbTab & CATTR_ALLOW_TEXT_ENTRY & vbTab)
        sbAUI.Append("Panel Actions:")
        sbAUI.Append(vbTab & "No Action" & vbTab & "Arm Panel" & vbTab & "Disarm Panel")
        sbAUI.Append(vbCr)

        sbAUI.Append(vbTab)
        sbAUI.Append(vbTab & UI_TEXT_BOX & vbTab & Attributes & vbTab)
        sbAUI.Append("Display message on keypad:")
        sbAUI.Append(vbTab & "")
        sbAUI.Append(vbCr)

        sbAUI.Append(vbTab)
        sbAUI.Append(vbTab & UI_CHECK_BOX & vbTab & Attributes & vbTab)
        sbAUI.Append("Close Relay 1:")
        sbAUI.Append(vbTab & "1")
        sbAUI.Append(vbCr)

        sbAUI.Append(vbTab)
        sbAUI.Append(vbTab & UI_LABEL & vbTab & Attributes & vbTab)
        sbAUI.Append("Notes")
        sbAUI.Append(vbTab)
        sbAUI.Append(vbCr)

        sbAUI.Append(vbTab)
        sbAUI.Append(vbTab & UI_BUTTON & vbTab & Attributes & vbTab)
        sbAUI.Append("More 2...")
        sbAUI.Append(vbTab)
#End If

        Return sbAUI.ToString

    End Function


    ' When HS needs to display actions for an event, it will call this function with the action
    ' string. A formatted string should be returned describing all actions
    ' This is primarily displayed on web page, so HTML is allowed. The HTML will be stripped if
    ' its displayed in the windows UI
    Public Function ActionUIFormat(ByRef action As String) As String
        Dim sbAUF As New StringBuilder
        Dim values() As String

        If action Is Nothing Then Return "Bad action string passed to ActionUIFormat."
        If InStr(action, Chr(2)) = 0 Then Return "Bad action string passed to ActionUIFormat."

        values = Split(action, Chr(2))
        ' index
        ' 0 = plugin name (not used here, but used in HS to identify owner of trigger)
        ' 1 = not used
        ' 2 = value of first control in trigger tab
        ' 3 = value of second, etc.

        ' bypass zone
        sbAUF.Append("Bypass Zone: " & values(2) & "<br>")

        ' zone action
        sbAUF.Append("Action: " & values(3) & "<br>")

        ' keypad display
        sbAUF.Append("Keypad Display: " & values(4) & "<br>")

        ' relay
        If values(5) = 1 Then
            sbAUF.Append("Set Relay 1")
        End If

        Return sbAUF.ToString

    End Function


    ' when a user clicks OK after editing a trigger, this function is called so
    ' the data entry can be validated
    ' the actionstr parameter is the actual action string that will be saved
    ' return an empty string if no error, else return a string describing
    ' which entry is bad. HS will pop up a dialog with this information
    ' On the web interface, the page will be redisplayed with the error text
    ' displayed in RED across the top
    Public Function ValidateActionUI(ByRef actionstr As String) As String
        Dim items() As String

        If actionstr Is Nothing Then Return ""
        If actionstr.Trim = "" Then Return ""
        If InStr(actionstr, Chr(2)) = 0 Then Return "Invalid action string passed to ValidateActionUI"

        items = Split(actionstr, Chr(2))

        If UBound(items) < 2 Then   ' This is determined by the number of controls in your ActionUI
            Return "Invalid action string passed to ValidateActionUI"
        End If

        ' simulate error if user selects zone kitchen
        If items(2) = "Kitchen" Then
            Return "Invalid zone, Kitchen"
        Else
            Return "" ' actions are OK
        End If

    End Function


    ' When an HS event is triggered, and the action is set to control this plugin, this function
    ' is called. The passed string is a chr(2) seperated list of actions in the order as added with ActionUI
    Public Sub TriggerAction(ByRef actionstr As String)
        Dim actions() As String
        Dim zone As Short
        Dim action As Short
        Dim message As String = ""

        If actionstr Is Nothing Then Exit Sub
        If InStr(actionstr, Chr(2)) = 0 Then Exit Sub

        actions = Split(actionstr, Chr(2))
        ' index 0 = plugin name
        ' index 1 = not used
        ' index 2 = first action
        ' index 3 = second action, etc.
        If UBound(actions) < 3 Then
            hs.WriteLog(IFACE_NAME, "Bad action string passed to TriggerAction.")
            Exit Sub
        End If

        zone = Val(actions(2).Trim)
        action = Val(actions(3).Trim)
        message = actions(4).Trim

        If zone <> 0 Then
            hs.WriteLog(IFACE_NAME, "Bypassing zone " & (zone - 1).ToString) 'gZoneNames(zone - 1))
        End If

        If action <> 0 Then
            Select Case action
                Case 1
                    hs.WriteLog(IFACE_NAME, "Arm the security panel.")
                Case 2
                    hs.WriteLog(IFACE_NAME, "Disarm the security panel.")
            End Select
        End If

        If message <> "" Then
            hs.WriteLog(IFACE_NAME, "Display this message on panel: " & message)
        End If

    End Sub

#End Region

#Region "    Button Related Procedures    "

    ' In HomeSeer version 2.0.2027.0 or later - replaces older ButtonPress functionality
    '   by also passing a device object reference, which is needed to support multiple devices
    '   having the same button names.
    Public ReadOnly Property SupportsExtendedButtons() As Boolean
        Get
            Return True
        End Get
    End Property

    ' if you added buttons to your devices, this function is called when a user
    ' clicks the button. The name of the button is passed.
    ' This is used only if SupportsExtendedButtons is False or not in the plug-in.
    Public Sub ButtonPress(ByRef button_name As String)
        Select Case button_name
            Case "Arm"
                hs.SetDeviceString(gBaseCode & "1", "Armed")
            Case "Disarm"
                hs.SetDeviceString(gBaseCode & "1", "Disamed")
        End Select
    End Sub

    ' if you added buttons to your devices, this function is called when a user
    ' clicks the button. The name of the button is passed.
    ' This is used only if SupportsExtendedButtons is False or not in the plug-in.
    Public Sub ButtonPressEx(ByRef button_name As String, ByRef dv As Object)

        If Not dv Is Nothing Then
            hs.WriteLog(IFACE_NAME, "In ButtonPressEx for " & dv.location & " " & dv.name & ", button pressed is " & button_name)
        Else
            hs.WriteLog(IFACE_NAME, "In ButtonPressEx, bad DeviceClass object provided.")
        End If

        If button_name Is Nothing Then
            hs.WriteLog(IFACE_NAME, "In ButtonPressEx, bad button name provided.")
        End If
        If button_name.Trim = "" Then
            hs.WriteLog(IFACE_NAME, "In ButtonPressEx, bad button name provided.")
        End If

        Select Case button_name
            Case "Arm"
                hs.SetDeviceString(gBaseCode & "1", "Armed")
            Case "Disarm"
                hs.SetDeviceString(gBaseCode & "1", "Disamed")
        End Select

    End Sub

#End Region

#Region "    Web Page Procedures (for HSPI's Web Page)  "

    ' This function is called when a user clicks the link in HS
    ' It returns a web page.
    Private SelectedUser As String = ""

    Private Class evWho
        Public Sub New()
            MyBase.New()
        End Sub
        Public WhoEmail As String = ""
        Public WhoName As String = ""
        Public Entry As Google.GData.Calendar.EventEntry
        Public GC As GoogleCalendar
    End Class
    Public Function GenPage(Optional ByRef lnk As String = "") As String
        Dim p As New StringBuilder
        Dim g As GoogleCalendar
        Dim u As GoogleCalendar.UserPair
        Dim iUser As Integer = 0
        Dim tp() As Pair
        Dim uCount As Integer = 0
        Dim evWho As evWho

        Dim colEvents As New Collections.SortedList
        Dim evList() As Google.GData.Calendar.EventEntry
        Dim sDate As String

        Dim ev As Google.GData.Calendar.EventEntry
        'hs.WriteLog(IFACE_NAME, "GenPage called with: " & lnk)

        Try
            p.Append(HTML_NewLine)

            p.Append(HTML_StartForm())

            If Users IsNot Nothing AndAlso Users.Count > 0 Then
                ReDim tp(Users.Count)
                tp(0).Name = "All Calendars"
                tp(0).Value = "All"
                For i As Integer = 0 To Users.Count - 1
                    g = Users.GetByIndex(i)
                    If g IsNot Nothing Then
                        u = g.User
                        If u IsNot Nothing Then
                            tp(i + 1).Name = u.FriendlyName
                            tp(i + 1).Value = u.UserName
                            uCount += 1
                            If Not String.IsNullOrEmpty(SelectedUser) Then
                                If u.UserName.Trim.ToLower = SelectedUser.Trim.ToLower Then
                                    iUser = i + 1
                                End If
                            End If
                        End If
                    End If
                Next
                If uCount <> tp.Length - 1 Then
                    ReDim Preserve tp(uCount)
                End If
            Else
                ReDim tp(0)
                tp(0).Name = "All Calendars"
                tp(0).Value = "All"
            End If
            p.Append(FormDropDown("Select Calendar", "calendar_user", tp, tp.Length, iUser, True))
            p.Append("&nbsp;&nbsp;&nbsp;&nbsp;")
            p.Append(FormButton("Update", "Refresh Google Data", , , , True))
            p.Append(HTML_NewLine)
            p.Append(HTML_NewLine)


            p.Append(HTML_StartTable(1, 0, 100, ALIGN_LEFT))

TryAgain:
            If Not String.IsNullOrEmpty(SelectedUser) Then
                If SelectedUser.Trim.ToLower = "all" Then
                    For x As Integer = 0 To Users.Count - 1
                        g = Users.GetByIndex(x)
                        If g IsNot Nothing Then
                            evList = g.Get_Exact_Date_Range(Now.Date, Now.AddDays(1).Date)
                            'evList = g.GetAll
                            If evList IsNot Nothing Then
                                If evList.Length > 0 Then
                                    For e As Integer = 0 To evList.Length - 1
                                        sDate = Format(evList(e).Times(0).StartTime, DateSort) & "_" & Rnd.ToString
                                        Try
                                            evWho = New evWho
                                            evWho.WhoEmail = g.User.UserName
                                            evWho.WhoName = g.User.FriendlyName
                                            evWho.Entry = evList(e)
                                            evWho.GC = g
                                            colEvents.Add(sDate & "_" & evList(e).Authors(0).Email.Trim, evWho)
                                        Catch ex As Exception
                                            hs.WriteLog(IFACE_NAME & " Error", "Exception adding event to collection: " & evList(e).Title.Text)
                                        End Try
                                    Next
                                End If
                            End If
                        End If
                    Next
                Else
                    g = Nothing
                    Try
                        g = Users.Item(SelectedUser.Trim.ToLower)
                    Catch ex As Exception
                        g = Nothing
                    End Try
                    If g Is Nothing Then
                        SelectedUser = "All"
                        GoTo TryAgain
                    Else
                        evList = g.Get_Exact_Date_Range(Date.Today, Now.AddDays(1))
                        'evList = g.GetAll
                        If evList IsNot Nothing Then
                            If evList.Length > 0 Then
                                For e As Integer = 0 To evList.Length - 1
                                    sDate = Format(evList(e).Times(0).StartTime, DateSort) & "_" & Rnd.ToString
                                    Try
                                        evWho = New evWho
                                        evWho.WhoEmail = g.User.UserName
                                        evWho.WhoName = g.User.FriendlyName
                                        evWho.Entry = evList(e)
                                        evWho.GC = g
                                        colEvents.Add(sDate, evWho)
                                    Catch ex As Exception
                                        hs.WriteLog(IFACE_NAME & " Error", "Exception adding event to collection: " & evList(e).Title.Text & " Using Key:" & sDate)
                                    End Try
                                Next
                            End If
                        End If
                    End If
                End If
            Else
                SelectedUser = "All"
                GoTo TryAgain
            End If

            Dim recDate As Date = Date.MinValue
            Dim rec As GoogleCalendar.HSGCalEntry = Nothing

            If colEvents IsNot Nothing AndAlso colEvents.Count > 0 Then
                p.Append(HTML_StartRow)
                p.Append(HTML_StartCell("", 20, ALIGN_LEFT, True))
                p.Append("Found " & colEvents.Count.ToString & " events.")
                p.Append(HTML_EndCell)
                p.Append(HTML_EndRow)

                For i As Integer = 0 To colEvents.Count - 1
                    evWho = colEvents.GetByIndex(i)
                    If evWho Is Nothing Then Continue For
                    ev = evWho.Entry
                    If ev Is Nothing Then Continue For
                    rec = New GoogleCalendar.HSGCalEntry(ev, evWho.GC.User)
                    If rec.StartDT.Date <> recDate.Date Then
                        recDate = rec.StartDT.Date
                        p.Append(HTML_StartRow)
                        p.Append(HTML_StartCell("tablecolumn", 20, ALIGN_LEFT, True))
                        p.Append(recDate.ToLongDateString)
                        p.Append(HTML_EndCell)
                        p.Append(HTML_EndRow)
                    End If

                    p.Append(HTML_StartRow)
                    p.Append(HTML_StartCell("", 20, ALIGN_LEFT, True))
                    p.Append(HTML_StartTable(0, 0, 100, ALIGN_LEFT))


                    p.Append(HTML_StartRow)
                    p.Append(HTML_StartCell("", 20, ALIGN_LEFT, True))
                    If SelectedUser.Trim.ToLower = "all" Then
                        p.Append("Appointment for " & evWho.WhoName & HTML_NewLine)
                    End If
                    p.Append(HTML_StartBold & rec.TitleHTML & HTML_EndBold)
                    If rec.Author.Trim.ToLower <> evWho.WhoEmail.Trim.ToLower Then
                        p.Append(HTML_NewLine & "(Created by " & rec.Author.Trim & ")")
                    End If
                    p.Append(HTML_EndCell)
                    p.Append(HTML_EndRow)

                    'hs.writelog(IFACE_NAME & " Debug", "Appointment for " & evWho.WhoName & ": " & rec.Title)

                    p.Append(HTML_StartRow)
                    p.Append(HTML_StartCell("", 20, ALIGN_LEFT, True))
                    p.Append(HTML_StartBold & "Starts: " & HTML_EndBold & rec.StartDT.ToShortTimeString)
                    p.Append("&nbsp;&nbsp;")
                    If rec.EndDT.Date = recDate.Date Then
                        p.Append(HTML_StartBold & "Ends: " & HTML_EndBold & rec.EndDT.ToShortTimeString)
                    Else
                        p.Append(HTML_StartBold & "Ends: " & HTML_EndBold & rec.EndDT.ToLongDateString & " " & rec.EndDT.ToShortTimeString)
                    End If
                    p.Append("&nbsp;&nbsp;")
                    p.Append(HTML_StartBold & "Duration: " & HTML_EndBold & rec.EndDT.Subtract(rec.StartDT).ToString)
                    p.Append(HTML_EndCell)
                    p.Append(HTML_EndRow)

                    Dim RR() As GoogleCalendar.HSGCalEntry.Remind
                    RR = rec.Reminders(evWho.GC)
                    If RR IsNot Nothing Then
                        If RR.Length > 0 Then
                            Dim sRemind As String = ""
                            For r As Integer = 0 To RR.Length - 1
                                Select Case RR(r).RemindMethod
                                    Case GoogleCalendar.HSGCalEntry.enumRemindMethod.Alert
                                        If String.IsNullOrEmpty(sRemind) Then
                                            sRemind = "Alert " & RR(r).RemindWhen.ToString & " prior"
                                        Else
                                            sRemind &= ", Alert " & RR(r).RemindWhen.ToString & " prior"
                                        End If
                                    Case GoogleCalendar.HSGCalEntry.enumRemindMethod.All
                                        If String.IsNullOrEmpty(sRemind) Then
                                            sRemind = "All Methods " & RR(r).RemindWhen.ToString & " prior"
                                        Else
                                            sRemind &= ", All Methods " & RR(r).RemindWhen.ToString & " prior"
                                        End If
                                    Case GoogleCalendar.HSGCalEntry.enumRemindMethod.eMail
                                        If String.IsNullOrEmpty(sRemind) Then
                                            sRemind = "eMail " & RR(r).RemindWhen.ToString & " prior"
                                        Else
                                            sRemind &= ", eMail " & RR(r).RemindWhen.ToString & " prior"
                                        End If
                                    Case GoogleCalendar.HSGCalEntry.enumRemindMethod.None
                                        If String.IsNullOrEmpty(sRemind) Then
                                            sRemind = "(None)"
                                        Else
                                            sRemind &= ", (None)"
                                        End If
                                    Case GoogleCalendar.HSGCalEntry.enumRemindMethod.SMS
                                        If String.IsNullOrEmpty(sRemind) Then
                                            sRemind = "SMS " & RR(r).RemindWhen.ToString & " prior"
                                        Else
                                            sRemind &= ", SMS " & RR(r).RemindWhen.ToString & " prior"
                                        End If
                                    Case Else
                                        If String.IsNullOrEmpty(sRemind) Then
                                            sRemind = "(Unknown) " & RR(r).RemindWhen.ToString & " prior"
                                        Else
                                            sRemind &= ", (Unknown) " & RR(r).RemindWhen.ToString & " prior"
                                        End If
                                End Select
                            Next
                            If Not String.IsNullOrEmpty(sRemind) Then
                                p.Append(HTML_StartRow)
                                p.Append(HTML_StartCell("", 20, ALIGN_LEFT, True))
                                p.Append(HTML_StartBold & "Reminders Set: " & HTML_EndBold & sRemind)
                                p.Append(HTML_EndCell)
                                p.Append(HTML_EndRow)
                            End If
                        End If
                    End If

                    If Not String.IsNullOrEmpty(rec.Description) Then
                        p.Append(HTML_StartRow)
                        p.Append(HTML_StartCell("", 20, ALIGN_LEFT, False))
                        p.Append(rec.Description)
                        p.Append(HTML_EndCell)
                        p.Append(HTML_EndRow)
                    End If
                    If Not String.IsNullOrEmpty(rec.Location) Then
                        p.Append(HTML_StartRow)
                        p.Append(HTML_StartCell("", 20, ALIGN_LEFT, False))
                        p.Append(HTML_StartBold & "Location: " & HTML_EndBold & rec.Location)
                        p.Append(HTML_EndCell)
                        p.Append(HTML_EndRow)
                    End If

                    p.Append(HTML_StartRow)
                    p.Append(HTML_StartCell("", 20, ALIGN_LEFT, True))
                    p.Append(HTML_StartBold & "Status: " & HTML_EndBold & rec.StatusText)
                    p.Append("&nbsp;&nbsp;")
                    p.Append(HTML_StartBold & "Free/Busy: " & HTML_EndBold & rec.FreeBusy.ToString)
                    p.Append("&nbsp;&nbsp;")
                    Dim sVisible As String = ""
                    Select Case rec.Visibility
                        Case GoogleCalendar.HSGCalEntry.enumVisibility.vConfidential
                            sVisible = "Confidential"
                        Case GoogleCalendar.HSGCalEntry.enumVisibility.vDefault
                            sVisible = "Default"
                        Case GoogleCalendar.HSGCalEntry.enumVisibility.vPrivate
                            sVisible = "Private"
                        Case GoogleCalendar.HSGCalEntry.enumVisibility.vPublic
                            sVisible = "Public"
                        Case Else
                            sVisible = "(Unknown)"
                    End Select
                    p.Append(HTML_StartBold & "Visibility: " & HTML_EndBold & sVisible)
                    Dim sCat As String = rec.CategoryList
                    If Not String.IsNullOrEmpty(sCat) Then
                        p.Append("&nbsp;&nbsp;")
                        p.Append(HTML_StartBold & "Categories: " & HTML_EndBold & sCat)
                    End If
                    p.Append(HTML_EndCell)
                    p.Append(HTML_EndRow)

                    'If Not String.IsNullOrEmpty(rec.Recurrence) Then
                    '    p.Append(HTML_StartRow)
                    '    p.Append(HTML_StartCell("", 20, ALIGN_LEFT, False))
                    '    p.Append(HTML_StartBold & "Recurrence: " & HTML_EndBold & rec.Recurrence)
                    '    p.Append(HTML_EndCell)
                    '    p.Append(HTML_EndRow)
                    'End If

                    Dim GAtt() As GoogleCalendar.HSGCalEntry.GCalAttendee
                    GAtt = rec.AttendeesFilt(evWho.WhoEmail)
                    If GAtt IsNot Nothing Then
                        If GAtt.Length > 0 Then
                            Dim sAttend As String = ""
                            For Each att As GoogleCalendar.HSGCalEntry.GCalAttendee In GAtt
                                If att IsNot Nothing Then
                                    If String.IsNullOrEmpty(sAttend) Then
                                        sAttend = att.eMail
                                    Else
                                        sAttend &= ", " & att.eMail
                                    End If
                                    Select Case att.Type
                                        Case GoogleCalendar.HSGCalEntry.enumAttendeeType.AttendanceOptional
                                            sAttend &= "(Optional)"
                                        Case GoogleCalendar.HSGCalEntry.enumAttendeeType.AttendanceRequired
                                            sAttend &= "(Required)"
                                        Case Else
                                    End Select
                                    Select Case att.Status
                                        Case GoogleCalendar.HSGCalEntry.enumAttendeeStatus.Accepted
                                            sAttend &= "=Accepted"
                                        Case GoogleCalendar.HSGCalEntry.enumAttendeeStatus.Declined
                                            If att.Type = GoogleCalendar.HSGCalEntry.enumAttendeeType.AttendanceRequired Then
                                                sAttend &= "=" & HTML_StartBold & "Declined" & HTML_EndBold
                                            Else
                                                sAttend &= "=Declined"
                                            End If
                                        Case GoogleCalendar.HSGCalEntry.enumAttendeeStatus.Invited
                                            sAttend &= "=Invited"
                                        Case GoogleCalendar.HSGCalEntry.enumAttendeeStatus.Tentative
                                            If att.Type = GoogleCalendar.HSGCalEntry.enumAttendeeType.AttendanceRequired Then
                                                sAttend &= "=" & HTML_StartBold & "Tentative" & HTML_EndBold
                                            Else
                                                sAttend &= "=Tentative"
                                            End If
                                        Case Else
                                            sAttend &= "=" & HTML_StartBold & "(Unknown)" & HTML_EndBold
                                    End Select
                                End If
                            Next
                            If Not String.IsNullOrEmpty(sAttend) Then
                                p.Append(HTML_StartRow)
                                p.Append(HTML_StartCell("", 20, ALIGN_LEFT, False))
                                p.Append(HTML_StartBold & "Attendee Status: " & HTML_EndBold & sAttend)
                                p.Append(HTML_EndCell)
                                p.Append(HTML_EndRow)
                            End If
                        End If
                    End If

                    p.Append(HTML_EndTable)
                    p.Append(HTML_EndCell)
                    p.Append(HTML_EndRow)
                Next

            Else
                p.Append(HTML_StartRow)
                p.Append(HTML_StartCell("", 20, ALIGN_LEFT, True))
                p.Append("No events found for today or tomorrow.")
                p.Append(HTML_EndCell)
                p.Append(HTML_EndRow)
            End If

            p.Append(HTML_EndTable)

            ' this is REQUIRED so HS knows what page to display when put is complete
            p.Append(AddHidden("ref_page", Me.link)) ' this must match the registered link
            p.Append(HTML_EndForm)

            'If RAW_PAGE Then
            '    ' optional, return a complete page so our displayed page does not contain the links bar or header
            '    st.Append("HTTP/1.0 200 OK" & vbCrLf)
            '    st.Append("Server: HomeSeer" & vbCrLf)
            '    st.Append("Expires: Sun, 22 Mar 1993 16:18:35 GMT" & vbCrLf)
            '    st.Append("Content-Type: text/html" & vbCrLf)
            '    st.Append("Accept-Ranges: bytes" & vbCrLf)
            '    st.Append("Content-Length: " & Len(data).ToString & vbCrLf & vbCrLf)
            '    ' In raw form, I have to provide the HTML head/body tags, but I am going to let HS provide them
            '    '   anyway with the GetPageHeader procedure.
            '    st.Append(hs.GetPageHeader("My Sample Plug-In Page", "", "", False, False, False, False, False))
            '    st.Append(data.ToString)
            '    ' In raw form, I have to provide the HTML head/body ending tags, but I am going to let HS provide them
            '    '   anyway with the GetPageFooter procedure.
            '    st.Append(hs.GetPageFooter(False))
            'Else
            '    st.Append(data.ToString)
            'End If

        Catch ex As Exception
            hs.WriteLog(IFACE_NAME & " Error", "Exception (" & Erl.ToString & ") generating calendar page: " & ex.Message)
        End Try
        Return p.ToString

    End Function

    ' put requests call here
    Public Function PagePut(ByRef data As String) As String
        Dim i As Integer

        GetFormData(data, Me.lPairs, Me.tPair)


        For i = 0 To lPairs - 1
            Select Case tPair(i).Name.Trim.ToLower
                Case "calendar_user"
                    SelectedUser = tPair(i).Value.Trim
                Case "update"
                    UpdateThread_Trigger = True
                    Return ""
            End Select
        Next

        'done:
        '        If RAW_PAGE Then
        '            ' we supply the page to display
        '            PagePut = GenPage()
        '        Else
        '            ' display page with headers and links
        '            'PagePut = "Ok"
        '            ' display original page
        '            PagePut = ""
        '        End If
        Return ""
    End Function

#End Region

#Region "        Plug-In Script Commands       "

    ' ByRef variables can be changed in a Sub and typically use less storage memory than ByVal.
    ' ByVal variables pass the value (a copy) of the variable rather than a pointer to the variable.
    Public Function GetEvents_DateOnly(ByVal User As String, ByVal Days As Integer) As GoogleCalendar.HSGCalEntry()
        Return _GetEvents(User, Now.Date, Now.AddDays(Days).Date, True)
    End Function
    Public Function GetEventsSelect_DateOnly(ByVal User As String, ByVal StartDate As Date, ByVal Days As Integer) As GoogleCalendar.HSGCalEntry()
        Return _GetEvents(User, StartDate.Date, Now.AddDays(Days).Date, True)
    End Function
    Public Function GetEventsRange_DateOnly(ByVal User As String, ByVal StartDate As Date, ByVal EndDate As Date) As GoogleCalendar.HSGCalEntry()
        Return _GetEvents(User, StartDate.Date, EndDate.Date, True)
    End Function


    Public Function GetEvents_DateTime(ByVal User As String, ByVal Days As Integer) As GoogleCalendar.HSGCalEntry()
        Return _GetEvents(User, Now, Now.AddDays(Days), False)
    End Function
    Public Function GetEventsSelect_DateTime(ByVal User As String, ByVal StartDate As Date, ByVal Days As Integer) As GoogleCalendar.HSGCalEntry()
        Return _GetEvents(User, StartDate, Now.AddDays(Days), False)
    End Function
    Public Function GetEventsRange_DateTime(ByVal User As String, ByVal StartDate As Date, ByVal EndDate As Date) As GoogleCalendar.HSGCalEntry()
        Return _GetEvents(User, StartDate, EndDate, False)
    End Function

    Public Sub UpdateNow()
        UpdateThread_Trigger = True
    End Sub
    Public Function Fetch(ByVal User As String, ByVal DaysForward As Integer, Optional ByVal DaysBackward As Integer = 0) As Boolean
        Return _Fetch(User, DaysForward, DaysBackward)
    End Function

    Private Function _Fetch(ByVal User As String, ByVal DaysForward As Integer, ByVal DaysBackward As Integer) As Boolean
        If String.IsNullOrEmpty(User) Then
            hs.WriteLog(IFACE_NAME & " Error", "Fetch called with an invalid or missing user ID.")
            Return False
        End If

        Try
            Dim g As GoogleCalendar

            g = Nothing
            If User.Contains("@") Then
                Try
                    g = Users.Item(SelectedUser.Trim.ToLower)
                Catch ex As Exception
                    g = Nothing
                End Try
            Else
                Dim Found As Boolean = False
                Try
                    For i As Integer = 0 To Users.Count - 1
                        g = Users.GetByIndex(i)
                        If g IsNot Nothing Then
                            If g.FriendlyName.Trim.ToLower = User.Trim.ToLower Then
                                Found = True
                                Exit For
                            End If
                        End If
                    Next
                    If Not Found Then g = Nothing
                Catch ex As Exception
                    g = Nothing
                End Try
            End If
            If g Is Nothing Then
                If User.Trim.ToLower <> "all" Then
                    hs.WriteLog(IFACE_NAME & " Error", "Fetch could not find the calendar entries for user " & User)
                    Return False
                End If
            End If

            If User.Trim.ToLower = "all" Then
                For i As Integer = 0 To Users.Count - 1
                    g = Users.GetByIndex(i)
                    If g IsNot Nothing Then
                        g.InitRecords(DaysForward, DaysBackward)
                    End If
                Next
            Else
                g.InitRecords(DaysForward, DaysBackward)
            End If

            Return True
        Catch ex As Exception
            hs.WriteLog(IFACE_NAME & " Error", "Fetch exception: " & ex.Message)
            Return False
        End Try

    End Function



    Private Function _GetEvents(ByVal User As String, _
                                ByVal StartDate As Date, _
                                ByVal EndDate As Date, _
                                Optional ByVal DateOnly As Boolean = True) As GoogleCalendar.HSGCalEntry()
        If String.IsNullOrEmpty(User) Then
            hs.WriteLog(IFACE_NAME & " Error", "GetEvents called with an invalid or missing user ID.")
            Return Nothing
        End If

        Dim g As GoogleCalendar

        g = Nothing
        If User.Contains("@") Then
            Try
                g = Users.Item(SelectedUser.Trim.ToLower)
            Catch ex As Exception
                g = Nothing
            End Try
        Else
            Dim Found As Boolean = False
            Try
                For i As Integer = 0 To Users.Count - 1
                    g = Users.GetByIndex(i)
                    If g IsNot Nothing Then
                        If g.FriendlyName.Trim.ToLower = User.Trim.ToLower Then
                            Found = True
                            Exit For
                        End If
                    End If
                Next
                If Not Found Then g = Nothing
            Catch ex As Exception
                g = Nothing
            End Try
        End If

        Dim evList() As Google.GData.Calendar.EventEntry
        Dim sDate As String = ""
        Dim colEvents As New Collections.Generic.List(Of GoogleCalendar.HSGCalEntry)
        Dim gev As GoogleCalendar.HSGCalEntry

        If g Is Nothing Then
            If User.Trim.ToLower <> "all" Then
                hs.WriteLog(IFACE_NAME & " Error", "GetEvents called with a user that could not be found: " & User)
                Return Nothing
            End If
        End If

        If User.Trim.ToLower = "all" Then
            For i As Integer = 0 To Users.Count - 1
                g = Users.GetByIndex(i)
                If g IsNot Nothing Then
                    If DateOnly Then
                        evList = g.Get_Exact_Date_Range(StartDate.Date, EndDate.Date)
                    Else
                        evList = g.Get_DateTime_Range(StartDate, EndDate)
                    End If
                    If evList IsNot Nothing Then
                        If evList.Length > 0 Then
                            For e As Integer = 0 To evList.Length - 1
                                gev = New GoogleCalendar.HSGCalEntry(evList(e), g.User)
                                colEvents.Add(gev)
                            Next
                        End If
                    End If
                End If
            Next
        Else
            If DateOnly Then
                evList = g.Get_Exact_Date_Range(StartDate.Date, EndDate.Date)
            Else
                evList = g.Get_DateTime_Range(StartDate, EndDate)
            End If
            If evList IsNot Nothing Then
                If evList.Length > 0 Then
                    For e As Integer = 0 To evList.Length - 1
                        gev = New GoogleCalendar.HSGCalEntry(evList(e), g.User)
                        colEvents.Add(gev)
                    Next
                End If
            End If
        End If


        If colEvents Is Nothing Then Return Nothing
        If colEvents.Count < 1 Then Return Nothing
        Return colEvents.ToArray

    End Function

#End Region

End Class
