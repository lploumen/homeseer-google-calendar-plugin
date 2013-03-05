Option Explicit On 
Imports System
Imports System.Diagnostics
Imports System.Text
Imports System.Threading
Imports System.IO
Imports VB = Microsoft.VisualBasic

Public Enum UserRight
    User_Guest = 1
    User_Admin = 2
    User_Local = 4
    User_Normal = 8
    User_Guest_Local = 5
    User_Admin_Local = 6
    User_Normal_Local = 12
    User_Invalid = -1
End Enum

Public Class HSUser
    Public UserName As String
    Public Rights As UserRight
End Class

Public Structure Pair
    Public Name As String
    Public Value As String
End Structure


Module HTMLPublic

    Public sWebPage As String

    Public Const USER_GUEST As Integer = 1 ' user can view web pages only, cannot make changes
    Public Const USER_ADMIN As Integer = 2 ' user can make changes
    Public Const USER_LOCAL As Integer = 4 ' this user is used when logging in on a local subnet
    Public Const USER_NORMAL As Integer = 8 ' Not guest, not admin, just NORMAL!
    Public HSUsers As New Collections.SortedList

'   For HyperJump function
    Private Const SW_SHOW As Integer = 5
    Private Const SW_SHOWNORMAL As Integer = 1         'Restores Window if Minimized or Maximized
    Public gHSServerPort As String          'The HomeSeer web server port number.
    Public gIEAppPath As String             'The path to Internet Explorer
    Public gUseIEBrowser As Boolean = False 'Whether to use Internet Explorer or the system default browser (usually IE anyway)
    Public gWebPath As String               'Path to root HS server, usually http://localhost:port
    Public gEXEPath As String               'Path to this executable, usually the HS path.

    ' HTTP constants
    Public Const HTML_StartHead As String = "<head>" & vbCrLf
    Public Const HTML_EndHead As String = "</head>" & vbCrLf
    Public Const HTML_StartPage As String = "<html>" & vbCrLf
    Public Const HTML_EndPage As String = "</html>" & vbCrLf
    Public Const HTML_StartBody As String = "<body>" & vbCrLf
    Public Const HTML_EndBody As String = "</body>" & vbCrLf
    Public Const HTML_EndForm As String = "</form>" & vbCrLf
    Public Const HTML_NewLine As String = "<br>" & vbCrLf
    Public Const HTML_StartPara As String = "<p>"
    Public Const HTML_EndPara As String = "</p>" & vbCrLf
    Public Const HTML_Line As String = "<hr noshade>"
    Public Const HTML_EndTable As String = "</table>" & vbCrLf
    Public Const HTML_StartRow As String = "<tr>"
    Public Const HTML_EndRow As String = "</tr>" & vbCrLf
    Public Const HTML_EndCell As String = "</td>"
    Public Const ALIGN_RIGHT As Integer = 1 ' for cell alignment in tables
    Public Const ALIGN_LEFT As Integer = 2
    Public Const ALIGN_CENTER As Integer = 3
    Public Const ALIGN_TOP As Integer = 4
    Public Const ALIGN_BOTTOM As Integer = 5
    Public Const ALIGN_MIDDLE As Integer = 6
    Public Const HTML_StartBold As String = "<b>"
    Public Const HTML_EndBold As String = "</b>"
    Public Const HTML_StartItalic As String = "<i>"
    Public Const HTML_EndItalic As String = "</i>"
    Public Const HTML_StartHead2 As String = "<h2>"
    Public Const HTML_EndHead2 As String = "</h2>"
    Public Const HTML_StartHead3 As String = "<h3>"
    Public Const HTML_EndHead3 As String = "</h3>"
    Public Const HTML_StartHead4 As String = "<h4>"
    Public Const HTML_EndHead4 As String = "</h4>"
    Public Const HTML_StartTitle As String = "<title>"
    Public Const HTML_EndTitle As String = "</title>" & vbCrLf
    Public Const HTML_EndFont As String = "</font>"

    Public Const COLOR_WHITE As String = "#FFFFFF"
    Public Const COLOR_KEWARE As String = "#0080C0"
    Public Const COLOR_RED As String = "#FF0000"
    Public Const COLOR_BLACK As String = "#000000"
    Public Const COLOR_LT_BLUE As String = "#D9F2FF"
    Public Const COLOR_LT_GRAY As String = "#E1E1E1"
    Public Const COLOR_LT_PINK As String = "#FFB6C1"
    Public Const COLOR_ORANGE As String = "#D58000"
    Public Const COLOR_GREEN As String = "#008000"

    '---- End New Web Stuff ----


Public Class clsLaunchApp
        Public file As String
        Public params As String
        Public directory As String

        <System.MTAThread()> Public Sub LaunchAppThread()
            Dim startInfo As New ProcessStartInfo
            Dim PID As System.Diagnostics.Process
            Dim s As String
            Dim i As Integer
            Dim j As Integer
            On Error Resume Next

            s = file
            If Not directory Is Nothing Then
                If (InStr(s, "\") > 0) Or (InStr(s, "/") > 0) Or (InStr(s, ":") > 0) Then
                Else
                    If Right(Trim(directory), 1) = "\" Then
                        s = directory & file
                    Else
                        If directory = "" Then
                        Else
                            s = directory & "\" & file
                        End If
                    End If
                End If
            End If

            startInfo.FileName = s
            If params Is Nothing Then
            Else
                startInfo.Arguments = params
            End If
            If directory Is Nothing Then
                startInfo.WorkingDirectory = gEXEPath
            Else
                startInfo.WorkingDirectory = directory
            End If
            Err.Clear()
            PID = System.Diagnostics.Process.Start(startInfo)
            If Err.Number <> 0 Then
                hs.WriteLog(IFACE_NAME, "Error Launching application (LaunchAppThread): " & Err.Description)
                hs.WriteLog(IFACE_NAME, "Error (Cont) File is:" & file & " and params is:" & params)
            'Else
            '    PID.PriorityClass = ProcessPriorityClass.BelowNormal
            End If
        End Sub
    End Class

    Public Sub GetServerPortAndPath()
        Dim j As Short
        On Error Resume Next
        gHSServerPort = Trim(hs.GetINISetting("Settings", "svrport", "0"))
        j = Val(gHSServerPort)
        gWebPath = "http://localhost"
        If (j <> 0) And (j <> 80) Then
            gWebPath = gWebPath & ":" & gHSServerPort
        End If
    End Sub


    Public Function AddFuncLink(ByRef ref As String, ByRef label As String, Optional ByVal bSelected As Boolean = False) As String
        Dim st As String = ""
        On Error Resume Next
        If bSelected Then
            st = "<input type=""button"" class=""functionrowbuttonselected"" value=""" & label & """ onClick=""location.href='" & ref & "'""  onmouseover=""this.className='functionrowbutton';"" onmouseout=""this.className='functionrowbuttonselected';"">"
        Else
            st = "<input type=""button"" class=""functionrowbutton"" value=""" & label & """ onClick=""location.href='" & ref & "'"" onmouseover=""this.className='functionrowbuttonselected';"" onmouseout=""this.className='functionrowbutton';"">"
        End If

        AddFuncLink = st
    End Function


    Public Function AddHidden(ByRef Name As String, ByRef Value As String) As String
        Dim st As String = ""
        On Error Resume Next

        st = "<input type=""hidden"" name=""" & Name & """ value=""" & Value & """>"
        Return st
    End Function

    Public Function AddHiddenWithID(ByRef Name As String, ByRef Value As String) As String
        Dim st As String = ""
        On Error Resume Next

        st = "<input type=""hidden"" ID=""" & Name & """ name=""" & Name & """ value=""" & Value & """>"
        Return st
    End Function

    Public Function AddLink(ByRef ref As String, _
                            ByRef label As String, _
                            Optional ByRef image As Object = Nothing, _
                            Optional ByRef w As Short = 0, _
                            Optional ByRef H As Short = 0) As String
        Dim st As String = ""
        On Error Resume Next

        If IsNothing(image) Then
            st = "<a href=""" & ref & """>" & label & "</a>" & vbCrLf
        Else
            If w = 0 Then
                st = "<a href=""" & ref & """>" & label & "<img src=""" & image & """ border=""0""></a>" & vbCrLf
            Else
                st = "<a href=""" & ref & """>" & label & "<img src=""" & image & """ width=""" & w.ToString & """ height=""" & H.ToString & """ border=""0""></a>" & vbCrLf
            End If
        End If
        AddLink = st
    End Function

    Public Function AddNavLink(ByRef ref As String, ByRef label As String, _
                               Optional ByVal bSelected As Boolean = False, _
                               Optional ByVal AltText As String = "", _
                               Optional ByVal target As String = "") As String
        Dim st As String = ""
        On Error Resume Next

        If AltText Is Nothing Then
            AltText = " "
        End If
        If bSelected Then
            If target <> "" Then
                st = "<input type=""button"" class=""linkrowbuttonselected"" value=""" & label & """ alt=""" & AltText & """ onClick=""window.open('" & ref & "','" & target & "')" & """ onmouseover=""this.className='linkrowbutton';"" onmouseout=""this.className='linkrowbuttonselected';"">"
            Else
                st = "<input type=""button"" class=""linkrowbuttonselected"" value=""" & label & """ alt=""" & AltText & """ onClick=""location.href='" & ref & "'"" onmouseover=""this.className='linkrowbutton';"" onmouseout=""this.className='linkrowbuttonselected';"">"
            End If
        Else
            If target <> "" Then
                st = "<input type=""button"" class=""linkrowbutton"" value=""" & label & """ alt=""" & AltText & """ onClick=""window.open('" & ref & "','" & target & "')" & """ onmouseover=""this.className='linkrowbuttonselected';"" onmouseout=""this.className='linkrowbutton';"">"
            Else
                st = "<input type=""button"" class=""linkrowbutton"" value=""" & label & """ alt=""" & AltText & """ onClick=""location.href='" & ref & "'"" onmouseover=""this.className='linkrowbuttonselected';"" onmouseout=""this.className='linkrowbutton';"">"
            End If
        End If

        AddNavLink = st
    End Function

    Public Function AddNavLinkPlugin(ByRef ref As String, ByRef label As String, _
                               Optional ByVal bSelected As Boolean = False, _
                               Optional ByVal AltText As String = "") As String
        Dim st As String = ""
        On Error Resume Next

        If AltText Is Nothing Then
            AltText = " "
        End If

        If bSelected Then
            st = "<input type=""button"" class=""linkrowbuttonselectedplugin"" value=""" & label & """ alt=""" & AltText & """ onClick=""location.href='" & ref & "'"" onmouseover=""this.className='linkrowbuttonplugin';"" onmouseout=""this.className='linkrowbuttonselectedplugin';"">"
        Else
            st = "<input type=""button"" class=""linkrowbuttonplugin"" value=""" & label & """ alt=""" & AltText & """ onClick=""location.href='" & ref & "'"" onmouseover=""this.className='linkrowbuttonselectedplugin';"" onmouseout=""this.className='linkrowbuttonplugin';"">"
        End If

        AddNavLinkPlugin = st
    End Function

    Public Function FormButton(ByRef name As String, _
                                ByRef Value As String, _
                                Optional ByVal AltText As String = "", _
                                Optional ByVal onClick As String = "", _
                                Optional ByVal css_style As String = "formbutton", _
                                Optional ByVal SubmitButton As Boolean = True) As String
        Dim st As String = ""
        Dim btype As String = ""

        On Error Resume Next

        If SubmitButton Then
            btype = "submit"
        Else
            btype = "button"
        End If

        If AltText Is Nothing Then
            If onClick.Trim.Length = 0 Then
                st = st & "<input class=""" & css_style & """ type=""" & btype & """ name=""" & name & """ value=""" & Value & """>" & vbCrLf
            Else
                st = st & "<input class=""" & css_style & """ type=""" & btype & """ name=""" & name & """ value=""" & Value & """ OnClick=""" & onClick & """>" & vbCrLf
            End If
        Else
            If onClick.Trim.Length = 0 Then
                st = st & "<input class=""" & css_style & """ type=""" & btype & """ name=""" & name & """ value=""" & Value & """ alt=""" & AltText & """>" & vbCrLf
            Else
                st = st & "<input class=""" & css_style & """ type=""" & btype & """ name=""" & name & """ value=""" & Value & """ alt=""" & AltText & """ OnClick=""" & onClick & """>" & vbCrLf
            End If
        End If

        FormButton = st
    End Function

    Public Function FormButtonEx(ByRef name As String, ByRef Value As String, Optional ByRef PrevNNL As Boolean = False, _
                              Optional ByRef NNL As Boolean = False, Optional ByVal AltText As String = "") As String
        Dim st As String = ""
        On Error Resume Next

        If PrevNNL Then

            st = HTML_StartCell("", 1, ALIGN_LEFT, True)
        Else
            st = HTML_StartTable(0) & HTML_StartCell("", 1, ALIGN_LEFT, True)
        End If

        If AltText Is Nothing Then
            st = st & "<input class=""formbutton"" type=""submit"" name=""" & name & """ value=""" & Value & """>" & vbCrLf
        Else
            st = st & "<input class=""formbutton"" type=""submit"" name=""" & name & """ value=""" & Value & """ alt=""" & AltText & """>" & vbCrLf
        End If

        If NNL Then
            st = st & HTML_EndCell
        Else
            st = st & HTML_EndCell & HTML_EndTable
        End If

        FormButtonEx = st
    End Function


    Public Function FormCheckBox(ByRef label As String, _
                                  ByRef name As String, _
                                  ByRef Value As String, _
                                  ByRef checked As Boolean, _
                                  Optional ByRef onChange As Boolean = False) As String
        Dim st As String = ""
        Dim chk As String = ""
        On Error Resume Next

        If checked Then
            chk = " checked "
        Else
            chk = ""
        End If
        If onChange Then
            st = "<input class=""formcheckbox"" type=""checkbox""" & chk & " name=""" & name & """ value=""" & Value & """ onClick=""submit();"" > " & label & vbCrLf
        Else
            st = "<input class=""formcheckbox"" type=""checkbox""" & chk & " name=""" & name & """ value=""" & Value & """ > " & label & vbCrLf
        End If
        FormCheckBox = st
    End Function

    Public Function FormCheckBoxEx(ByRef label As String, _
                                   ByRef name As String, _
                                   ByRef Value As String, _
                                   ByRef checked As Boolean, _
                                   Optional ByRef PrevNNL As Boolean = False, _
                                   Optional ByRef NNL As Boolean = False) As String
        Dim st As New StringBuilder
        Dim chk As String = ""
        On Error Resume Next

        If PrevNNL Then
            st.Append(HTML_StartCell("", 1, ALIGN_LEFT, True))
        Else
            st.Append(HTML_EndRow & HTML_StartRow & HTML_StartCell("", 1, ALIGN_LEFT, True))
        End If

        If checked Then
            chk = " checked "
        Else
            chk = ""
        End If

        st.Append("<input class=""FormCheckBoxEx"" type=""checkbox""" & chk & " name=""" & name & """ value=""" & Value & """ > " & label & vbCrLf)

        If NNL Then
            st.Append(HTML_EndCell)
        Else
            st.Append(HTML_EndCell & HTML_EndRow)
        End If

        Return st.ToString
    End Function


    Public Function FormDropDown(ByRef label As String, _
                                 ByRef sName As String, _
                                 ByRef options() As Pair, _
                                 ByRef option_count As Integer, _
                                 ByRef selected As Integer, _
                                 Optional ByRef OnChange As Boolean = False) As String
        Dim st As New StringBuilder
        Dim i As Short
        Dim sel As String = ""
        Dim newline As String = ""
        On Error Resume Next

        If label <> "" Then
            newline = "<br>"
        End If

        st.Append(label & newline)
        If OnChange Then
            st.Append("<select class=""formdropdown"" name=""" & sName & """ size=""1"" onchange=""submit();"">" & vbCrLf)

        Else
            st.Append("<select class=""formdropdown"" name=""" & sName & """ size=""1"">" & vbCrLf)
        End If
        For i = 0 To option_count - 1
            If i = selected Then
                sel = "selected "
            Else
                sel = ""
            End If
            st.Append("<option " & sel & "value=""" & options(i).Value & """>" & options(i).Name & "</option>" & vbCrLf)
        Next
        st.Append("</select>" & vbCrLf)
        FormDropDown = st.ToString
    End Function

    Public Function FormDropDownEx(ByRef label As String, _
                                   ByRef sName As String, _
                                   ByRef options() As Pair, _
                                   ByRef option_count As Integer, _
                                   ByRef selected As Integer, _
                                   Optional ByRef OnChange As Boolean = False, _
                                   Optional ByRef Width As Integer = 1, _
                                   Optional ByRef PrevNNL As Boolean = False, _
                                   Optional ByRef NNL As Boolean = False) As String
        Dim st As New StringBuilder
        Dim i As Short
        Dim sel As String = ""
        Dim newline As String = ""
        On Error Resume Next

        If label <> "" Then
            newline = "<br>"
        End If

        If PrevNNL Then
            st.Append(HTML_StartCell("", 1, ALIGN_LEFT, True))
        Else
            st.Append(HTML_EndRow & HTML_StartRow & HTML_StartCell("", 1, ALIGN_LEFT, True))
        End If

        st.Append(label & newline)

        If OnChange Then
            st.Append("<select class=""FormDropDown"" name=""" & sName & """ size=""" & Width.ToString & """ onchange=""submit();"">" & vbCrLf)
        Else
            st.Append("<select class=""FormDropDown"" name=""" & sName & """ size=""" & Width.ToString & """>" & vbCrLf)
        End If
        For i = 0 To option_count - 1
            If i = selected Then
                sel = "selected "
            Else
                sel = ""
            End If
            st.Append("<option " & sel & "value=""" & options(i).Value & """>" & options(i).Name & "</option>" & vbCrLf)
        Next

        If NNL Then
            st.Append("</select>" & vbCrLf & HTML_EndCell)
        Else
            st.Append("</select>" & vbCrLf & HTML_EndCell & HTML_EndRow)
        End If

        FormDropDownEx = st.ToString

    End Function


    Public Function FormDropDownPair(ByRef label As String, _
                                     ByRef sName As String, _
                                     ByRef options() As Pair, _
                                     ByRef option_count As Integer, _
                                     ByRef selected As Integer, _
                                     Optional ByRef OnChange As Boolean = False) As String
        Dim st As String = ""
        Dim i As Integer
        Dim sel As String = ""
        Dim newline As String = ""
        On Error Resume Next

        If label <> "" Then
            newline = "<br>"
        End If

        st = st & label & newline
        If OnChange Then
            st = st & "<select class=""formdropdown"" name=""" & sName & """ size=""1"" onchange=""submit();"">" & vbCrLf

        Else
            st = st & "<select class=""formdropdown"" name=""" & sName & """ size=""1"">" & vbCrLf
        End If
        For i = 0 To option_count - 1
            If i = selected Then
                sel = "selected "
            Else
                sel = ""
            End If
            st = st & "<option " & sel & "value=""" & options(i).Value & """>" & options(i).Name & "</option>" & vbCrLf
        Next
        st = st & "</select>" & vbCrLf

        Return st

    End Function

    Public Function FormGraphicInputButton(ByVal ButName As String, ByVal ButSource As String, ByVal Value As String, _
                                            Optional ByVal AltText As String = "") As String
        Dim st As String = ""
        On Error Resume Next
        If AltText Is Nothing Then
            st = st & "<input type=""image"" name=""" & ButName & """ src=""" & ButSource & """ value=""" & Value & """ class=""graphicbutton"">" & vbCrLf
        Else
            st = st & "<input type=""image"" name=""" & ButName & """ src=""" & ButSource & """ value=""" & Value & """ class=""graphicbutton"" alt=""" & AltText & """>" & vbCrLf
        End If

        FormGraphicInputButton = st
    End Function


    Public Function FormIRButton(ByRef Name_Renamed As String, ByRef Value As String) As String
        Dim st As String = ""
        On Error Resume Next

        st = st & "<input class=""irbutton"" type=""submit"" name=""" & Name_Renamed & """ value=""" & Value & """>" & vbCrLf
        FormIRButton = st
    End Function

    Public Function FormPasswordTextBox(ByRef label As String, _
                                        ByRef Name As String, _
                                        ByRef Value As String, _
                                        ByRef fieldsize As Integer) As String
        Dim st As String = ""
        Dim newline As String = ""
        On Error Resume Next

        If label <> "" Then
            newline = "<br>"
        End If
        st = label & newline & "<input class=""formtext"" type=""password"" size=""" & fieldsize.ToString & """ name=""" & Name & """ value=""" & Value & """>" & vbCrLf
        FormPasswordTextBox = st
    End Function



    Public Function FormPlugTextArea(ByRef label As String, _
                                     ByRef sName As String, _
                                     ByRef Value As String, _
                                     ByRef fieldsize As Integer, _
                                     ByRef rows As Integer, _
                                     Optional ByRef PrevNNL As Boolean = False, _
                                     Optional ByRef NNL As Boolean = False) As String
        Dim st As String = ""
        Dim newline As String = ""
        On Error Resume Next

        If label <> "" Then
            newline = "<br>"
        End If

        If rows = 0 Then rows = 1

        If PrevNNL Then

            st = HTML_StartCell("", 1, ALIGN_LEFT, True)
        Else
            st = HTML_StartTable(0) & HTML_StartCell("", 1, ALIGN_LEFT, True)
        End If

        st = st & label & newline & "<textarea rows=""" & rows.ToString & """ class=""formtext"" cols=""" & fieldsize.ToString & """ name=""" & sName & """>" & vbCrLf
        st = st & Value & "</textarea>"

        If NNL Then
            st = st & HTML_EndCell
        Else
            st = st & HTML_EndCell & HTML_EndTable
        End If

        FormPlugTextArea = st
    End Function

    Public Function FormRadio(ByRef label As String, _
                               ByRef name As String, _
                               ByRef Value As String, _
                               ByRef checked As Boolean, _
                               Optional ByRef OnChange As Boolean = False) As String
        Dim st As String = ""
        Dim chk As String = ""
        On Error Resume Next

        If checked Then
            chk = " checked "
        Else
            chk = ""
        End If
        If OnChange Then
            st = "<input class=""formradio"" type=""radio""" & chk & " name=""" & name & """ value=""" & Value & """ onClick=""submit();"" > " & label & vbCrLf
        Else
            st = "<input class=""formradio"" type=""radio""" & chk & " name=""" & name & """ value=""" & Value & """> " & label & vbCrLf
        End If

        FormRadio = st
    End Function

    Public Function FormRadioEx(ByRef label As String, _
                                ByRef name As String, _
                                ByRef Value As String, _
                                ByRef checked As Boolean, _
                                Optional ByRef PrevNNL As Boolean = False, _
                                Optional ByRef NNL As Boolean = False) As String
        Dim st As String = ""
        Dim chk As String = ""
        On Error Resume Next

        If PrevNNL Then

            st = HTML_StartCell("", 1, ALIGN_LEFT, True)
        Else
            st = HTML_StartTable(0) & HTML_StartCell("", 1, ALIGN_LEFT, True)
        End If

        If checked Then
            chk = " checked "
        Else
            chk = ""
        End If
        st = st & "<input class=""FormRadioEx"" type=""radio""" & chk & " name=""" & name & """ value=""" & Value & """> " & label & vbCrLf

        If NNL Then
            st = st & HTML_EndCell
        Else
            st = st & HTML_EndCell & HTML_EndTable
        End If

        FormRadioEx = st
    End Function

    Public Function FormTextArea(ByRef label As String, _
                                 ByRef name As String, _
                                 ByRef value As String, _
                                 ByRef rows As Integer, _
                                 ByRef cols As Integer) As String
        Dim st As String = ""

        st = label & "<textarea rows=""" & rows.ToString & """ cols=""" & cols.ToString & """ name=""" & name & """>" & value & "</textarea>" & vbCrLf

        FormTextArea = st
    End Function

    Public Function FormTextBox(ByRef label As String, _
                                ByRef Name As String, _
                                ByRef Value As String, _
                                ByRef fieldsize As Integer) As String
        Dim st As String = ""
        Dim newline As String = ""
        On Error Resume Next

        If label <> "" Then
            newline = "<br>"
        End If
        st = label & newline & "<input class=""formtext"" type=""text"" size=""" & fieldsize.ToString & """ name=""" & Name & """ value=""" & Value & """>" & vbCrLf
        FormTextBox = st
    End Function

    Public Function FormTextBoxEx(ByRef label As String, _
                                  ByRef Name As String, _
                                  ByRef Value As String, _
                                  ByRef fieldsize As Integer, _
                                  Optional ByRef PrevNNL As Boolean = False, _
                                  Optional ByRef NNL As Boolean = False) As String
        Dim st As String = ""
        Dim newline As String = ""
        On Error Resume Next

        If label <> "" Then
            newline = "<br>"
        End If

        If PrevNNL Then

            st = HTML_StartCell("", 1, ALIGN_LEFT, True)
        Else
            st = HTML_EndRow & HTML_StartRow & HTML_StartCell("", 1, ALIGN_LEFT, True)
        End If

        st = st & label & newline & "<input class=""formtext"" type=""text"" size=""" & fieldsize.ToString & """ name=""" & Name & """ value=""" & Value & """>" & vbCrLf

        If NNL Then
            st = st & HTML_EndCell
        Else
            st = st & HTML_EndCell & HTML_EndRow
        End If

        FormTextBoxEx = st
    End Function


    Public Function HTML_Header(Optional ByRef level As Short = 1, Optional ByRef align As Short = 0) As String
        Dim stalign As String = ""

        If IsNothing(align) Then
            stalign = ""
        Else
            If align = ALIGN_RIGHT Then
                stalign = " align=""right"""
            ElseIf align = ALIGN_LEFT Then
                stalign = " align=""left"""
            ElseIf align = ALIGN_CENTER Then
                stalign = " align=""center"""
            End If
        End If

        HTML_Header = "<h" & level.ToString & stalign & ">"

    End Function

    Public Function HTML_HeaderEnd(Optional ByRef level As Short = 1) As String

        HTML_HeaderEnd = "</h" & level.ToString & ">"

    End Function

    Public Function HTML_StartCell(ByRef Class_name As String, _
                                   ByRef colspan As Integer, _
                                   Optional ByRef align As Integer = 0, _
                                   Optional ByRef nowrap As Boolean = False) As String
        Dim st As String = ""
        Dim stalign As String = ""
        Dim wrap As String = ""
        On Error Resume Next

        If IsNothing(nowrap) Then
            wrap = ""
        Else
            If nowrap Then
                wrap = " nowrap"
            Else
                wrap = ""
            End If
        End If

        If IsNothing(align) Then
            stalign = ""
        Else
            If align = ALIGN_RIGHT Then
                stalign = " align=""right"""
            ElseIf align = ALIGN_LEFT Then
                stalign = " align=""left"""
            ElseIf align = ALIGN_CENTER Then
                stalign = " align=""center"""
            ElseIf align = ALIGN_TOP Then
                stalign = " align=""top"""
            ElseIf align = ALIGN_BOTTOM Then
                stalign = " align=""bottom"""
            ElseIf align = ALIGN_MIDDLE Then
                stalign = " align=""middle"""
            End If
        End If
        If Class_name = "" Then
            st = "<td" & wrap & stalign & " colspan=""" & colspan.ToString & """>"
        Else
            st = "<td" & wrap & stalign & " colspan=""" & colspan.ToString & """ class=""" & Class_name & """>"
        End If
        HTML_StartCell = st
    End Function

    Public Function HTML_StartCellW(ByRef Class_name As String, _
                                    ByRef colspan As Integer, _
                                    Optional ByRef cwidth As Integer = -1, _
                                    Optional ByRef align As Integer = 0, _
                                    Optional ByRef nowrap As Boolean = False) As String
        Dim st As String = ""
        Dim stalign As String = ""
        Dim wrap As String = ""
        Dim cw As String = ""
        On Error Resume Next

        If (IsNothing(cwidth) Or (cwidth = -1)) Then
            cw = ""
        Else
            cw = " width=""" & cwidth.ToString & Chr(34)
        End If

        If IsNothing(nowrap) Then
            wrap = ""
        Else
            If nowrap Then
                wrap = " nowrap"
            Else
                wrap = ""
            End If
        End If

        If IsNothing(align) Then
            stalign = ""
        Else
            If align = ALIGN_RIGHT Then
                stalign = " align=""right"""
            ElseIf align = ALIGN_LEFT Then
                stalign = " align=""left"""
            ElseIf align = ALIGN_CENTER Then
                stalign = " align=""center"""
            ElseIf align = ALIGN_TOP Then
                stalign = " align=""top"""
            ElseIf align = ALIGN_BOTTOM Then
                stalign = " align=""bottom"""
            End If
        End If
        If Class_name = "" Then
            st = "<td" & cw & wrap & stalign & " colspan=""" & colspan.ToString & """>"
        Else
            st = "<td" & cw & wrap & stalign & " colspan=""" & colspan.ToString & """ class=""" & Class_name & """>"
        End If
        HTML_StartCellW = st
    End Function

    Public Function HTML_StartFont(ByRef color As String) As String
        Dim st As String = ""
        On Error Resume Next

        st = "<font color=""" & color & """>"
        HTML_StartFont = st
    End Function

    Public Function HTML_StartForm(Optional ByVal formname As String = "") As String
        On Error Resume Next
        If formname.Trim.Length = 0 Then
            Return "<form method=""post"">"
        Else
            Return "<form method=""post"" name=""" & formname & """>"
        End If

    End Function

    Public Function HTML_StartTable(ByRef border As Short, _
                                    Optional ByRef spacing As Integer = 0, _
                                    Optional ByRef width As Integer = 0, _
                                    Optional ByVal align As Integer = 0) As String
        Dim st As String = ""
        Dim w As String
        On Error Resume Next

        If IsNothing(width) Then
            w = ""
        Else
            w = "width=""" & width.ToString & "%"""
        End If

        If align = ALIGN_CENTER Then
            st = "<div align=""center"">"
        ElseIf align = ALIGN_LEFT Then
            st = "<div align=""left"">"
        ElseIf align = ALIGN_RIGHT Then
            st = "<div align=""right"">"
        End If
        If IsNothing(spacing) Then
            st = st & "<table border=""" & border.ToString & """ cellpadding=""0"" cellspacing=""0"" " & w & ">" & vbCrLf
        Else
            st = st & "<table border=""" & border.ToString & """ cellpadding=""0"" cellspacing=""" & spacing.ToString & """" & " " & w & "> " & vbCrLf
        End If
        If align <> 0 Then
            st = st & "</div>"
        End If
        HTML_StartTable = st
    End Function

    Public Function HTML_WrapSpan(ByVal id As String, ByVal wraptext As String, _
                                   Optional ByVal bdisplay As Boolean = False, _
                                   Optional ByVal WithTable As Boolean = False) As String
        Dim st As New StringBuilder
        On Error Resume Next

        If WithTable Then
            st.Append("<td>" & vbCrLf)
        End If
        If bdisplay Then
            st.Append("<span style=""display='';"" id=""" & id & """>" & vbCrLf)
        Else
            st.Append("<span style=""display='none';"" id=""" & id & """>" & vbCrLf)
        End If
        st.Append(wraptext)
        st.Append("</span>" & vbCrLf)
        If WithTable Then
            st.Append("</td>")
        End If

        Return st.ToString

    End Function

    Public Function PageError(ByRef Error_Renamed As String, ByRef clr As String, ByRef info As String) As String
        Dim st As String = ""
        On Error Resume Next
        st = "<html><head></head><body>"
        st = st & "<p><table width=""100%"" border=""2"" cellspacing=""0"" cellpadding=""30"">"
        st = st & "<tr><td bgcolor=""" & clr & """ align=""center"">"
        st = st & "<b><font size=""7"">" & Error_Renamed & "</font></b>"
        st = st & "</td></tr></table></p>"
        st = st & "<center>"
        st = st & info
        st = st & "</body></html>"
        Return st
    End Function


    Public Function UrlDecode(ByRef sEncoded As String) As String
        '========================================================
        ' Accept url-encoded string
        ' Return decoded string
        '========================================================

        Dim x As Integer ' sEncoded position pointer
        Dim pos As Integer ' position of InStr target
        On Error Resume Next

        UrlDecode = sEncoded
        If sEncoded = "" Then Exit Function

        ' convert "+" to space
        x = 1
        Do
            pos = InStr(x, sEncoded, "+")
            If pos = 0 Then Exit Do
            Mid(sEncoded, pos, 1) = " "
            x = pos + 1
        Loop

        x = 1

        ' convert "%xx" to character
        Do
            pos = InStr(x, sEncoded, "%")
            If pos = 0 Then Exit Do

            Mid(sEncoded, pos, 1) = Chr(CInt("&H" & (Mid(sEncoded, pos + 1, 2))))
            sEncoded = VB.Left(sEncoded, pos) & Mid(sEncoded, pos + 3)
            x = pos + 1
        Loop
        UrlDecode = sEncoded
    End Function

    Public Function URLEncode(ByRef strData As String) As String

        Dim i As Short
        Dim strTemp As String = ""
        Dim strChar As String = ""
        Dim strOut As String = ""
        Dim intAsc As Short

        strTemp = Trim(strData)
        For i = 1 To Len(strTemp)
            strChar = Mid(strTemp, i, 1)
            intAsc = Asc(strChar)
            If (intAsc >= 48 And intAsc <= 57) Or (intAsc >= 97 And intAsc <= 122) Or (intAsc >= 65 And intAsc <= 90) Then
                strOut = strOut & strChar
            Else
                strOut = strOut & "%" & Hex(intAsc)
            End If
        Next i

        URLEncode = strOut

    End Function

    Public Sub GetFormData(ByRef sFormData As String, ByRef lPairs As Integer, ByRef tPair() As Pair)
        '================================================
        ' Get the CGI data from STDIN or from QueryString
        ' Store name/value pairs
        '================================================
        Dim pointer As Integer ' sFormData position pointer
        Dim N As Integer ' name/value pair counter
        Dim delim1 As Integer ' position of "="
        Dim delim2 As Integer ' position of "&"
        Dim iQPtr As Short ' position of "?"

        'On Error Resume Next

        '=========================================
        ' Parse and decode form data
        '=========================================
        ' Data is received from browser as "name=value&name=value&...name=value"
        ' Names and values are URL-encoded
        '
        ' Store name/value pairs in array tPair(), and decode the values
        '
        pointer = 1
        lPairs = 0

        If InStr(sFormData, "?") > 0 Then
            ' Get the name of the web page with the trailing backslash
            sWebPage = Left(sFormData, InStr(sFormData, "?") - 1)
            ' Now get rid of the trailing backslash on the front of the web page name
            sWebPage = Right(sWebPage, Len(sWebPage) - 1)
        Else
            sWebPage = ""
        End If

        Do
            delim1 = InStr(pointer, sFormData, "=")
            If delim1 = 0 Then Exit Do
            pointer = delim1 + 1

            lPairs = lPairs + 1
        Loop

        ReDim tPair(lPairs)

        pointer = 1

        For N = 0 To (lPairs - 1)
            delim1 = InStr(pointer, sFormData, "=") ' find next equal sign
            If delim1 = 0 Then Exit For ' parse complete
            If pointer = 1 Then
                iQPtr = InStr(sFormData, "?")
                If iQPtr < delim1 Then
                    tPair(N).Name = UrlDecode(Mid(sFormData, iQPtr + 1, delim1 - (iQPtr + 1)))
                Else
                    tPair(N).Name = UrlDecode(Mid(sFormData, pointer, delim1 - pointer))
                End If
            Else
                tPair(N).Name = UrlDecode(Mid(sFormData, pointer, delim1 - pointer))
            End If
            delim2 = InStr(delim1, sFormData, "&")
            ' if no trailing ampersand, we are at the end of data
            If delim2 = 0 Then delim2 = Len(sFormData) + 1
            ' value is between the "=" and the "&"
            tPair(N).Value = UrlDecode(Mid(sFormData, delim1 + 1, delim2 - delim1 - 1))
            pointer = delim2 + 1
        Next N

    End Sub


    Public Sub HyperJump(ByVal url As String)
        Dim w As String

        If url.Trim.ToLower.StartsWith("http") Then
            w = url
        Else
            If url.Trim.StartsWith("\") Or url.Trim.StartsWith("/") Then
                w = gWebPath & url
            Else
                w = gWebPath & "\" & url
            End If
        End If
        Try
            If gIEAppPath = "" Then
                gIEAppPath = IEPath()
            End If
            If gUseIEBrowser Then
                If gIEAppPath <> "" Then
                    LaunchApp(gIEAppPath, "-new " & w)
                Else
                    LaunchApp(w)
                End If
            Else
                LaunchApp(w)
            End If
        Catch ex As Exception
            hs.WriteLog(IFACE_NAME, "Error in HyperJump, error is " & ex.Message & " Address:" & w)
        End Try

    End Sub

    Public Function IEPath() As String
        Dim regkey As Microsoft.Win32.RegistryKey
        Dim regstr As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\IEXPLORE.EXE"

        Try
            regkey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(regstr, False)
            IEPath = regkey.GetValue("")
        Catch ex As Exception
            IEPath = ""
        End Try

    End Function

    Public Function LaunchApp(ByVal file As String, Optional ByVal params As String = "", Optional ByVal directory As String = "") As Integer
        Try
            Dim la As New clsLaunchApp
            Dim la_th As New Thread(AddressOf la.LaunchAppThread)
            la.file = file
            la.params = params
            la.directory = directory
            la_th.Name = "Launch_" & file
            la_th.Start()
        Catch ex As Exception
            hs.WriteLog(IFACE_NAME, "Error Launching application: " & file & "->" & ex.Message)
        End Try
    End Function

    Public Sub GetEXEPath()
        gEXEPath = System.IO.Path.GetDirectoryName(System.AppDomain.CurrentDomain.BaseDirectory)
    End Sub

    Public Function GetUserRight(ByVal sUser As String) As UserRight
        Dim sAllUsers As String = ""
        Dim UserPairs() As String = Nothing
        Dim sTemp As String = ""
        Dim User() As String
        Dim iRight As Short
        Dim i As Short, x As Short
        Dim HSU As New HSUser
        Dim DE As DictionaryEntry

        Try
            sAllUsers = hs.GetUsers
            UserPairs = Split(sAllUsers, ",")
            i = UBound(UserPairs)
        Catch ex As Exception
            hs.WriteLog(IFACE_NAME, "Error (1) processing HomeSeer users: " & ex.Message)
            HSU.Rights = UserRight.User_Invalid
            Return HSU.Rights
        End Try

        Try
            HSUsers.Clear()
            For x = 0 To i
                sTemp = UserPairs(x)
                User = Split(sTemp, "|")
                sTemp = User(0).Trim.ToUpper
                iRight = Val(User(1).Trim)
                HSU = New HSUser
                HSU.UserName = sTemp
                HSU.Rights = iRight
                Try
                    HSUsers.Add(sTemp, HSU)
                Catch ex As Exception
                End Try
            Next
        Catch ex As Exception
            hs.WriteLog(IFACE_NAME, "Error (2) processing HomeSeer users: " & ex.Message)
        End Try

        For Each DE In HSUsers
            HSU = DE.Value
            If HSU.UserName = sUser.Trim.ToUpper Then
                Return HSU.Rights
            End If
        Next
        Return UserRight.User_Invalid

    End Function


End Module


