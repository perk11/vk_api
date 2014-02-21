Imports System.Windows.Forms
Imports vk.notification
Imports System.Xml

Public Class notification
    Shared notifications As New List(Of String)
    Public Shared Sub notify(s As String)
        notifications.Add(s)
    End Sub
    Public Shared Function GetList() As List(Of String)
        Dim result As List(Of String) = notifications
        notifications = New List(Of String)
        Return result
    End Function
End Class

Public Class api
    Public Enum permissions
        groups = 262144
    End Enum
    Public Shared vkauthurl As String
    Const maxqueries As Byte = 3
    Shared token_locker As New Object
    Public Shared token As String = Nothing, tokenexpires As Date, lastqueriesdate(maxqueries - 1) As Date
    Private Shared waitfornextquerylock As New Object
    Public Shared client_id As String
    '    Public Shared userid As ULong
    Public Shared debug As Boolean
    Public Shared user_notified As Boolean = False

    Public Shared showform As Boolean = False
    Public Shared apprights As Long
    Public Shared actual_current_user As New users.user
    Public Shared ReadOnly Property current_user As users.user
        Get
            If actual_current_user.uid = 0 Then
                actual_current_user.uid = users._get()(0).uid
            End If
            Return actual_current_user
        End Get
    End Property
    Public Shared ReadOnly Property IsLoggedIn As Boolean
        Get
            If token = Nothing Then
                Return False
            Else
                Dim compareresult As Integer
#If Not Debug Then
                Try
#End If
                    compareresult = tokenexpires.CompareTo(Now)
#If Not Debug Then
                Catch
                    compareresult = -1
                End Try
#End If
                If compareresult < 0 Then
                    Return False
                End If
            End If
            Return True
        End Get
    End Property
    Public Shared Sub setoptions(Optional clientid As String = Nothing, Optional rights As UInteger = Nothing, Optional scope As List(Of String) = Nothing, Optional updatetoken As Boolean = False, Optional debug_notifications As Boolean = False)
        If Not IsNothing(clientid) Then client_id = clientid
        If IsNothing(client_id) Then Throw New Exception("Client ID not set")
        vkauthurl = "https://oauth.vk.com/authorize?client_id=" & client_id & "&redirect_uri=https%3A%2F%2Foauth.vk.com%2Fblank.html&display=page&response_type=token"

        If Not IsNothing(scope) Then
            vkauthurl = vkauthurl & "&scope="

            Dim current_right As UInt16 = 0
            Dim total_rights As UInt16 = scope.Count
            For Each rght In scope
                vkauthurl = vkauthurl & System.Web.HttpUtility.UrlEncode(rght)
                current_right = current_right + 1
                If current_right < total_rights Then
                    vkauthurl = vkauthurl & ","
                End If
            Next
        ElseIf Not rights = Nothing Then
            vkauthurl = vkauthurl & "&scope=" & rights
            apprights = rights
        Else
            Throw New Exception("App rights not set in setoptions call")

        End If
        If updatetoken Then token = Nothing
        debug = debug_notifications
        If Not My.Settings.tokenrights = apprights Then
            token = Nothing
        End If
        GetToken()
        'users.GetUserName(current_user)
    End Sub


    Shared Sub New()
        '   MsgBox("init")

        For i = 0 To maxqueries - 1
            lastqueriesdate(i) = Now
        Next
        token = My.Settings.vktoken
        tokenexpires = My.Settings.tokenexpires

    End Sub

    Public Shared Sub GetToken()
        If Not IsLoggedIn Then
            Dim t = New Threading.Thread(AddressOf GetNewToken)
            t.SetApartmentState(Threading.ApartmentState.STA)

            t.Start()
            t.Join()
            If Not IsLoggedIn Then
                Dim ex As New Exception("Failed to authorize at vk.com")
                ex.Data("reason") = "Auth window closed"
                Throw ex
            End If
            My.Settings.tokenrights = apprights
        End If
    End Sub
    Public Shared Sub GetNewToken()
        If IsNothing(vkauthurl) Then Throw New Exception("Call setoptions before trying to get token")
        Dim auth As New auth_window
        Dim currurl As String = Nothing
        If Not IsNothing(auth.WebBrowser1.Url) Then
            Try
                currurl = auth.WebBrowser1.Url.ToString
            Catch ex As Exception
                'MsgBox("failed to get current url from browser")
            End Try
        End If
        If Not currurl = vkauthurl Then
            auth.WebBrowser1.Navigate(vkauthurl)
            showform = True
            While showform
                System.Threading.Thread.Sleep(5)
                Application.DoEvents()
            End While
        End If
    End Sub

    Public Shared Sub savetoken()
        My.Settings.vktoken = token
        My.Settings.tokenexpires = tokenexpires
        '    My.Settings.userid = userid
        My.Settings.Save()
    End Sub
    Private Shared Function api_get_url(method As String, parameters As Dictionary(Of String, String)) As String
        GetToken()
        Dim request_text As String
        request_text = method & ".xml"
        If Not IsNothing(parameters) Then
            request_text = request_text & "?"
            For Each s In parameters
                request_text = request_text & "&" & System.Web.HttpUtility.UrlEncode(s.Key) & "=" & System.Web.HttpUtility.UrlEncode(s.Value)
            Next
        End If
        Return "https://api.vk.com/method/" & request_text & "&access_token=" & token
    End Function
    Private Shared Sub wait_for_next_request_time()
        Dim waittime As Int32
        SyncLock waitfornextquerylock
            ' Try
            waittime = 1050 - Now.Subtract(lastqueriesdate(0)).TotalMilliseconds
            If waittime > 0 Then
                System.Threading.Thread.Sleep(waittime)
            End If
            '  Catch
            'End Try
        End SyncLock
        For i = 0 To maxqueries - 2
            lastqueriesdate(i) = lastqueriesdate(i + 1)
        Next i
        lastqueriesdate(maxqueries - 1) = Now
    End Sub
    Public Shared Function api_request(method As String, Optional parameters As Dictionary(Of String, String) = Nothing) As XmlDocument
        Dim url As String = api_get_url(method, parameters)
        Dim doc As XmlDocument = New XmlDocument()
        wait_for_next_request_time()
        Try
            doc.Load(New XmlTextReader(url))
        Catch ex As Exception
            Throw New Exception("Failed to read from URL " & url & ":" & ex.Message, ex)
            Return Nothing
        End Try
        Dim node As XmlNode = doc.SelectSingleNode("error")
        If node IsNot Nothing Then
            Dim code As Long, msg As String
            Try
                code = Long.Parse(node.SelectSingleNode("error_code").InnerText)
                msg = node.SelectSingleNode("error_msg").InnerText
            Catch ex As Exception
                Throw New Exception("Improperly formatted error response: " + doc.InnerXml, ex)
            End Try
            error_handler(code, msg, method)
            doc = Nothing
        End If
        Return doc
    End Function
    Public Shared Function api_request_no_parse(method As String, Optional parameters As Dictionary(Of String, String) = Nothing) As XmlTextReader
        Dim url As String = api_get_url(method, parameters)
        wait_for_next_request_time()
        Try
            Return New XmlTextReader(url)
        Catch ex As Exception
            Throw New Exception("Failed to read from URL " & url & ":" & ex.Message, ex)
            Return Nothing
        End Try
    End Function
    Public Class status
        Inherits api
        Public Shared Sub setaudio(audio_id As ULong, owner_id As Long)
            Dim response_code As Long = -1
            Dim parameters As New Dictionary(Of String, String)
            parameters.Add("audio", owner_id.ToString & "_" & audio_id.ToString)
            Dim s As Xml.XmlTextReader = api_request_no_parse("status.set", parameters)
            While s.Read
                If s.NodeType = Xml.XmlNodeType.Element Then
                    If s.Name = "response" Then
                        response_code = s.ReadElementContentAsLong
                    ElseIf s.Name = "error" Then
                        error_handler_TextReader(s, "status.setaudio")
                    End If
                End If
            End While
            If Not response_code = 1 Then
                Throw New Exception("Got response code " & response_code & "from vk")
            End If
        End Sub
    End Class
    Public Class audio
        Public Shared Sub setBroadcast(Optional audio_id As ULong = 0, Optional owner_id As Long = 0, Optional target_ids As List(Of Long) = Nothing) 'Если аудиозапись не задана, броадкаст убирается
            Dim parameters As New Dictionary(Of String, String)
            Dim successful_ids As New List(Of Long)
            If Not (audio_id = 0 Or owner_id = 0) Then
                parameters.Add("audio", owner_id.ToString & "_" & audio_id.ToString)
            End If
            If Not IsNothing(target_ids) Then
                parameters.Add("target_ids", String.Join(",", target_ids))
            End If
            Dim s As Xml.XmlTextReader = api_request_no_parse("audio.setBroadcast", parameters)
            While s.Read

                If s.NodeType = Xml.XmlNodeType.Element Then
                    If s.Name = "response" Then
                        Dim response As XmlReader = s.ReadSubtree
                        While response.Read
                            If response.NodeType = XmlNodeType.Element Then
                                If response.Name = "item" Then
                                    successful_ids.Add(response.ReadElementContentAsLong)
                                End If
                            End If
                        End While
                    ElseIf s.Name = "error" Then
                        error_handler_TextReader(s, "audio.setBroadcast")
                    End If
                End If
            End While
            If IsNothing(target_ids) Then
                If Not successful_ids.Contains(current_user.uid) Then
                    error_handler(-1, "Не удалось установить трансляцию на страницу пользователя")
                End If
            Else
                For Each id In target_ids
                    If Not successful_ids.Contains(id) Then
                        error_handler(-1, "Не удалось установить трансляцию для ID " & id.ToString)
                    End If
                Next
            End If

        End Sub
        Public Shared Function search(query As String, Optional count As UInt32 = 0, Optional offset As UInt32 = 0) As SearchResult
            If debug Then
                notify("Поиск " & query & " в ВК")
            End If
            Dim parameters As New Dictionary(Of String, String)
            parameters.Add("q", query)
            If count > 0 Then
                parameters.Add("count", count)
            End If
            If offset > 0 Then
                parameters.Add("offset", offset)
            End If
            parameters.Add("auto_complete", "0")
            parameters.Add("sort", "2") 'По популярности
            parameters.Add("v", "5.0")
            Dim s As Xml.XmlTextReader = api_request_no_parse("audio.search", parameters)
            Dim rslt As New SearchResult
            's.ReadToFollowing("count")
            '      rslt.count = s.ReadElementContentAsLong
            '     Debug.Print(s.ReadContentAsLong)
            While s.Read()
                If s.NodeType = Xml.XmlNodeType.Element Then
                    Select Case s.Name
                        Case "audio"
                            Dim record As Xml.XmlReader = s.ReadSubtree, aid As ULong, owner_id As Long, duration As ULong, artist As String = Nothing, title As String = Nothing
                            While record.Read
                                Select Case record.Name
                                    Case "id"
                                        aid = s.ReadElementContentAsLong
                                    Case "owner_id"
                                        owner_id = s.ReadElementContentAsLong
                                    Case "duration"
                                        duration = s.ReadElementContentAsLong
                                    Case "artist"
                                        artist = s.ReadElementContentAsString
                                    Case "title"
                                        title = s.ReadElementContentAsString
                                End Select
                            End While
                            rslt.tracks.Add(New track(aid, owner_id, duration, artist, title))
                        Case "count"
                            rslt.count = s.ReadElementContentAsLong
                        Case "error"
                            error_handler_TextReader(s, "audio.search")
                    End Select
                End If
            End While
            Return rslt
        End Function
        Public Class SearchResult
            Public count As ULong, tracks As New List(Of track)
        End Class
        Public Shared Function getDictionary(Optional userid As ULong = Nothing) As Dictionary(Of track.trackname, track)
            If debug Then
                notify("Получаем список аудиозаписей пользователя в ВК..")
            End If
            Dim parameters As New Dictionary(Of String, String)
            If Not (userid = 0) Then
                parameters.Add("uid", userid)
            End If
            parameters.Add("v", "5.0")
            Dim result As New Dictionary(Of track.trackname, track)
            Dim s As Xml.XmlReader = api_request_no_parse("audio.get", parameters)
            While s.Read
                If s.NodeType = Xml.XmlNodeType.Element Then
                    If s.Name = "audio" Then
                        Dim audio As Xml.XmlReader = s.ReadSubtree, id As ULong, owner_id As Long, duration As ULong, artist As String = Nothing, title As String = Nothing
                        While audio.Read
                            If audio.NodeType = Xml.XmlNodeType.Element Then
                                Select Case audio.Name
                                    Case "id"
                                        id = audio.ReadElementContentAsLong
                                    Case "owner_id"
                                        owner_id = audio.ReadElementContentAsLong
                                    Case "artist"
                                        artist = audio.ReadElementContentAsString
                                    Case "title"
                                        title = audio.ReadElementContentAsString
                                    Case "duration"
                                        duration = audio.ReadElementContentAsLong
                                End Select

                            End If
                        End While
                        Try
                            result.Add(New track.trackname(artist, title), New track(id, owner_id, duration, artist, title))
                        Catch
#If DEBUG Then
                            notify("Track " & artist & "-" & title & " is already in  the list of tracks")
#End If
                        End Try
                    ElseIf s.Name = "error" Then
                        error_handler_TextReader(s)
                    End If
                End If
            End While
            Return result
        End Function
        <Serializable()> Public Class track
            <Serializable()> Public Class trackname
                Public artist As String, title As String
                Public Sub New(track_artist As String, track_title As String)
                    artist = track_artist
                    title = track_title
                End Sub
                Shadows ReadOnly Property ToString
                    Get
                        If IsNothing(Me) Then
                            Return Nothing
                        ElseIf (Not IsNothing(artist)) And (Not IsNothing(title)) Then
                            Return artist & " - " & title
                        Else
                            Return Nothing
                        End If
                    End Get
                End Property
                Public Overrides Function Equals(obj As Object) As Boolean
                    If IsNothing(Me) Then
                        If IsNothing(obj) Then
                            Return True
                        Else
                            Return False
                        End If
                    Else
                        If IsNothing(obj) Then
                            Return False
                        Else
                            If Me.title.Equals(obj.title) Then
                                If Trim(Me.artist).Equals(Trim(obj.artist)) Then
                                    If Trim(Me.title).Equals(Trim(obj.title)) Then
                                        Return True
                                    Else
                                        Return False
                                    End If
                                Else
                                    Return False
                                End If

                            Else
                                Return False
                            End If

                        End If

                    End If
                    '    Return MyBase.Equals(obj)
                End Function
                Public Shared Operator =(ByVal param1 As trackname, ByVal param2 As trackname)
                    If IsNothing(param1) Then
                        If IsNothing(param2) Then
                            Return True
                        Else
                            Return False
                        End If

                    Else
                        If IsNothing(param2) Then
                            Return False
                        Else
                            If UCase(Trim(param1.title)) = UCase(Trim(param2.title)) Then
                                If UCase(Trim(param1.artist)) = UCase(Trim(param2.artist)) Then
                                    Return True
                                Else
                                    Return False
                                End If
                            Else
                                Return False
                            End If

                        End If

                    End If
                End Operator
                Public Shared Operator <>(ByVal param1 As trackname, ByVal param2 As trackname)
                    If param1 = param2 Then Return False Else Return True
                End Operator
            End Class
            Public id As ULong, owner_id As Long, duration As ULong, artist As String, title As String
            Public Sub New(audio_id As ULong, user_owner_id As Long, track_duration As ULong)
                id = audio_id
                owner_id = user_owner_id
                duration = track_duration
            End Sub
            Public Sub New(audio_id As ULong, user_owner_id As Long, track_duration As ULong, track_artist As String, track_title As String)
                id = audio_id
                owner_id = user_owner_id
                duration = track_duration
                artist = track_artist
                title = track_title
            End Sub
            Public Property name As trackname
                Get
                    Return New trackname(artist, title)
                End Get
                Set(value As trackname)
                    artist = value.artist
                    title = value.title
                End Set
            End Property

        End Class
    End Class
    Public Class groups
        Public Class group
            Public gid As Long, name As String, screenname As String, is_closed As Boolean, is_admin As Boolean, is_member As Boolean, type As grouptype, photo As String, photo_medium As String, photo_big As String, description As String, members_count As Long, _
                activity As String, status As String, links As List(Of link), verified As Boolean, site As String, city As String, country As String
            Public Class counters
                Public photos As Long, albums As Long, topics As Long
            End Class
            Public Class link
                Public url As String, name As String, desc As String, photo_50 As String, photo_100 As String
            End Class
            Public Enum grouptype
                page
                group
                [event]
                unknown
            End Enum
            Public Sub New(group_id As Long)
                gid = group_id
            End Sub
            Public Sub New()
            End Sub
        End Class
        Public Shared Function _get(user_id As ULong, Optional extended As Boolean = False, Optional filter As List(Of getfilter) = Nothing, Optional fields As List(Of groupfields) = Nothing, Optional offset As ULong = 0, Optional count As Integer = 0) As List(Of group)
            Dim result As New List(Of group), parameters As New Dictionary(Of String, String)
            With parameters
                .Add("user_id", user_id)
                .Add("extended", IIf(extended, "1", "0"))
                If Not IsNothing(filter) Then .Add("filter", String.Join(",", filter))
                If Not IsNothing(fields) Then .Add("filter", String.Join(",", fields))
                If offset > 0 Then .Add("offset", offset)
                If count > 0 Then .Add("count", count)
            End With
            Dim s As XmlTextReader = api_request_no_parse("groups.get", parameters)

            While s.Read

                If s.NodeType = Xml.XmlNodeType.Element Then
                    If s.Name = "group" Then
                        result.Add(GroupParser(s.ReadSubtree))

                    ElseIf s.Name = "error" Then
                        error_handler_TextReader(s, "groups.get")
                    End If
                End If
            End While
            Return result
        End Function
        Public Shared Shadows Function getById(group_id As String, Optional fields As List(Of String) = Nothing) As group
            Return getById(New List(Of String)({group_id}), fields)(0)
        End Function
        Public Shared Shadows Function getById(group_ids As List(Of String), Optional fields As List(Of String) = Nothing) As List(Of group)
            Dim result As New List(Of group), parameters As New Dictionary(Of String, String)
            With parameters
                If group_ids.Count = 1 Then .Add("group_id", group_ids(0)) Else .Add("group_ids", String.Join(",", group_ids))
                .Add("V", "5.11")
                If Not IsNothing(fields) Then .Add("filter", String.Join(",", fields))
            End With
            Dim s As XmlTextReader = api_request_no_parse("groups.getById", parameters)

            While s.Read

                If s.NodeType = Xml.XmlNodeType.Element Then
                    If s.Name = "group" Then
                        result.Add(GroupParser(s.ReadSubtree))
                    ElseIf s.Name = "error" Then
                        error_handler_TextReader(s, "groups.get")
                    End If
                End If
            End While
            Return result
        End Function
        Private Shared Function GroupParser(group As XmlReader) As group
            'Поля из fields НЕ парсятся, надо дописать при необходимости
            Dim result As New group
            While group.Read
                If group.NodeType = XmlNodeType.Element Then
                    With result
                        Select Case group.Name
                            Case "gid", "id"
                                .gid = group.ReadElementContentAsLong
                            Case "name"
                                .name = group.ReadElementContentAsString
                            Case "screenname", "screen_name"
                                .screenname = group.ReadElementContentAsString
                            Case "is_closed"
                                .is_closed = IIf(group.ReadElementContentAsInt = 1, True, False)
                            Case "is_admin"
                                .is_admin = IIf(group.ReadElementContentAsInt = 1, True, False)
                            Case "is_member"
                                .is_member = IIf(group.ReadElementContentAsInt = 1, True, False)
                            Case "type"
                                Try
                                    .type = [Enum].Parse(GetType(groups.group.grouptype), group.ReadElementContentAsString)
                                Catch
                                    .type = groups.group.grouptype.unknown
                                End Try
                            Case "photo", "photo_50"
                                .photo = group.ReadElementContentAsString
                            Case "photo_medium", "photo_100"
                                .photo_medium = group.ReadElementContentAsString
                            Case "photo_big", "photo_200"
                                .photo_big = group.ReadElementContentAsString
                        End Select
                    End With
                End If
            End While
            Return result
        End Function

        Public Enum getfilter
            admin
            editor
            moder
            groups
            publics
            events
        End Enum
        Public Enum groupfields
            city
            country
            place
            description
            wiki_page
            members_count
            counters
            start_date
            end_date
            can_post
            can_see_all_posts
            activity
            status
            contacts
            links
            fixed_post
            verified
            site
        End Enum
    End Class
    Public Class messages
        <Serializable()> Public Class message
            Public mid As UInt64, uid As Long, sent_date As Date, read_state As Boolean, out As Boolean, title As String, body As String, fwd_messages As New List(Of Message), chat_id As UInt64, chat_active As List(Of users.user), users_count As UInt16, admin_id As UInt64, deleted As Boolean
            'read_state - True - message is read
        End Class
        <Serializable()> Public Class chat
            Public chat_id As UInt64
            Public users As List(Of users.user)
            Public admin As UInt64
        End Class
        Private Shared Function MessageParser(message As XmlReader) As messages.message
            Dim curr_message As New messages.message
            While message.Read
                If message.NodeType = Xml.XmlNodeType.Element Then
                    With curr_message
                        Select Case message.Name
                            Case "mid"
                                .mid = message.ReadElementContentAsLong
                            Case "uid"
                                .uid = message.ReadElementContentAsLong
                            Case "date"
                                .sent_date = New Date(1970, 1, 1, 0, 0, 0, 0).AddSeconds(message.ReadElementContentAsLong).ToLocalTime ' Unix time to Date conversion
                            Case "read_state"
                                .read_state = IIf(message.ReadElementContentAsInt = 1, True, False)
                            Case "out"
                                .out = IIf(message.ReadElementContentAsInt = 1, True, False)
                            Case "title"
                                .title = message.ReadElementContentAsString
                            Case "body"
                                .body = message.ReadElementContentAsString
                            Case "fwd_messages"
                                'Throw New Exception("fwd_messages not implemented")
                                Dim fwd_messages As XmlReader = message.ReadSubtree
                                While fwd_messages.Read
                                    If fwd_messages.NodeType = Xml.XmlNodeType.Element Then
                                        If fwd_messages.Name = "message" Then
                                            .fwd_messages.Add(MessageParser(fwd_messages))
                                        End If
                                    End If
                                End While
                                'Debug.Print(message.ReadOuterXml)
                            Case "chat_id"
                                .chat_id = message.ReadElementContentAsLong
                            Case "chat_active"
                                Dim chat_active As String = message.ReadElementContentAsString, chat_users As New List(Of users.user)
                                For Each uid In chat_active.Split(",")
                                    chat_users.Add(New users.user(uid))
                                Next
                                .chat_active = chat_users
                            Case "users_count"
                                .users_count = message.ReadElementContentAsLong
                            Case "admin_id"
                                .admin_id = message.ReadElementContentAsLong
                        End Select
                    End With

                End If
            End While
            Return curr_message

        End Function
        Public Shared Function getDialogs(Optional uid As UInt64 = 0, Optional chat_id As UInt64 = 0, Optional offset As UInt64 = 0, Optional count As UInt64 = 0) As List(Of Message)
            Dim result As New List(Of message), parameters As New Dictionary(Of String, String)
            With parameters
                If uid > 0 Then .Add("uid", uid)
                If chat_id > 0 Then .Add("chat_id", chat_id)
                If offset > 0 Then .Add("offset", offset)
                If count > 0 Then
                    If count > 200 Then
                        result.AddRange(getDialogs(uid, chat_id, offset + 200, count - 200))
                        .Add("count", 200)
                    Else
                        .Add("count", count)
                    End If
                End If

            End With

            Dim s As XmlTextReader = api_request_no_parse("messages.getDialogs", parameters)

            While s.Read
                If s.NodeType = Xml.XmlNodeType.Element Then
                    If s.Name = "response" Then
                        Dim response As XmlReader = s.ReadSubtree
                        While response.Read
                            If response.NodeType = Xml.XmlNodeType.Element Then
                                If response.Name = "message" Then
                                    result.Add(MessageParser(response.ReadSubtree))
                                End If
                            End If
                        End While
                    End If
                End If


            End While
            Return result
        End Function
        Public Shared Function getDialogsCount(Optional uid As UInt64 = 0) As UInt32
            Dim parameters As New Dictionary(Of String, String)
            If uid > 0 Then parameters.Add("uid", uid)
            parameters.Add("count", 1)
            Dim doc As XmlDocument = api_request("messages.getDialogs", parameters)
            Dim count As Xml.XmlNode = doc.SelectNodes("/response/count").Item(0)
            Return Long.Parse(count.InnerText)
        End Function
        Public Shared Function getHistory(Optional uid As Long = 0, Optional chat_id As Long = 0, Optional offset As UInt64 = 0, Optional count As UInt64 = 0, Optional start_mid As UInt64 = 0, Optional rev As Boolean = False, Optional progressbar As ProgressBar = Nothing, Optional recursive As Boolean = False) As List(Of messages.message)
            Dim parameters As New Dictionary(Of String, String), result As New List(Of messages.message)
            If Not recursive And count > 200 Then
                If Not IsNothing(progressbar) Then
                    progressbar.Maximum = count
                End If
            End If
            With parameters
                If uid > 0 Then
                    .Add("uid", uid)
                ElseIf chat_id > 0 Then
                    .Add("chat_id", chat_id)
                Else
                    Throw New Exception("Both uid and chat_id are zero, getHistory fails")
                End If
                If offset > 0 Then .Add("offset", offset)
                If start_mid > 0 Then .Add("start_mid", start_mid)

                If count > 0 Then
                    If count > 200 Then
                        result.AddRange(getHistory(uid, chat_id, offset + 200, count - 200, start_mid, rev, progressbar, True))
                        If Not IsNothing(progressbar) Then
                            progressbar.PerformStep()
                            progressbar.Update()
                        End If
                        .Add("count", 200)
                    Else
                        .Add("count", count)
                    End If
                End If
                .Add("rev", IIf(rev, 1, 0))
            End With
            Dim s As XmlTextReader = api_request_no_parse("messages.getHistory", parameters)

            While s.Read
                If s.NodeType = Xml.XmlNodeType.Element Then
                    If s.Name = "response" Then
                        Dim response As XmlReader = s.ReadSubtree
                        While response.Read
                            If response.NodeType = Xml.XmlNodeType.Element Then
                                If response.Name = "message" Then
                                    result.Add(MessageParser(response.ReadSubtree))
                                End If
                            End If
                        End While
                    End If
                End If


            End While
            Return result
        End Function
    End Class
    Protected Overloads Shared Function fields_to_string(fields As List(Of String)) As String
        Dim fields_string As String = Nothing
        If Not IsNothing(fields) Then
            For Each field In fields
                If fields_string = Nothing Then
                    fields_string = field
                Else
                    fields_string = fields_string & "," & field
                End If
            Next
        End If
        Return fields_string
    End Function
    Protected Overloads Shared Function fields_to_string(fields As List(Of Long)) As String
        Dim fields_string As String = Nothing
        If Not IsNothing(fields) Then
            For Each field In fields
                If fields_string = Nothing Then
                    fields_string = field
                Else
                    fields_string = fields_string & "," & field.ToString
                End If
            Next
        End If
        Return fields_string
    End Function
    Public Class friends
        Public Shared Function _get(Optional uid As ULong = Nothing, Optional fields As List(Of String) = Nothing) As List(Of users.user)
            Dim result As New List(Of users.user)
            Dim parameters As New Dictionary(Of String, String)
            If Not uid = Nothing Then
                parameters.Add("uid", uid)
            End If
            Dim fields_string As String = fields_to_string(fields)
            If Not IsNothing(fields_string) Then
                parameters.Add("fields", fields_string)
            End If
            Dim s As XmlTextReader = api_request_no_parse("friends.get", parameters)

            If Not IsNothing(fields) Then
                If fields.Count = 0 Then fields = Nothing
            End If
            If IsNothing(fields) Then
                While s.Read
                    If s.NodeType = XmlNodeType.Element Then
                        If s.Name = "uid" Then
                            result.Add(New users.user(s.ReadElementContentAsLong))
                        End If
                    End If
                End While
            Else
                While s.Read
                    If s.NodeType = XmlNodeType.Element Then
                        If s.Name = "user" Then
                            result.Add(users.parse_user_reader(s.ReadSubtree))
                        ElseIf s.Name = "error" Then
                            error_handler_TextReader(s, "friends.get")

                        End If
                    End If
                End While
            End If
            Return result
        End Function
    End Class
    Public Class likes
        Public Enum like_object
            post
            comment
            photo
            audio
            video
            note
            sitepage
        End Enum
        Public Enum like_filter
            likes
            copies
        End Enum
        Public Shared Function getList(type As like_object, Optional owner_id As Int64 = 0, Optional item_id As Int64 = 0, Optional page_url As String = Nothing, Optional filter As like_filter = like_filter.likes, Optional friends_only As Boolean = False, Optional offset As ULong = 0, Optional count As ULong = 0) As List(Of users.user)
            Throw New Exception("Not implemented")
        End Function
    End Class
    Public Class account
        Public Shared Function getAppPermissions() As Long 'НЕ РАБОТАЕТ!
            Dim response As XmlTextReader = api_request_no_parse("account.getAppPermissions"), result As Long = 0
            While response.Read
                If response.NodeType = XmlNodeType.Element Then
                    If response.Name = "uid" Then
                        result = response.ReadContentAsLong
                    ElseIf response.Name = "error" Then
                        error_handler_TextReader(response)
                    End If
                End If
            End While
            Return result
        End Function
    End Class
    Public Class users
        Private Shared Sub GetUserName(ByRef us As users.user)
            Dim info As List(Of users.user) = users._get(New List(Of String)({us.uid}), New List(Of String)({"first_name", "last_name"}))
            If info.Count > 0 Then
                us.first_name = info(0).first_name
                us.last_name = info(0).last_name
            End If
        End Sub
        Public Shared Function parse_user_reader(s As XmlReader) As user
            Dim curr_user As New user
            While s.Read
                While s.Read
                    If s.NodeType = XmlNodeType.Element Then
                        Select Case s.Name
                            Case "first_name"
                                curr_user.first_name = s.ReadElementContentAsString
                            Case "last_name"
                                curr_user.last_name = s.ReadElementContentAsString
                            Case "uid"
                                curr_user.uid = s.ReadElementContentAsLong
                            Case "error"
                                error_handler_TextReader(s)
                        End Select
                    End If
                End While

            End While
            Return curr_user
        End Function
        <Serializable()> Public Class user
            Private p_usertracks As Dictionary(Of audio.track.trackname, audio.track) = Nothing, p_firstname As String = Nothing, p_lastname As String = Nothing

            Public uid As ULong, first_name As String, last_name As String, user_friends As New Dictionary(Of String, List(Of user)), user_messages As New List(Of messages.message)
            Public Shadows ReadOnly Property ToString As String
                Get
                    If IsNothing(first_name) And IsNothing(last_name) Then
                        GetUserName(Me)
                    End If
                    If IsNothing(first_name) And IsNothing(last_name) Then
                        Return Nothing
                    Else
                        Return first_name & " " & last_name
                    End If

                End Get
            End Property
            Public ReadOnly Property tracks As Dictionary(Of audio.track.trackname, audio.track)
                Get
                    If IsNothing(p_usertracks) Then
                        p_usertracks = audio.getDictionary
                    End If
                    Return p_usertracks
                End Get
            End Property
            Public Sub New(Optional user_id As ULong = 0)
                If user_id > 0 Then uid = user_id
            End Sub
            Public ReadOnly Property friends(Optional fields As List(Of String) = Nothing) As List(Of user)
                Get
                    Dim fieldsstring
                    If Not IsNothing(fields) Then
                        Dim bldr As New Text.StringBuilder
                        For Each fld In fields
                            bldr.Append(fld)
                        Next
                        fieldsstring = bldr.ToString
                    Else
                        fieldsstring = "Nofields"
                    End If
                    If Not user_friends.ContainsKey(fieldsstring) Then
                        user_friends(fieldsstring) = api.friends._get(uid, fields)
                    End If
                    Return user_friends(fieldsstring)
                End Get
            End Property
        End Class

        Shared Function _get(Optional uids As List(Of String) = Nothing, Optional fields As List(Of String) = Nothing, Optional name_case As String = Nothing) As List(Of user)
            Dim parameters As New Dictionary(Of String, String)
            Dim uids_string As String = Nothing, fields_string As String = Nothing

            fields_string = fields_to_string(fields)
            With parameters
                If Not IsNothing(uids) Then .Add("uids", String.Join(",", uids))
                If Not IsNothing(fields_string) Then
                    .Add("fields", fields_string)
                End If
                If Not IsNothing(name_case) Then
                    .Add("name_case", name_case)
                End If
            End With
            Dim s As XmlDocument = api_request("users.get", parameters), result As New List(Of user)
            If s IsNot Nothing Then
                For Each node As XmlNode In s.SelectNodes("/response/user")
                    Dim curr_user As New user
                    curr_user.uid = Long.Parse(node.SelectSingleNode("uid").InnerText)
                    curr_user.first_name = node.SelectSingleNode("first_name").InnerText
                    curr_user.last_name = node.SelectSingleNode("last_name").InnerText
                    result.Add(curr_user)
                Next
            End If


            Return result
        End Function

    End Class
    Public Class photos
        Public Shared Function getAllCount(Optional owner_id As Int64 = 0) As UInt32
            Dim parameters As New Dictionary(Of String, String)
            If owner_id > 0 Then parameters.Add("owner_id", owner_id)
            parameters.Add("count", 1)
            Dim doc As XmlDocument = api_request("photos.getAll", parameters)
            Dim count As Xml.XmlNode = doc.SelectNodes("/response/count").Item(0)
            Return Long.Parse(count.InnerText)
        End Function
        Public Shared Function getAll(Optional owner_id As Int64 = 0, Optional no_service_albums As Boolean = False, Optional offset As UInt32 = 0, Optional count As UInt32 = 0, Optional extended As Boolean = False) As List(Of photo)
            Dim result As New List(Of photo)
            Dim parameters As New Dictionary(Of String, String)
            If owner_id <> 0 Then
                parameters.Add("owner_id", owner_id)
            End If
            parameters.Add("no_service_albums", IIf(no_service_albums, "1", "0"))
            parameters.Add("extended", IIf(extended, "1", "0"))
            If count = 0 Then
                count = getAllCount(owner_id)
            End If
            If count > 0 Then
                With parameters
                    If count > 100 Then
                        result.AddRange(getAll(owner_id, no_service_albums, offset + 100, count - 100, extended))
                        .Add("count", 100)
                    Else
                        .Add("count", count)
                    End If
                End With
            End If
            If offset > 0 Then
                parameters.Add("offset", offset)
            End If
            result.AddRange(PhotoParser(api_request_no_parse("photos.getAll", parameters), "photos.getAll"))
            Return result
        End Function

        Public Shared Function getUserPhotos(Optional uid As UInteger = 0, Optional offset As UInt32 = 0, Optional count As UInt32 = 0, Optional sort As Boolean = False, Optional extended As Boolean = False) As List(Of photo)
            Dim result As New List(Of photo)
            Dim parameters As New Dictionary(Of String, String)

            parameters.Add("extended", IIf(extended, "1", "0"))
            parameters.Add("sort", IIf(sort, "1", "0"))
            If uid > 0 Then
                parameters.Add("uid", uid)
            End If
            If count = 0 Then
                count = getUserPhotosCount(uid)
            End If
            With parameters
                If count > 100 Then
                    result.AddRange(getUserPhotos(uid, offset + 100, count - 100, sort, extended))
                    .Add("count", 100)
                Else
                    .Add("count", count)
                End If
            End With
            If offset > 0 Then
                parameters.Add("offset", offset)
            End If
            result.AddRange(PhotoParser(api_request_no_parse("photos.getUserPhotos", parameters), "photos.getUserPhotos"))
            Return result
        End Function
        Public Shared Function getUserPhotosCount(Optional uid As UInteger = 0) As UInteger
            Dim parameters As New Dictionary(Of String, String)
            If uid > 0 Then parameters.Add("uid", uid)
            parameters.Add("count", 1)
            Dim doc As XmlDocument = api_request("photos.getUserPhotos", parameters)
            Dim count As Xml.XmlNode = doc.SelectNodes("/response/count").Item(0)
            Return Long.Parse(count.InnerText)

        End Function
        Public Overloads Shared Function getById(photo As photo, Optional extended As Boolean = False) As photo
            Dim l As New List(Of photo)
            l.Add(photo)
            Return getById(l, extended)(0)
        End Function
        Public Overloads Shared Function getById(photos As List(Of photo), Optional extended As Boolean = False) As List(Of photo)
            'Функция принимает фото с заполненными параметрами pid и owner_id, возвращает фото с полной информацией. Если нет доступа к фото, его не будет в списке
            Dim result As New List(Of photo)
            Dim parameters As New Dictionary(Of String, String)

            parameters.Add("extended", IIf(extended, "1", "0"))
            parameters.Add("photos", Nothing)
            Dim out_of_limit_photos As New List(Of photo)
            Dim parameters_new As Dictionary(Of String, String) = parameters, max_url_length_reached As Boolean = False
            For Each ph As photo In photos
                If Not max_url_length_reached Then
                    parameters_new("photos") = parameters("photos") & "," & CStr(ph.owner_id) & "_" & CStr(ph.pid)
                    If Len(api_get_url("photos.getById", parameters_new)) > 2000 Then
                        max_url_length_reached = True
                    Else
                        parameters = parameters_new
                    End If
                Else
                    out_of_limit_photos.Add(ph)
                End If
            Next
            If max_url_length_reached Then
                result.AddRange(getById(out_of_limit_photos, extended))
            End If
            If Not IsNothing(parameters("photos")) Then

                parameters("photos") = Right(parameters("photos"), Len(parameters("photos")) - 1)

            End If

            result.AddRange(PhotoParser(api_request_no_parse("photos.getById", parameters), "photos.getById"))

            Return result
        End Function
        Private Shared Function PhotoParser(s As XmlTextReader, Optional action As String = Nothing) As List(Of photo)
            Dim result As New List(Of photo)
            While s.Read
                If s.NodeType = XmlNodeType.Element Then
                    If s.Name = "photo" Then
                        Dim photo As XmlReader = s.ReadSubtree, curr_photo As New photo
                        While photo.Read
                            If photo.NodeType = XmlNodeType.Element Then
                                With curr_photo
                                    Select Case photo.Name
                                        Case "pid"
                                            .pid = photo.ReadElementContentAsLong
                                        Case "aid"
                                            .aid = photo.ReadElementContentAsLong
                                        Case "owner_id"
                                            .owner_id = photo.ReadElementContentAsLong
                                        Case "created"
                                            .created = New Date(1970, 1, 1, 0, 0, 0, 0).AddSeconds(photo.ReadElementContentAsLong).ToLocalTime ' Unix time to Date conversion
                                        Case "src"
                                            .src = photo.ReadElementContentAsString
                                        Case "src_small"
                                            .src_small = photo.ReadElementContentAsString
                                        Case "src_big"
                                            .src_big = photo.ReadElementContentAsString
                                        Case "src_xbig"
                                            .src_xbig = photo.ReadElementContentAsString
                                        Case "src_xxbig"
                                            .src_xxbig = photo.ReadElementContentAsString
                                        Case "text"
                                            .text = Web.HttpUtility.HtmlDecode(photo.ReadElementContentAsString.Replace("<br>", vbNewLine))
                                        Case "error"
                                            error_handler_TextReader(s, action)
                                    End Select
                                End With
                            End If
                        End While
                        result.Add(curr_photo)
                    End If
                End If

            End While
            Return result
        End Function
        Public Class photo
            Public pid As UInt64, aid As Long, owner_id As Int64, created As Date, src As String, src_small As String, src_big As String, src_xbig As String, src_xxbig As String, text As String
            Public ReadOnly Property src_largest As String
                Get
                    If Len(src_xxbig) > 0 Then
                        Return src_xxbig
                    ElseIf Len(src_xbig) > 0 Then
                        Return src_xbig
                    Else
                        Return src_big
                    End If
                End Get
            End Property

        End Class
        Public Class album
            Public aid As Long, thumb_id As UInt64, owner_id As Int64, title As String, description As String, created As Date, updated As Date, size As UInt32, privacy As Integer, thumb_src As String
        End Class
        Public Shared Function getAlbums(Optional owner_id As Int64 = 0, Optional album_ids As List(Of Long) = Nothing, Optional need_system As Boolean = True, Optional need_covers As Boolean = False, Optional photo_sizes As Boolean = False, Optional offset As ULong = 0, Optional count As ULong = 0) As List(Of album)
            Dim result As New List(Of album)
            Dim parameters As New Dictionary(Of String, String)
            parameters.Add("v", "5.0")
            parameters.Add("owner_id", owner_id.ToString)
            If Not IsNothing(album_ids) Then
                parameters.Add("album_ids", fields_to_string(album_ids))
            End If
            If offset > 0 Then
                parameters.Add("offset", offset.ToString)
            End If
            If count > 0 Then
                parameters.Add("count", count.ToString)
            End If
            parameters.Add("need_system", IIf(need_system, "1", "0"))
            parameters.Add("need_covers", IIf(need_covers, "1", "0"))
            parameters.Add("photo_sizes", IIf(photo_sizes, "1", "0"))
            Dim al As XmlTextReader = api_request_no_parse("photos.getAlbums", parameters)

            While al.Read
                If al.NodeType = XmlNodeType.Element Then
                    If al.Name = "album" Then
                        Dim album = al.ReadSubtree, current_album As New album
                        While album.Read
                            If al.NodeType = XmlNodeType.Element Then
                                With current_album
                                    Select Case al.Name
                                        Case "id"
                                            .aid = al.ReadElementContentAsLong
                                        Case "thumb_id"
                                            .thumb_id = al.ReadElementContentAsLong
                                        Case "owner_id"
                                            .owner_id = al.ReadElementContentAsLong
                                        Case "title"
                                            .title = Web.HttpUtility.HtmlDecode(al.ReadElementContentAsString)
                                        Case "description"
                                            .description = al.ReadElementContentAsString
                                        Case "created"
                                            .created = New Date(1970, 1, 1, 0, 0, 0, 0).AddSeconds(al.ReadElementContentAsLong).ToLocalTime ' Unix time to Date conversion
                                        Case "updated"
                                            .updated = New Date(1970, 1, 1, 0, 0, 0, 0).AddSeconds(al.ReadElementContentAsLong).ToLocalTime
                                        Case "size"
                                            .size = al.ReadElementContentAsLong
                                        Case "privacy"
                                            .privacy = al.ReadElementContentAsLong
                                        Case "thumb_src"
                                            .thumb_src = al.ReadElementContentAsLong
                                        Case "sizes"
                                            Throw New Exception("Getting thumbnails sizes is not implemented")
                                    End Select
                                End With
                            End If
                        End While
                        result.Add(current_album)
                    End If
                End If
            End While
            Return result
        End Function
        Public Class photoComparer
            Implements IEqualityComparer(Of vk.api.photos.photo)
            Public Overloads Function Equals(x As vk.api.photos.photo, y As vk.api.photos.photo) As Boolean Implements System.Collections.Generic.IEqualityComparer(Of vk.api.photos.photo).Equals
                If x.pid = y.pid AndAlso x.owner_id = y.owner_id AndAlso x.src_largest = y.src_largest Then
                    Return True
                Else
                    Return False
                End If
            End Function
            Public Overloads Function GetHashCode(obj As vk.api.photos.photo) As Integer Implements System.Collections.Generic.IEqualityComparer(Of vk.api.photos.photo).GetHashCode
                Return obj.owner_id Xor obj.pid
            End Function
        End Class
        Public Class photoIdcomparer
            Implements IEqualityComparer(Of vk.api.photos.photo)
            Public Overloads Function Equals(x As vk.api.photos.photo, y As vk.api.photos.photo) As Boolean Implements System.Collections.Generic.IEqualityComparer(Of vk.api.photos.photo).Equals
                If x.pid = y.pid AndAlso x.owner_id = y.owner_id Then
                    Return True
                Else
                    Return False
                End If
            End Function
            Public Overloads Function GetHashCode(obj As vk.api.photos.photo) As Integer Implements System.Collections.Generic.IEqualityComparer(Of vk.api.photos.photo).GetHashCode
                Return obj.owner_id Xor obj.pid
            End Function
        End Class
    End Class
    Shared Sub error_handler_TextReader(s As XmlTextReader, Optional action As String = Nothing)
        Dim error_code As Long, error_message As String = Nothing
        While s.Read
            If s.Name = "error_code" Then
                error_code = s.ReadElementContentAsLong
            ElseIf s.Name = "error_msg" Then
                error_message = s.ReadElementContentAsString
            End If
        End While
        error_handler(error_code, error_message)
    End Sub
    Shared Sub error_handler(error_code As Long, error_message As String, Optional action As String = Nothing)
        Dim error_text As String = Nothing
        Select Case error_code
            Case 221
                If Not user_notified Then
                    MsgBox("Ошибка! Трансляция отключена на странице пользователя", MsgBoxStyle.Exclamation, "Ошибка!")
                    user_notified = True
                End If
                error_text = "Translation is disabled"
            Case 5
                GetNewToken()
                error_text = "Ошибка аутентификации в ВК"
                '  Case 15
                '    error_text = "User deactivated
            Case Else
                error_text = "Error " & error_code & ": " & Chr(34) & error_message & Chr(34) & " получена от ВК"

                '  Throw New Exception(error_text)

        End Select
        If Not IsNothing(action) Then
            error_text = error_text & " при попытке совершить " & action
        End If
        Dim ex As New Exception(error_text)
        ex.Data("error_code") = error_code
        Throw (ex)
    End Sub

    Public Class custom
        Public Shared playend As Date
        Public Shared Sub set_status_by_track_name(name As audio.track.trackname, Optional target_ids As List(Of Long) = Nothing)
            '   Dim vkaudio As New audio, vkstatus As New status
            Dim track_to_set As audio.track = Nothing, from_cache As Boolean = False
            For Each track In api.current_user.tracks
                If track.Key = name Then
                    If debug Then
                        notify("Аудиозапись найдена в кэше")
                    End If
                    track_to_set = track.Value
                    from_cache = True
                    Exit For
                End If
            Next

            If Not from_cache Then track_to_set = searchtrack(name)

            If Not (IsNothing(track_to_set)) Then
                Try
                    audio.setBroadcast(track_to_set.id, track_to_set.owner_id, target_ids)

                    playend = Now.AddSeconds(track_to_set.duration)
                    If Not from_cache Then
                        current_user.tracks.Add(track_to_set.name, track_to_set)
                    End If
                Catch ex As Exception
                    notify("Не удалось установить в статус трек: " & ex.Message)
                    If from_cache Then
                        track_to_set = searchtrack(name)

                        Try
                            audio.setBroadcast(track_to_set.id, track_to_set.owner_id, target_ids)
                            playend = Now.AddSeconds(track_to_set.duration)
                            current_user.tracks.Remove(track_to_set.name)
                            If debug Then
                                notify("Аудиозапись удалена из кэша")
                            End If
                        Catch e As Exception
                            notify("Не удалось установить в статус трек: " & e.Message)
                        End Try
                    End If


                End Try
            Else
                '     Debug.Print("Track not found on vk")
                '   notify("Track " & name.ToString & " not found on vk ")
                playend = Now.AddSeconds(9999999)
                Dim ex As New Exception("Запись " & name.ToString & " не найдена в ВК")
                ex.Data("reason") = "not found"
                Throw ex
            End If



        End Sub
        Private Shared Function searchtrack(name As audio.track.trackname) As audio.track
            Dim track_to_set As audio.track = Nothing, exact_found As Boolean = False
            Dim rslt As audio.SearchResult = audio.search(name.ToString, 1)
            If rslt.count > 0 Then
                track_to_set = rslt.tracks(0)
                If rslt.tracks(0).name = name Then
                    If debug Then
                        notify("Найдено точное совпадение аудиозаписи")
                    End If
                Else
                    If debug Then
                        notify("Первый результат в поиске " & rslt.tracks(0).name.ToString & " не совпадает с записью с last.fm " & name.ToString & ", запросим ещё 300 записей...")
                    End If
                    rslt = audio.search(name.ToString, 300, 0)
                    For Each tr In rslt.tracks
                        If tr.name = name Then
                            track_to_set = tr
                            exact_found = True
                            Exit For
                        End If
                    Next
                    If Not exact_found Then
                        Dim mindistance As Integer = Integer.MaxValue
                        For Each tr In rslt.tracks
                            Dim currdistance As Integer = LevenshteinDistance(tr.artist.ToString, name.artist.ToString) + LevenshteinDistance(tr.title.ToString, name.title.ToString)
                            If mindistance > currdistance Then
                                mindistance = currdistance
                                track_to_set = tr
                            End If
                        Next

                        notify("Трек с названием " & name.ToString & " не найден в ВК, в статус отправлен наиболее близкий результат: " & track_to_set.name.ToString)
                    End If
                End If

            End If
            Return track_to_set
        End Function
        Private Shared Function LevenshteinDistance(ByVal s As String, ByVal t As String) As Integer
            Dim n As Integer = s.Length
            Dim m As Integer = t.Length
            Dim d(n + 1, m + 1) As Integer
            If n = 0 Then
                Return m
            End If
            If m = 0 Then
                Return n
            End If
            Dim i As Integer
            Dim j As Integer
            For i = 0 To n
                d(i, 0) = i
            Next
            For j = 0 To m
                d(0, j) = j
            Next
            For i = 1 To n
                For j = 1 To m
                    Dim cost As Integer
                    If t(j - 1) = s(i - 1) Then
                        cost = 0
                    Else
                        cost = 1
                    End If
                    d(i, j) = Math.Min(Math.Min(d(i - 1, j) + 1, d(i, j - 1) + 1),
                         d(i - 1, j - 1) + cost)
                Next
            Next
            Return d(n, m)
        End Function
        Public Shared Function get_userid_from_string(txt As String) As Long
            Dim uid As Long = 0
            If txt = "" Then

            ElseIf ULong.TryParse(txt, uid) Then

            ElseIf txt.IndexOf("vk.com/") > -1 Then
                uid = get_userid_from_string(txt.Substring(txt.IndexOf("vk.com/") + 7))
            ElseIf txt.Substring(0, 2) = "id" Then
                Long.TryParse(txt.Substring(2, txt.Length - 2), uid)
            Else
                Try
                    ULong.TryParse(users._get(New List(Of String)({txt}))(0).uid, uid)
                Catch ex As Exception
                    If ex.Data("error_code") = 113 Then
                        uid = 0
                    Else
                        Throw ex
                    End If

                End Try
            End If
            Return uid
        End Function

    End Class

    Public Shared Sub Logout()
        Dim auth As New auth_window
        auth.WebBrowser1.Navigate(vkauthurl & "&revoke=1") 'This is an only known reliant way to actually make cookies ineffective
        'auth.WebBrowser1.Navigate("https://oauth.vk.com/oauth/logout?hash=f398e2d9497d7a6d53&success_url=Y2xpZW50X2lkPTM0MDEwMDgmcmVkaXJlY3RfdXJpPWh0dHAlM0ElMkYlMkZvYXV0aC52ay5jb20lMkZibGFuay5odG1sJnJlc3BvbnNlX3R5cGU9dG9rZW4mc2NvcGU9NCZzdGF0ZT0mcmV2b2tlPTEmZGlzcGxheT1wYWdl&success_hash=dce8f92b2cb731189b")

        token = Nothing
        Application.DoEvents()
    End Sub

End Class



