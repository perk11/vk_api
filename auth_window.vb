Imports System.Collections.Specialized
Imports System.Web
Imports vk.notification
Imports vk
Public Class auth_window
    Public showform As Boolean = False
    Private Sub WebBrowser1_DocumentCompleted(sender As System.Object, e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted
        Const successauthurl As String = "https://oauth.vk.com/blank.html"
        Dim currurl As String
        'Debug.Print("Document completed")
        currurl = WebBrowser1.Url.ToString
        If api.debug Then
            notify("Загружена страница " & currurl)
        End If
        If currurl = "about:blank" Then
            vk.api.showform = False
            Me.Dispose()
        ElseIf (IIf(Len(currurl) >= Len(successauthurl), Microsoft.VisualBasic.Left(currurl, Len(successauthurl)) = successauthurl, False)) Then

            ' Debug.Print("hey, auth!")
            If api.debug Then
                notify("Получен результат аутентификации!")
            End If
            If InStr(currurl, "error") > 0 Then
                notify("Ошибка авторизации... Пробуем снова")
                MsgBox("Ошибка авторизации!")
                WebBrowser1.Navigate(api.vkauthurl)
                WebBrowser1.Visible = True
            ElseIf InStr(currurl, "access_token=") > 0 Then
                '    Debug.Print("hey, we've got token!")
                If api.debug Then
                    notify("Получен токен!")
                End If
                '  MainWindow.started = True

                Dim parameters As NameValueCollection = HttpUtility.ParseQueryString(Microsoft.VisualBasic.Right(currurl, Len(currurl) - InStr(currurl, "#")))
#If Not Debug Then

                Try
#End If

                api.token = parameters.Item("access_token")
                api.tokenexpires = Now.AddSeconds(parameters.Item("expires_in"))
                api.current_user.uid = parameters("user_id")

                WebBrowser1.Navigate("about:blank")
#If Not Debug Then
                Catch
                    MsgBox("Неизвестная ошибка при попытке авторизации, попробуйте ещё раз")
                    WebBrowser1.Navigate(vk.api.vkauthurl)
                End Try
#End If
            Else
                notify("Что-то странное происходит. IE6? Адрес: " & currurl)
            End If

        ElseIf InStr(currurl, "navcancl.htm") > 0 Then
            notify("Failed to connect to vk, probably network error")
        ElseIf InStr(currurl, "cancel=1") Or InStr(currurl, "user_denied") > 0 Then
            notify("vk authorization canceled")
            WebBrowser1.Navigate(vk.api.vkauthurl)

        Else
            '  If MainWindow.Timer1.Enabled Then MainWindow.Timer1.Enabled = False
            '   vk_api.showform = True
            If api.debug Then
                notify("Загружена страница ВК, которую необходимо показать")
            End If
            '  Debug.Print("something else")
            ' MainWindow.started = False
            Me.TopMost = True
            Me.TopLevel = True
            Me.Show()
            '     Me.Activate()

        End If
    End Sub

    Private Sub vk_auth_window_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ' If MainWindow.Timer1.Enabled = False Then
        ' MainWindow.Timer1.Enabled = True
        ' End If
        api.showform = False
    End Sub

    Private Sub vk_auth_window_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

    End Sub




    Private Sub vk_auth_window_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub WebBrowser1_Navigating(sender As Object, e As System.Windows.Forms.WebBrowserNavigatingEventArgs) Handles WebBrowser1.Navigating
        Me.Visible = False
    End Sub

End Class