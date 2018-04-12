Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports EAGetMail

Public Class Form1
    Private m_arUidl As ArrayList = New ArrayList
    Private m_bcancel As Boolean = False
    Private m_uidlfile As String = "uidl.txt"
    Private m_curpath As String = ""

#Region "EAGetMail Event Handler"
    Private Sub OnConnected(ByVal sender As Object, ByRef cancel As Boolean)
        lblStatus.Text = "Connected ..."
        cancel = m_bcancel
        Application.DoEvents()
    End Sub

    Private Sub OnQuit(ByVal sender As Object, ByRef cancel As Boolean)
        lblStatus.Text = "Quit ..."
        cancel = m_bcancel
        Application.DoEvents()
    End Sub

    Private Sub OnReceivingDataStream(ByVal sender As Object, ByVal info As MailInfo, ByVal received As Integer, ByVal total As Integer, ByRef cancel As Boolean)
        pgBar.Minimum = 0
        pgBar.Maximum = total
        pgBar.Value = received
        cancel = m_bcancel
        Application.DoEvents()
    End Sub

    Private Sub OnIdle(ByVal sender As Object, ByRef cancel As Boolean)
        cancel = m_bcancel
        Application.DoEvents()
    End Sub

    Private Sub OnAuthorized(ByVal sender As Object, ByRef cancel As Boolean)
        lblStatus.Text = "Authorized ..."
        cancel = m_bcancel
        Application.DoEvents()
    End Sub

    Private Sub OnSecuring(ByVal sender As Object, ByRef cancel As Boolean)
        lblStatus.Text = "Securing ..."
        cancel = m_bcancel
        Application.DoEvents()
    End Sub
#End Region

#Region "UIDL Functions"
    ' uidl is the identifier of every email on POP3/IMAP4 server, to avoid retrieve
    ' the same email from server more than once, we record the email uidl retrieved every time
    ' if you delete the email from server every time and not to leave a copy of email on
    ' the server, then please remove all the function about uidl.
    Private Function _FindUIDL(ByVal infos() As MailInfo, ByVal uidl As String) As Boolean
        Dim count As Integer = infos.Length
        For i As Integer = 0 To count - 1
            If String.Compare(infos(i).UIDL, uidl, False) = 0 Then
                _FindUIDL = True
                Exit Function
            End If
        Next
        _FindUIDL = False
    End Function

    'remove the local uidl which is not existed on the server.
    Private Sub _SyncUIDL(ByVal oServer As MailServer, ByVal infos() As MailInfo)
        Dim s As String = String.Format("{0}#{1} ", oServer.Server, oServer.User)
        Dim bcontinue As Boolean = False
        Dim n As Integer = 0

        Do
            bcontinue = False
            Dim count As Integer = m_arUidl.Count
            For i As Integer = n To count - 1
                Dim x As String = m_arUidl(i)
                If String.Compare(s, 0, x, 0, s.Length, True) = 0 Then

                    Dim pos As Integer = x.LastIndexOf(" ")
                    If pos <> -1 Then
                        Dim uidl As String = x.Substring(pos + 1)
                        If (Not _FindUIDL(infos, uidl)) Then
                            'this uidl doesn't exist on server, 
                            'so we should remove it from local uidl list to save the storage.
                            bcontinue = True
                            n = i
                            m_arUidl.RemoveAt(i)
                            Exit For
                        End If
                    End If
                End If
            Next
        Loop While (bcontinue)

    End Sub

    Private Function _FindExistedUIDL(ByVal oServer As MailServer, ByVal uidl As String) As Boolean
        Dim s As String = String.Format("{0}#{1} {2}", oServer.Server.ToLower(), oServer.User.ToLower(), uidl)
        Dim count As Integer = m_arUidl.Count
        For i As Integer = 0 To count - 1
            Dim x As String = m_arUidl(i)
            If String.Compare(s, x, False) = 0 Then
                _FindExistedUIDL = True
                Exit Function
            End If
        Next
        _FindExistedUIDL = False
    End Function

    Private Sub _AddUIDL(ByVal oServer As MailServer, ByVal uidl As String)
        Dim s As String = String.Format("{0}#{1} {2}", oServer.Server.ToLower(), oServer.User.ToLower(), uidl)
        m_arUidl.Add(s)
    End Sub

    Private Sub _UpdateUIDL()
        Dim s As New StringBuilder
        Dim count As Integer = m_arUidl.Count
        For i As Integer = 0 To count - 1
            s.Append(m_arUidl(i))
            s.Append(vbCrLf)
        Next

        Dim file As String = String.Format("{0}\{1}", m_curpath, m_uidlfile)
        Dim fs As FileStream = Nothing

        Try
            fs = New FileStream(file, FileMode.Create, FileAccess.Write, FileShare.None)
            Dim data() As Byte = System.Text.Encoding.Default.GetBytes(s.ToString())
            fs.Write(data, 0, data.Length)
            fs.Close()
        Catch ep As Exception
            If Not (fs Is Nothing) Then
                fs.Close()
            End If
            Throw ep
        End Try
    End Sub

    Private Sub _LoadUIDL()

        m_arUidl.Clear()
        Dim filename As String = String.Format("{0}\{1}", m_curpath, m_uidlfile)
        Dim read As StreamReader = Nothing

        Try
            read = File.OpenText(filename)
            Do While (True)
                Dim line As String = read.ReadLine().Trim(vbCrLf & " " & vbTab.ToCharArray())
                m_arUidl.Add(line)
            Loop
        Catch ep As Exception
        End Try

        If Not (read Is Nothing) Then
            read.Close()
        End If
    End Sub

#End Region

#Region "Parse and Display Mails"
    Private Sub LoadMails()
        lstMail.Items.Clear()
        Dim mailFolder As String = String.Format("{0}\inbox", m_curpath)

        If (Not Directory.Exists(mailFolder)) Then
            Directory.CreateDirectory(mailFolder)
        End If

        Dim files() As String = Directory.GetFiles(mailFolder, "*.eml")
        Dim count As Integer = files.Length
        For i As Integer = 0 To count - 1
            Dim fullname As String = files(i)
            'For evaluation usage, please use "TryIt" as the license code, otherwise the 
            '"invalid license code" exception will be thrown. However, the object will expire in 1-2 months, then
            '"trial version expired" exception will be thrown.
            Dim oMail As New Mail("TryIt")

            'Load( file, true ) only load the email header to Mail object to save the CPU and memory
            ' the Mail object will load the whole email file later automatically if bodytext or attachment is required..
            oMail.Load(fullname, True)

            Dim item As ListViewItem = New ListViewItem(oMail.From.ToString())
            item.SubItems.Add(oMail.Subject)
            item.SubItems.Add(oMail.ReceivedDate.ToString("yyyy-MM-dd HH:mm:ss"))
            item.Tag = fullname
            lstMail.Items.Add(item)

            Dim pos As Integer = fullname.LastIndexOf(".")
            Dim mainName As String = fullname.Substring(0, pos)
            Dim htmlName As String = mainName + ".htm"
            If Not File.Exists(htmlName) Then
                ' this email is unread, we set the font style to bold.
                'item.Font = New System.Drawing.Font(item.Font, FontStyle.Bold)
            End If
            oMail.Clear()
        Next
    End Sub

    Private Function _FormatHtmlTag(ByVal src As String) As String
        src = src.Replace(">", "&gt;")
        src = src.Replace("<", "&lt;")
        _FormatHtmlTag = src
    End Function

    'we generate a html + attachment folder for every email, once the html is create,
    ' next time we don't need to parse the email again.
    Private Sub _GenerateHtmlForEmail(ByVal htmlName As String, ByVal emlFile As String, ByVal tempFolder As String)
        'For evaluation usage, please use "TryIt" as the license code, otherwise the 
        '"invalid license code" exception will be thrown. However, the object will expire in 1-2 months, then
        '"trial version expired" exception will be thrown.
        Dim oMail As New Mail("TryIt")
        oMail.Load(emlFile, False)



        Dim html As String = oMail.HtmlBody
        Dim hdr As New StringBuilder
        hdr.Append("<font face=""Courier New,Arial"" size=2>")
        hdr.Append("<b>From:</b> " + _FormatHtmlTag(oMail.From.ToString()) + "<br>")

        Dim addrs() As MailAddress = oMail.To
        Dim count As Integer = addrs.Length
        If (count > 0) Then
            hdr.Append("<b>To:</b> ")
            For i As Integer = 0 To count - 1
                hdr.Append(_FormatHtmlTag(addrs(i).ToString()))
                If (i < count - 1) Then
                    hdr.Append("")
                End If
            Next
            hdr.Append("<br>")
        End If

        addrs = oMail.Cc
        count = addrs.Length
        If (count > 0) Then
            hdr.Append("<b>Cc:</b> ")
            For i As Integer = 0 To count - 1

                hdr.Append(_FormatHtmlTag(addrs(i).ToString()))
                If (i < count - 1) Then
                    hdr.Append("")
                End If
            Next
            hdr.Append("<br>")
        End If

        hdr.Append(String.Format("<b>Subject:</b>{0}<br>" & vbCrLf, _FormatHtmlTag(oMail.Subject)))

        Dim atts() As Attachment = oMail.Attachments
        count = atts.Length
        If (count > 0) Then

            If (Not Directory.Exists(tempFolder)) Then
                Directory.CreateDirectory(tempFolder)
            End If

            hdr.Append("<b>Attachments:</b>")
            For i As Integer = 0 To count - 1
                Dim att As Attachment = atts(i)
                'this attachment is in OUTLOOK RTF format, decode it here.
                If (String.Compare(att.Name, "winmail.dat") = 0) Then
                    Dim tatts() As Attachment
                    Try
                        tatts = Mail.ParseTNEF(att.Content, True)
                        Dim y As Integer = tatts.Length
                        For x As Integer = 0 To y - 1
                            Dim tatt As Attachment = tatts(x)
                            Dim tattname As String = String.Format("{0}\{1}", tempFolder, tatt.Name)
                            tatt.SaveAs(tattname, True)
                            hdr.Append(String.Format("<a href=""{0}"" target=""_blank"">{1}</a> ", tattname, tatt.Name))
                        Next
                    Catch ep As Exception
                        MessageBox.Show(ep.Message)
                    End Try
                Else
                    Dim attname As String = String.Format("{0}\{1}", tempFolder, att.Name)
                    att.SaveAs(attname, True)
                    hdr.Append(String.Format("<a href=""{0}"" target=""_blank"">{1}</a> ", attname, att.Name))
                    If (att.ContentID.Length > 0) Then
                        'show embedded image.
                        html = html.Replace("cid:" + att.ContentID, attname)
                    ElseIf (String.Compare(att.ContentType, 0, "image/", 0, "image/".Length, True) = 0) Then
                        'show attached image.
                        html = html + String.Format("<hr><img src=""{0}"">", attname)
                    End If
                End If
            Next
        End If

        Dim reg As Regex = New Regex("(<meta[^>]*charset[ \t]*=[ \t""]*)([^<> \r\n""]*)", RegexOptions.Multiline Or RegexOptions.IgnoreCase)
        html = reg.Replace(html, "$1utf-8")
        If Not (reg.IsMatch(html)) Then
            hdr.Insert(0, "<meta HTTP-EQUIV=""Content-Type"" Content=""text-html charset=utf-8"">")
        End If

        html = hdr.ToString() + "<hr>" + html
        Dim fs As New FileStream(htmlName, FileMode.Create, FileAccess.Write, FileShare.None)
        Dim data() As Byte = System.Text.UTF8Encoding.UTF8.GetBytes(html)
        fs.Write(data, 0, data.Length)
        fs.Close()
        oMail.Clear()
    End Sub

    Private Sub ShowMail(ByVal fileName As String)

        Try
            Dim pos As Integer = fileName.LastIndexOf(".")
            Dim mainName As String = fileName.Substring(0, pos)
            Dim htmlName As String = mainName + ".htm"

            Dim tempFolder As String = mainName
            If Not (File.Exists(htmlName)) Then
                'we haven't generate the html for this email, generate it now.
                _GenerateHtmlForEmail(htmlName, fileName, tempFolder)
            End If
            webMail.Navigate(New System.Uri(htmlName))
        Catch ep As Exception
            MessageBox.Show(ep.Message)
        End Try
    End Sub
#End Region

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        webMail.Navigate(New System.Uri("about:blank"))

        lstProtocol.Items.Add("POP3")
        lstProtocol.Items.Add("IMAP4")
        lstProtocol.Items.Add("Exchange Web Service - 2007/2010")
        lstProtocol.Items.Add("Exchange WebDAV - Exchange 2000/2003")

        lstProtocol.SelectedIndex = 0

        lstAuthType.Items.Add("USER/LOGIN")
        lstAuthType.Items.Add("APOP")
        lstAuthType.Items.Add("NTLM")
        lstAuthType.SelectedIndex = 0

        Dim asb As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
        Dim ms() As System.Reflection.Module = asb.GetModules()

        Dim file As String = "pocketpc.mobile.vb.exe"
        Dim count As Integer = ms.Length
        For i As Integer = 0 To count - 1
            If String.Compare(file, ms(i).Name, True) = 0 Then
                file = ms(i).FullyQualifiedName
                Exit For
            End If
        Next

        Dim path As String = file
        Dim pos As Integer = path.LastIndexOf("\")
        If (pos <> -1) Then
            path = path.Substring(0, pos)
        End If
        m_curpath = path

        LoadMails()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        m_bcancel = True
    End Sub

    Private Sub btnDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDel.Click
        Dim items As ListView.SelectedIndexCollection = lstMail.SelectedIndices
        If items.Count = 0 Then
            Exit Sub
        End If

        If MessageBox.Show("Do you want to delete all selected emails", _
                                 "", _
                                 MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) = DialogResult.No Then
            Exit Sub
        End If

        Dim ar As New ArrayList()
        For i As Integer = 0 To items.Count - 1
            Dim index As Integer = items(i)
            Dim oItem As ListViewItem = lstMail.Items(index)
            ar.Add(oItem)
            Try
                Dim fileName As String = oItem.Tag
                File.Delete(fileName)
                Dim pos As Integer = fileName.LastIndexOf(".")
                Dim tempFolder As String = fileName.Substring(0, pos)
                Dim htmlName As String = tempFolder + ".htm"
                If File.Exists(htmlName) Then
                    File.Delete(htmlName)
                End If

                If (Directory.Exists(tempFolder)) Then

                    Directory.Delete(tempFolder, True)
                End If

            Catch ep As Exception
                MessageBox.Show(ep.Message)
                Exit For
            End Try
        Next

        For i As Integer = 0 To ar.Count - 1
            Dim oItem As ListViewItem = ar(i)
            lstMail.Items.Remove(oItem)
        Next
        webMail.Navigate(New System.Uri("about:blank"))
    End Sub

    Private Sub btnStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStart.Click
        Dim server, user, password As String
        server = textServer.Text.Trim()
        user = textUser.Text.Trim()
        password = textPassword.Text.Trim()

        If (server.Length = 0 Or user.Length = 0 Or password.Length = 0) Then
            MessageBox.Show("Please input server, user and password.")
            Exit Sub
        End If

        btnStart.Enabled = False
        btnCancel.Enabled = True

        Dim authType As ServerAuthType = ServerAuthType.AuthLogin
        If (lstAuthType.SelectedIndex = 1) Then
            authType = ServerAuthType.AuthCRAM5
        ElseIf (lstAuthType.SelectedIndex = 2) Then
            authType = ServerAuthType.AuthNTLM
        End If

        Dim protocol As ServerProtocol = lstProtocol.SelectedIndex

        Dim oServer As MailServer = New MailServer(server, user, password, _
    chkSSL.Checked, authType, protocol)

        'For evaluation usage, please use "TryIt" as the license code, otherwise the 
        '"invalid license code" exception will be thrown. However, the object will expire in 1-2 months, then
        '"trial version expired" exception will be thrown.
        Dim oClient As MailClient = New MailClient("TryIt")

        'Catching the following events is not necessary, 
        'just make the application more user friendly.
        'If you use the object in asp.net/windows service or non-gui application, 
        'You need not to catch the following events.
        'To learn more detail, please refer to the code in EAGetMail EventHandler region

        AddHandler oClient.OnAuthorized, AddressOf OnAuthorized
        AddHandler oClient.OnConnected, AddressOf OnConnected
        AddHandler oClient.OnIdle, AddressOf OnIdle
        AddHandler oClient.OnSecuring, AddressOf OnSecuring
        AddHandler oClient.OnReceivingDataStream, AddressOf OnReceivingDataStream
        Dim bLeaveCopy As Boolean = chkLeaveCopy.Checked

        Try
            ' uidl is the identifier of every email on POP3/IMAP4 server, to avoid retrieve
            ' the same email from server more than once, we record the email uidl retrieved every time
            ' if you delete the email from server every time and not to leave a copy of email on
            ' the server, then please remove all the function about uidl.
            _LoadUIDL()


            Dim mailFolder As String = String.Format("{0}\inbox", m_curpath)
            If Not Directory.Exists(mailFolder) Then
                Directory.CreateDirectory(mailFolder)
            End If

            m_bcancel = False
            oClient.Connect(oServer)
            Dim infos() As MailInfo = oClient.GetMailInfos()
            lblStatus.Text = String.Format("Total {0} email(s)", infos.Length)

            _SyncUIDL(oServer, infos)
            Dim count As Integer = infos.Length

            For i As Integer = 0 To count - 1
                Dim info As MailInfo = infos(i)
                If _FindExistedUIDL(oServer, info.UIDL) Then
                    'this email has existed on local disk.
                Else
                    lblStatus.Text = String.Format("Retrieving {0}/{1}...", info.Index, count)
                    Dim oMail As Mail = oClient.GetMail(info)
                    Dim d As System.DateTime = System.DateTime.Now
                    Dim cur As System.Globalization.CultureInfo = New System.Globalization.CultureInfo("en-US")
                    Dim sdate As String = d.ToString("yyyyMMddHHmmss", cur)
                    Dim fileName As String = String.Format("{0}\{1}{2}{3}.eml", mailFolder, sdate, d.Millisecond.ToString("d3"), i)
                    oMail.SaveAs(fileName, True)

                    Dim item As ListViewItem = New ListViewItem(oMail.From.ToString())
                    item.SubItems.Add(oMail.Subject)
                    item.SubItems.Add(oMail.ReceivedDate.ToString("yyyy-MM-dd HH:mm:ss"))
                    ' item.Font = New System.Drawing.Font(item.Font, FontStyle.Bold)
                    item.Tag = fileName
                    lstMail.Items.Insert(0, item)
                    oMail.Clear()


                    If (bLeaveCopy) Then
                        'add the email uidl to uidl file to avoid we retrieve it next time. 
                        _AddUIDL(oServer, info.UIDL)
                    End If
                End If
            Next

            If Not (bLeaveCopy) Then
                lblStatus.Text = "Deleting ..."
                For i As Integer = 0 To count - 1
                    oClient.Delete(infos(i))
                Next
            End If
            ' Delete method just mark the email as deleted, 
            ' Quit method pure the emails from server exactly.
            oClient.Quit()

        Catch ep As Exception
            MessageBox.Show(ep.Message)
        End Try

        'update the uidl list to a text file and then we can load it next time.
        _UpdateUIDL()

        lblStatus.Text = "Completed"
        pgBar.Maximum = 100
        pgBar.Minimum = 0
        pgBar.Value = 0
        btnStart.Enabled = True
        btnCancel.Enabled = False
    End Sub

    Private Sub lstMail_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstMail.SelectedIndexChanged
        Dim items As ListView.SelectedIndexCollection = lstMail.SelectedIndices
        If (items.Count = 0) Then
            Exit Sub
        End If

        Dim oItem As ListViewItem = lstMail.Items(items(0))
        ShowMail(oItem.Tag)

    End Sub
End Class
