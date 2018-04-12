'  ===============================================================================
' |    THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF      |
' |    ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO    |
' |    THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A         |
' |    PARTICULAR PURPOSE.                                                    |
' |    Copyright (c)2006-2012  ADMINSYSTEM SOFTWARE LIMITED                         |
' |
' |    Project: It demonstrates how to use EAGetMail to receive/parse non-delivery report.
' |        
' |
' |    File: Form1 : implementation file
' |
' |    Author: Ivan Lui ( ivan@emailarchitect.net )
'  ===============================================================================
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports EAGetMail

Public Class Form1
    Inherits System.Windows.Forms.Form
    Implements IComparer

    Private m_arUidl As ArrayList = New ArrayList
    Private m_bcancel As Boolean = False
    Private m_uidlfile As String = "uidl.txt"
    Friend WithEvents textReport As System.Windows.Forms.TextBox
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

            oMail.Clear()
        Next
    End Sub

    ' Report non-delivery report
    Private Sub ShowReport(ByVal fileName As String)
        Try
            Dim oMail As New Mail("TryIt")
            oMail.Load(fileName, False)

            If Not oMail.IsReport Then
                MessageBox.Show("This is not a report")
                Return
            End If

            Dim s As String = ""
            Dim oReport As MailReport = oMail.GetReport()
            Select Case oReport.ReportType
                Case DeliveryReportType.DeliveryReceipt
                    s = "This is a deliver receipt." & vbCr & vbLf & vbCr & vbLf
                    Exit Select
                Case DeliveryReportType.ReadReceipt
                    s = "This is a read receipt." & vbCr & vbLf & vbCr & vbLf
                    Exit Select
                Case Else
                    s = "This is a failure report." & vbCr & vbLf & vbCr & vbLf
                    Exit Select
            End Select

            s += [String].Format("Original Sender: {0}" & vbCr & vbLf, oReport.OriginalSender)
            s += [String].Format("Original Recipient: {0}" & vbCr & vbLf, oReport.OriginalRecipient)
            s += [String].Format("Original Message-ID: {0}" & vbCr & vbLf & vbCr & vbLf, oReport.OriginalMessageID)

            If oReport.ReportType = DeliveryReportType.FailureReport Then
                s += [String].Format("Original Subject: {0}" & vbCr & vbLf, oReport.OriginalSubject)
                s += [String].Format("Report MTA: {0}" & vbCr & vbLf, oReport.ReportMTA)
                s += [String].Format("Error Code: {0}" & vbCr & vbLf, oReport.ErrCode)
                s += [String].Format("Error Description: {0}" & vbCr & vbLf & vbCr & vbLf, oReport.ErrDescription)

                s += "---- Original Message Header ----" & vbCr & vbLf & vbCr & vbLf
                Dim oHeaders As HeaderCollection = oReport.OriginalHeaders
                Dim count As Integer = oHeaders.Count
                For i As Integer = 0 To count - 1
                    Dim oHeader As HeaderItem = TryCast(oHeaders(i), HeaderItem)
                    s += [String].Format("{0}: {1}" & vbCr & vbLf, oHeader.HeaderKey, oHeader.HeaderValue)
                Next
            End If


            textReport.Text = s
        Catch ep As Exception
            MessageBox.Show(ep.Message)
        End Try
    End Sub
#End Region

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    Friend WithEvents groupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents label1 As System.Windows.Forms.Label
    Friend WithEvents label2 As System.Windows.Forms.Label
    Friend WithEvents label3 As System.Windows.Forms.Label
    Friend WithEvents textServer As System.Windows.Forms.TextBox
    Friend WithEvents textUser As System.Windows.Forms.TextBox
    Friend WithEvents textPassword As System.Windows.Forms.TextBox
    Friend WithEvents label4 As System.Windows.Forms.Label
    Friend WithEvents lstAuthType As System.Windows.Forms.ComboBox
    Friend WithEvents label5 As System.Windows.Forms.Label
    Friend WithEvents lstProtocol As System.Windows.Forms.ComboBox
    Friend WithEvents chkLeaveCopy As System.Windows.Forms.CheckBox
    Friend WithEvents btnStart As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents pgBar As System.Windows.Forms.ProgressBar
    Friend WithEvents chkSSL As System.Windows.Forms.CheckBox

    Friend WithEvents lstMail As System.Windows.Forms.ListView
    Friend WithEvents colFrom As System.Windows.Forms.ColumnHeader
    Friend WithEvents colSubject As System.Windows.Forms.ColumnHeader
    Friend WithEvents colDate As System.Windows.Forms.ColumnHeader
    Friend WithEvents btnDel As System.Windows.Forms.Button
    Friend WithEvents lblTotal As System.Windows.Forms.Label
    Friend WithEvents label6 As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.groupBox1 = New System.Windows.Forms.GroupBox
        Me.pgBar = New System.Windows.Forms.ProgressBar
        Me.lblStatus = New System.Windows.Forms.Label
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnStart = New System.Windows.Forms.Button
        Me.chkLeaveCopy = New System.Windows.Forms.CheckBox
        Me.lstProtocol = New System.Windows.Forms.ComboBox
        Me.label5 = New System.Windows.Forms.Label
        Me.lstAuthType = New System.Windows.Forms.ComboBox
        Me.label4 = New System.Windows.Forms.Label
        Me.chkSSL = New System.Windows.Forms.CheckBox
        Me.textPassword = New System.Windows.Forms.TextBox
        Me.textUser = New System.Windows.Forms.TextBox
        Me.textServer = New System.Windows.Forms.TextBox
        Me.label3 = New System.Windows.Forms.Label
        Me.label2 = New System.Windows.Forms.Label
        Me.label1 = New System.Windows.Forms.Label
        Me.lstMail = New System.Windows.Forms.ListView
        Me.colFrom = New System.Windows.Forms.ColumnHeader
        Me.colSubject = New System.Windows.Forms.ColumnHeader
        Me.colDate = New System.Windows.Forms.ColumnHeader
        Me.btnDel = New System.Windows.Forms.Button
        Me.lblTotal = New System.Windows.Forms.Label
        Me.label6 = New System.Windows.Forms.Label
        Me.textReport = New System.Windows.Forms.TextBox
        Me.groupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'groupBox1
        '
        Me.groupBox1.Controls.Add(Me.pgBar)
        Me.groupBox1.Controls.Add(Me.lblStatus)
        Me.groupBox1.Controls.Add(Me.btnCancel)
        Me.groupBox1.Controls.Add(Me.btnStart)
        Me.groupBox1.Controls.Add(Me.chkLeaveCopy)
        Me.groupBox1.Controls.Add(Me.lstProtocol)
        Me.groupBox1.Controls.Add(Me.label5)
        Me.groupBox1.Controls.Add(Me.lstAuthType)
        Me.groupBox1.Controls.Add(Me.label4)
        Me.groupBox1.Controls.Add(Me.chkSSL)
        Me.groupBox1.Controls.Add(Me.textPassword)
        Me.groupBox1.Controls.Add(Me.textUser)
        Me.groupBox1.Controls.Add(Me.textServer)
        Me.groupBox1.Controls.Add(Me.label3)
        Me.groupBox1.Controls.Add(Me.label2)
        Me.groupBox1.Controls.Add(Me.label1)
        Me.groupBox1.Location = New System.Drawing.Point(8, 8)
        Me.groupBox1.Name = "groupBox1"
        Me.groupBox1.Size = New System.Drawing.Size(232, 328)
        Me.groupBox1.TabIndex = 0
        Me.groupBox1.TabStop = False
        Me.groupBox1.Text = "Account Information"
        '
        'pgBar
        '
        Me.pgBar.Location = New System.Drawing.Point(8, 312)
        Me.pgBar.Name = "pgBar"
        Me.pgBar.Size = New System.Drawing.Size(216, 8)
        Me.pgBar.TabIndex = 15
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Location = New System.Drawing.Point(8, 272)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(0, 13)
        Me.lblStatus.TabIndex = 14
        '
        'btnCancel
        '
        Me.btnCancel.Enabled = False
        Me.btnCancel.Location = New System.Drawing.Point(128, 240)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(88, 24)
        Me.btnCancel.TabIndex = 13
        Me.btnCancel.Text = "Cancel"
        '
        'btnStart
        '
        Me.btnStart.Location = New System.Drawing.Point(32, 240)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(88, 24)
        Me.btnStart.TabIndex = 12
        Me.btnStart.Text = "Start"
        '
        'chkLeaveCopy
        '
        Me.chkLeaveCopy.Location = New System.Drawing.Point(8, 208)
        Me.chkLeaveCopy.Name = "chkLeaveCopy"
        Me.chkLeaveCopy.Size = New System.Drawing.Size(208, 16)
        Me.chkLeaveCopy.TabIndex = 11
        Me.chkLeaveCopy.Text = "Leave a copy of message on server"
        '
        'lstProtocol
        '
        Me.lstProtocol.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.lstProtocol.Location = New System.Drawing.Point(80, 176)
        Me.lstProtocol.Name = "lstProtocol"
        Me.lstProtocol.Size = New System.Drawing.Size(136, 21)
        Me.lstProtocol.TabIndex = 10
        '
        'label5
        '
        Me.label5.AutoSize = True
        Me.label5.Location = New System.Drawing.Point(8, 178)
        Me.label5.Name = "label5"
        Me.label5.Size = New System.Drawing.Size(46, 13)
        Me.label5.TabIndex = 9
        Me.label5.Text = "Protocol"
        '
        'lstAuthType
        '
        Me.lstAuthType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.lstAuthType.Location = New System.Drawing.Point(80, 144)
        Me.lstAuthType.Name = "lstAuthType"
        Me.lstAuthType.Size = New System.Drawing.Size(136, 21)
        Me.lstAuthType.TabIndex = 8
        '
        'label4
        '
        Me.label4.AutoSize = True
        Me.label4.Location = New System.Drawing.Point(8, 146)
        Me.label4.Name = "label4"
        Me.label4.Size = New System.Drawing.Size(56, 13)
        Me.label4.TabIndex = 7
        Me.label4.Text = "Auth Type"
        '
        'chkSSL
        '
        Me.chkSSL.Location = New System.Drawing.Point(8, 120)
        Me.chkSSL.Name = "chkSSL"
        Me.chkSSL.Size = New System.Drawing.Size(208, 16)
        Me.chkSSL.TabIndex = 6
        Me.chkSSL.Text = "SSL connection"
        '
        'textPassword
        '
        Me.textPassword.Location = New System.Drawing.Point(80, 86)
        Me.textPassword.Name = "textPassword"
        Me.textPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.textPassword.Size = New System.Drawing.Size(136, 20)
        Me.textPassword.TabIndex = 5
        '
        'textUser
        '
        Me.textUser.Location = New System.Drawing.Point(80, 54)
        Me.textUser.Name = "textUser"
        Me.textUser.Size = New System.Drawing.Size(136, 20)
        Me.textUser.TabIndex = 4
        '
        'textServer
        '
        Me.textServer.Location = New System.Drawing.Point(80, 22)
        Me.textServer.Name = "textServer"
        Me.textServer.Size = New System.Drawing.Size(136, 20)
        Me.textServer.TabIndex = 3
        '
        'label3
        '
        Me.label3.AutoSize = True
        Me.label3.Location = New System.Drawing.Point(8, 88)
        Me.label3.Name = "label3"
        Me.label3.Size = New System.Drawing.Size(53, 13)
        Me.label3.TabIndex = 2
        Me.label3.Text = "Password"
        '
        'label2
        '
        Me.label2.AutoSize = True
        Me.label2.Location = New System.Drawing.Point(8, 56)
        Me.label2.Name = "label2"
        Me.label2.Size = New System.Drawing.Size(29, 13)
        Me.label2.TabIndex = 1
        Me.label2.Text = "User"
        '
        'label1
        '
        Me.label1.AutoSize = True
        Me.label1.Location = New System.Drawing.Point(8, 24)
        Me.label1.Name = "label1"
        Me.label1.Size = New System.Drawing.Size(38, 13)
        Me.label1.TabIndex = 0
        Me.label1.Text = "Server"
        '
        'lstMail
        '
        Me.lstMail.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.colFrom, Me.colSubject, Me.colDate})
        Me.lstMail.FullRowSelect = True
        Me.lstMail.HideSelection = False
        Me.lstMail.Location = New System.Drawing.Point(248, 16)
        Me.lstMail.Name = "lstMail"
        Me.lstMail.Size = New System.Drawing.Size(474, 168)
        Me.lstMail.TabIndex = 1
        Me.lstMail.UseCompatibleStateImageBehavior = False
        Me.lstMail.View = System.Windows.Forms.View.Details
        '
        'colFrom
        '
        Me.colFrom.Text = "From"
        Me.colFrom.Width = 100
        '
        'colSubject
        '
        Me.colSubject.Text = "Subject"
        Me.colSubject.Width = 200
        '
        'colDate
        '
        Me.colDate.Text = "Date"
        Me.colDate.Width = 150
        '
        'btnDel
        '
        Me.btnDel.Location = New System.Drawing.Point(650, 186)
        Me.btnDel.Name = "btnDel"
        Me.btnDel.Size = New System.Drawing.Size(72, 24)
        Me.btnDel.TabIndex = 3
        Me.btnDel.Text = "Delete"
        '
        'lblTotal
        '
        Me.lblTotal.AutoSize = True
        Me.lblTotal.Location = New System.Drawing.Point(256, 192)
        Me.lblTotal.Name = "lblTotal"
        Me.lblTotal.Size = New System.Drawing.Size(0, 13)
        Me.lblTotal.TabIndex = 4
        '
        'label6
        '
        Me.label6.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.label6.Location = New System.Drawing.Point(16, 352)
        Me.label6.Name = "label6"
        Me.label6.Size = New System.Drawing.Size(216, 48)
        Me.label6.TabIndex = 5
        Me.label6.Text = "Warning: if ""leave a copy of message on server"" is not checked,  the emails on th" & _
            "e server will be deleted !"
        '
        'textReport
        '
        Me.textReport.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.textReport.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.textReport.Location = New System.Drawing.Point(248, 216)
        Me.textReport.Multiline = True
        Me.textReport.Name = "textReport"
        Me.textReport.Size = New System.Drawing.Size(473, 183)
        Me.textReport.TabIndex = 6
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(734, 412)
        Me.Controls.Add(Me.textReport)
        Me.Controls.Add(Me.label6)
        Me.Controls.Add(Me.lblTotal)
        Me.Controls.Add(Me.btnDel)
        Me.Controls.Add(Me.lstMail)
        Me.Controls.Add(Me.groupBox1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.groupBox1.ResumeLayout(False)
        Me.groupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        lstProtocol.Items.Add("POP3")
        lstProtocol.Items.Add("IMAP4")
        lstProtocol.Items.Add("Exchange Web Service - 2007/2010")
        lstProtocol.Items.Add("Exchange WebDAV - Exchange 2000/2003")
        lstProtocol.SelectedIndex = 0

        lstAuthType.Items.Add("USER/LOGIN")
        lstAuthType.Items.Add("APOP")
        lstAuthType.Items.Add("NTLM")
        lstAuthType.SelectedIndex = 0

        Dim path As String = Application.ExecutablePath
        Dim pos As Integer = path.LastIndexOf("\")
        If pos <> -1 Then
            path = path.Substring(0, pos)
        End If
        m_curpath = path

        lstMail.Sorting = SortOrder.Descending
        lstMail.ListViewItemSorter = Me

        LoadMails()
        lblTotal.Text = String.Format("Total {0} delivery report(s)", lstMail.Items.Count)
    End Sub

#Region "IComparer Members"
    'sort the email as received data.
    Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
        Dim itemx As ListViewItem = x
        Dim itemy As ListViewItem = y

        Dim sx() As Char = itemx.SubItems(2).Text.ToCharArray()
        Dim sy() As Char = itemy.SubItems(2).Text.ToCharArray()
        If sx.Length <> sy.Length Then
            Compare = -1
            Exit Function 'should never occured.
        End If

        Dim count As Integer = sx.Length
        For i As Integer = 0 To count - 1

            If sx(i) > sy(i) Then
                Compare = -1
                Exit Function
            ElseIf sx(i) < sy(i) Then
                Compare = 1
                Exit Function
            End If
        Next
        Compare = 0
    End Function

#End Region

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        m_bcancel = True
    End Sub

    Private Sub btnDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDel.Click
        Dim items As ListView.SelectedListViewItemCollection = lstMail.SelectedItems
        If items.Count = 0 Then
            Exit Sub
        End If

        If MessageBox.Show("Do you want to delete all selected emails", _
                                 "", _
                                 MessageBoxButtons.YesNo) = DialogResult.No Then
            Exit Sub
        End If

        Do While (items.Count > 0)
            Try
                Dim fileName As String = items(0).Tag
                File.Delete(fileName)
                Dim pos As Integer = fileName.LastIndexOf(".")
                Dim tempFolder As String = fileName.Substring(0, pos)
                Dim htmlName As String = tempFolder + ".htm"
                If (File.Exists(htmlName)) Then
                    File.Delete(htmlName)
                End If

                If (Directory.Exists(tempFolder)) Then
                    Directory.Delete(tempFolder, True)
                End If
                lstMail.Items.Remove(items(0))
            Catch ep As Exception
                MessageBox.Show(ep.Message)
                Exit Do
            End Try
        Loop

        lblTotal.Text = String.Format("Total {0} email(s)", lstMail.Items.Count)

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

        ' UIDL is the identifier of every email on POP3/IMAP4/Exchange server, to avoid retrieve
        ' the same email from server more than once, we record the email UIDL retrieved every time
        ' if you delete the email from server every time and not to leave a copy of email on
        ' the server, then please remove all the function about uidl.
        ' UIDLManager wraps the function to write/read uidl record from a text file.
        Dim oUIDLManager As New UIDLManager()

        Try
            ' load existed uidl records to UIDLManager
            Dim uidlfile As String = [String].Format("{0}\{1}", m_curpath, m_uidlfile)
            oUIDLManager.Load(uidlfile)

            Dim mailFolder As String = [String].Format("{0}\inbox", m_curpath)
            If Not Directory.Exists(mailFolder) Then
                Directory.CreateDirectory(mailFolder)
            End If

            m_bcancel = False
            lblStatus.Text = "Connecting ..."
            oClient.Connect(oServer)
            Dim infos As MailInfo() = oClient.GetMailInfos()
            lblStatus.Text = [String].Format("Total {0} report(s)", infos.Length)

            ' remove the local uidl which is not existed on the server.
            oUIDLManager.SyncUIDL(oServer, infos)
            oUIDLManager.Update()

            Dim count As Integer = infos.Length

            Dim arReport As New ArrayList()
            For i As Integer = 0 To count - 1
                Dim info As MailInfo = infos(i)
                If oUIDLManager.FindUIDL(oServer, info.UIDL) IsNot Nothing Then
                    ' This email has been downloaded or checked before.
                    Continue For
                End If

                lblStatus.Text = [String].Format("Retrieving mail header {0}/{1}...", info.Index, count)

                Dim d As System.DateTime = System.DateTime.Now
                Dim cur As New System.Globalization.CultureInfo("en-US")
                Dim sdate As String = d.ToString("yyyyMMddHHmmss", cur)
                Dim fileName As String = [String].Format("{0}\{1}{2}{3}.eml", mailFolder, sdate, d.Millisecond.ToString("d3"), i)

                Dim oMail As New Mail("TryIt")
                oMail.Load(oClient.GetMailHeader(info))

                If Not oMail.IsReport Then
                    ' Not a report, continue
                    ' Add the email uidl to uidl file to avoid we check it next time. 
                    oUIDLManager.AddUIDL(oServer, info.UIDL, fileName)
                    Continue For
                End If

                ' This is a report, get the entire email.
                oMail = oClient.GetMail(info)
                oMail.SaveAs(fileName, True)

                Dim item As New ListViewItem(oMail.From.ToString())
                item.SubItems.Add(oMail.Subject)
                item.SubItems.Add(oMail.ReceivedDate.ToString("yyyy-MM-dd HH:mm:ss"))
                item.Font = New System.Drawing.Font(item.Font, FontStyle.Bold)
                item.Tag = fileName
                lstMail.Items.Insert(0, item)
                oMail.Clear()

                lblTotal.Text = [String].Format("Total {0} email(s)", lstMail.Items.Count)

                arReport.Add(info)
                ' Add the report mail info to arraylist, 
                ' then we can delete it later.
                If bLeaveCopy Then
                    ' Add the email uidl to uidl file to avoid we retrieve it next time. 
                    oUIDLManager.AddUIDL(oServer, info.UIDL, fileName)
                End If
            Next

            If Not bLeaveCopy Then
                lblStatus.Text = "Deleting ..."
                count = arReport.Count
                For i As Integer = 0 To count - 1
                    Dim info As MailInfo = TryCast(arReport(i), MailInfo)
                    oClient.Delete(info)
                    ' Remove UIDL from local uidl file.
                    oUIDLManager.RemoveUIDL(oServer, info.UIDL)
                Next
            End If
            ' Delete method just mark the email as deleted, 
            ' Quit method pure the emails from server exactly.

            oClient.Quit()
        Catch ep As Exception
            MessageBox.Show(ep.Message)
        End Try

        ' Update the uidl list to local uidl text file and then we can load it next time.
        oUIDLManager.Update()

        lblStatus.Text = "Completed"
        pgBar.Maximum = 100
        pgBar.Minimum = 0
        pgBar.Value = 0
        btnStart.Enabled = True
        btnCancel.Enabled = False
    End Sub

    Private Sub lstMail_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstMail.Click
        Dim items As ListView.SelectedListViewItemCollection = lstMail.SelectedItems
        If items.Count = 0 Then
            Exit Sub
        End If
        Dim item As ListViewItem = items(0)
        ShowReport(item.Tag)
        item.Font = New System.Drawing.Font(item.Font, FontStyle.Regular)
    End Sub

    Private Sub Form1_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        If Me.Width < 750 Then
            Me.Width = 750
        End If

        If Me.Height < 450 Then
            Me.Height = 450
        End If

        lstMail.Width = Me.Width - 270
        textReport.Width = lstMail.Width
        btnDel.Left = Me.Width - (btnDel.Width + 20)
        textReport.Height = Me.Height - (lstMail.Height + 100)
    End Sub

    Private Sub lstProtocol_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstProtocol.SelectedIndexChanged
        ' By default, Exchange Web Service requires SSL connection.
        ' For other protocol, please set SSL connection based on your server setting manually
        If lstProtocol.SelectedIndex = ServerProtocol.ExchangeEWS Then
            chkSSL.Checked = True
        End If
    End Sub
End Class
