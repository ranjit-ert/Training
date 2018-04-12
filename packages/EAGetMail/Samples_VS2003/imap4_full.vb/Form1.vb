'  ==============================================================================
' |    THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF      |
' |    ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO    |
' |    THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A         |
' |    PARTICULAR PURPOSE.                                                    |
' |    Copyright (c)2006-2012  ADMINSYSTEM SOFTWARE LIMITED                         |
' |
' |    Project: It demonstrates how to use EAGetMail to receive/parse email.
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

    Private oClient As MailClient = New MailClient("TryIt")
    Private oCurServer As MailServer = Nothing
    Private oUIDLManager As UIDLManager = Nothing
    Private m_bcancel As Boolean = False
    Private m_curpath As String = ""
    Private WithEvents btnQuit As System.Windows.Forms.Button
    Private WithEvents btnCopy As System.Windows.Forms.Button
    Private WithEvents label1 As System.Windows.Forms.Label
    Private WithEvents btnCancel As System.Windows.Forms.Button
    Private WithEvents btnStart As System.Windows.Forms.Button
    Private WithEvents lstProtocol As System.Windows.Forms.ComboBox
    Private WithEvents label5 As System.Windows.Forms.Label
    Private WithEvents lstAuthType As System.Windows.Forms.ComboBox
    Private WithEvents label4 As System.Windows.Forms.Label
    Private WithEvents chkSSL As System.Windows.Forms.CheckBox
    Private WithEvents btnUpload As System.Windows.Forms.Button
    Private WithEvents textPassword As System.Windows.Forms.TextBox
    Private WithEvents textUser As System.Windows.Forms.TextBox
    Private WithEvents textServer As System.Windows.Forms.TextBox
    Private WithEvents groupBox1 As System.Windows.Forms.GroupBox
    Private WithEvents label3 As System.Windows.Forms.Label
    Private WithEvents label2 As System.Windows.Forms.Label
    Private WithEvents label6 As System.Windows.Forms.Label

    Private WithEvents lblStatus As System.Windows.Forms.Label


    Private WithEvents openFileDialog1 As System.Windows.Forms.OpenFileDialog

    Private WithEvents lstMail As System.Windows.Forms.ListView
    Private WithEvents colFrom As System.Windows.Forms.ColumnHeader
    Private WithEvents colSubject As System.Windows.Forms.ColumnHeader
    Private WithEvents colDate As System.Windows.Forms.ColumnHeader
    Private WithEvents pgBar As System.Windows.Forms.ProgressBar
    Private WithEvents btnMove As System.Windows.Forms.Button
    Private WithEvents trAccounts As System.Windows.Forms.TreeView
    Private WithEvents btnPure As System.Windows.Forms.Button
    Private WithEvents btnUnread As System.Windows.Forms.Button
    Private WithEvents btnDelete As System.Windows.Forms.Button
    Private WithEvents btnUndelete As System.Windows.Forms.Button


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

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents webMail As AxSHDocVw.AxWebBrowser
    Friend WithEvents ContextMenuFolder As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form1))
        Me.btnQuit = New System.Windows.Forms.Button
        Me.btnCopy = New System.Windows.Forms.Button
        Me.label1 = New System.Windows.Forms.Label
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnStart = New System.Windows.Forms.Button
        Me.lstProtocol = New System.Windows.Forms.ComboBox
        Me.label5 = New System.Windows.Forms.Label
        Me.lstAuthType = New System.Windows.Forms.ComboBox
        Me.label4 = New System.Windows.Forms.Label
        Me.chkSSL = New System.Windows.Forms.CheckBox
        Me.btnUpload = New System.Windows.Forms.Button
        Me.textPassword = New System.Windows.Forms.TextBox
        Me.textUser = New System.Windows.Forms.TextBox
        Me.textServer = New System.Windows.Forms.TextBox
        Me.groupBox1 = New System.Windows.Forms.GroupBox
        Me.label3 = New System.Windows.Forms.Label
        Me.label2 = New System.Windows.Forms.Label
        Me.label6 = New System.Windows.Forms.Label
        Me.lblStatus = New System.Windows.Forms.Label
        Me.openFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.lstMail = New System.Windows.Forms.ListView
        Me.colFrom = New System.Windows.Forms.ColumnHeader
        Me.colSubject = New System.Windows.Forms.ColumnHeader
        Me.colDate = New System.Windows.Forms.ColumnHeader
        Me.pgBar = New System.Windows.Forms.ProgressBar
        Me.btnMove = New System.Windows.Forms.Button
        Me.trAccounts = New System.Windows.Forms.TreeView
        Me.ContextMenuFolder = New System.Windows.Forms.ContextMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.MenuItem4 = New System.Windows.Forms.MenuItem
        Me.MenuItem5 = New System.Windows.Forms.MenuItem
        Me.btnPure = New System.Windows.Forms.Button
        Me.btnUnread = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnUndelete = New System.Windows.Forms.Button
        Me.webMail = New AxSHDocVw.AxWebBrowser
        Me.groupBox1.SuspendLayout()
        CType(Me.webMail, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnQuit
        '
        Me.btnQuit.Location = New System.Drawing.Point(19, 254)
        Me.btnQuit.Name = "btnQuit"
        Me.btnQuit.Size = New System.Drawing.Size(208, 24)
        Me.btnQuit.TabIndex = 16
        Me.btnQuit.Text = "Quit"
        '
        'btnCopy
        '
        Me.btnCopy.Enabled = False
        Me.btnCopy.Location = New System.Drawing.Point(685, 176)
        Me.btnCopy.Name = "btnCopy"
        Me.btnCopy.Size = New System.Drawing.Size(56, 22)
        Me.btnCopy.TabIndex = 31
        Me.btnCopy.Text = "Copy"
        '
        'label1
        '
        Me.label1.AutoSize = True
        Me.label1.Location = New System.Drawing.Point(145, 272)
        Me.label1.Name = "label1"
        Me.label1.Size = New System.Drawing.Size(0, 16)
        Me.label1.TabIndex = 14
        '
        'btnCancel
        '
        Me.btnCancel.Enabled = False
        Me.btnCancel.Location = New System.Drawing.Point(19, 283)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(208, 24)
        Me.btnCancel.TabIndex = 13
        Me.btnCancel.Text = "Cancel Task"
        '
        'btnStart
        '
        Me.btnStart.Location = New System.Drawing.Point(18, 224)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(208, 24)
        Me.btnStart.TabIndex = 12
        Me.btnStart.Text = "Start"
        '
        'lstProtocol
        '
        Me.lstProtocol.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.lstProtocol.Location = New System.Drawing.Point(87, 183)
        Me.lstProtocol.Name = "lstProtocol"
        Me.lstProtocol.Size = New System.Drawing.Size(136, 21)
        Me.lstProtocol.TabIndex = 10
        '
        'label5
        '
        Me.label5.AutoSize = True
        Me.label5.Location = New System.Drawing.Point(15, 185)
        Me.label5.Name = "label5"
        Me.label5.Size = New System.Drawing.Size(46, 16)
        Me.label5.TabIndex = 9
        Me.label5.Text = "Protocol"
        '
        'lstAuthType
        '
        Me.lstAuthType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.lstAuthType.Location = New System.Drawing.Point(87, 151)
        Me.lstAuthType.Name = "lstAuthType"
        Me.lstAuthType.Size = New System.Drawing.Size(136, 21)
        Me.lstAuthType.TabIndex = 8
        '
        'label4
        '
        Me.label4.AutoSize = True
        Me.label4.Location = New System.Drawing.Point(15, 153)
        Me.label4.Name = "label4"
        Me.label4.Size = New System.Drawing.Size(56, 16)
        Me.label4.TabIndex = 7
        Me.label4.Text = "Auth Type"
        '
        'chkSSL
        '
        Me.chkSSL.Location = New System.Drawing.Point(18, 124)
        Me.chkSSL.Name = "chkSSL"
        Me.chkSSL.Size = New System.Drawing.Size(208, 16)
        Me.chkSSL.TabIndex = 6
        Me.chkSSL.Text = "SSL connection"
        '
        'btnUpload
        '
        Me.btnUpload.Enabled = False
        Me.btnUpload.Location = New System.Drawing.Point(810, 176)
        Me.btnUpload.Name = "btnUpload"
        Me.btnUpload.Size = New System.Drawing.Size(56, 22)
        Me.btnUpload.TabIndex = 30
        Me.btnUpload.Text = "Upload"
        '
        'textPassword
        '
        Me.textPassword.Location = New System.Drawing.Point(87, 88)
        Me.textPassword.Name = "textPassword"
        Me.textPassword.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.textPassword.Size = New System.Drawing.Size(136, 20)
        Me.textPassword.TabIndex = 5
        Me.textPassword.Text = ""
        '
        'textUser
        '
        Me.textUser.Location = New System.Drawing.Point(87, 56)
        Me.textUser.Name = "textUser"
        Me.textUser.Size = New System.Drawing.Size(136, 20)
        Me.textUser.TabIndex = 4
        Me.textUser.Text = ""
        '
        'textServer
        '
        Me.textServer.Location = New System.Drawing.Point(87, 24)
        Me.textServer.Name = "textServer"
        Me.textServer.Size = New System.Drawing.Size(136, 20)
        Me.textServer.TabIndex = 3
        Me.textServer.Text = ""
        '
        'groupBox1
        '
        Me.groupBox1.Controls.Add(Me.btnQuit)
        Me.groupBox1.Controls.Add(Me.label1)
        Me.groupBox1.Controls.Add(Me.btnCancel)
        Me.groupBox1.Controls.Add(Me.btnStart)
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
        Me.groupBox1.Controls.Add(Me.label6)
        Me.groupBox1.Location = New System.Drawing.Point(10, 5)
        Me.groupBox1.Name = "groupBox1"
        Me.groupBox1.Size = New System.Drawing.Size(232, 328)
        Me.groupBox1.TabIndex = 29
        Me.groupBox1.TabStop = False
        Me.groupBox1.Text = "Account Information"
        '
        'label3
        '
        Me.label3.AutoSize = True
        Me.label3.Location = New System.Drawing.Point(15, 90)
        Me.label3.Name = "label3"
        Me.label3.Size = New System.Drawing.Size(54, 16)
        Me.label3.TabIndex = 2
        Me.label3.Text = "Password"
        '
        'label2
        '
        Me.label2.AutoSize = True
        Me.label2.Location = New System.Drawing.Point(15, 58)
        Me.label2.Name = "label2"
        Me.label2.Size = New System.Drawing.Size(28, 16)
        Me.label2.TabIndex = 1
        Me.label2.Text = "User"
        '
        'label6
        '
        Me.label6.AutoSize = True
        Me.label6.Location = New System.Drawing.Point(15, 26)
        Me.label6.Name = "label6"
        Me.label6.Size = New System.Drawing.Size(38, 16)
        Me.label6.TabIndex = 0
        Me.label6.Text = "Server"
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Location = New System.Drawing.Point(10, 347)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(37, 16)
        Me.lblStatus.TabIndex = 27
        Me.lblStatus.Text = "Ready"
        '
        'openFileDialog1
        '
        Me.openFileDialog1.Filter = "Email File | *.EML"
        '
        'lstMail
        '
        Me.lstMail.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.colFrom, Me.colSubject, Me.colDate})
        Me.lstMail.FullRowSelect = True
        Me.lstMail.HideSelection = False
        Me.lstMail.Location = New System.Drawing.Point(425, 8)
        Me.lstMail.Name = "lstMail"
        Me.lstMail.Size = New System.Drawing.Size(446, 161)
        Me.lstMail.TabIndex = 20
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
        Me.colDate.Width = 100
        '
        'pgBar
        '
        Me.pgBar.Location = New System.Drawing.Point(425, 408)
        Me.pgBar.Name = "pgBar"
        Me.pgBar.Size = New System.Drawing.Size(445, 10)
        Me.pgBar.TabIndex = 26
        '
        'btnMove
        '
        Me.btnMove.Enabled = False
        Me.btnMove.Location = New System.Drawing.Point(748, 176)
        Me.btnMove.Name = "btnMove"
        Me.btnMove.Size = New System.Drawing.Size(56, 22)
        Me.btnMove.TabIndex = 25
        Me.btnMove.Text = "Move"
        '
        'trAccounts
        '
        Me.trAccounts.ContextMenu = Me.ContextMenuFolder
        Me.trAccounts.HideSelection = False
        Me.trAccounts.ImageIndex = -1
        Me.trAccounts.Location = New System.Drawing.Point(257, 8)
        Me.trAccounts.Name = "trAccounts"
        Me.trAccounts.SelectedImageIndex = -1
        Me.trAccounts.Size = New System.Drawing.Size(162, 408)
        Me.trAccounts.TabIndex = 19
        '
        'ContextMenuFolder
        '
        Me.ContextMenuFolder.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem2, Me.MenuItem3, Me.MenuItem4, Me.MenuItem5})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.Text = "Refresh Folders"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 1
        Me.MenuItem2.Text = "Refresh Mails"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 2
        Me.MenuItem3.Text = "Add Folder"
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 3
        Me.MenuItem4.Text = "Delete Folder"
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 4
        Me.MenuItem5.Text = "Rename Folder"
        '
        'btnPure
        '
        Me.btnPure.Enabled = False
        Me.btnPure.Location = New System.Drawing.Point(621, 176)
        Me.btnPure.Name = "btnPure"
        Me.btnPure.Size = New System.Drawing.Size(56, 22)
        Me.btnPure.TabIndex = 24
        Me.btnPure.Text = "Pure"
        '
        'btnUnread
        '
        Me.btnUnread.Enabled = False
        Me.btnUnread.Location = New System.Drawing.Point(559, 176)
        Me.btnUnread.Name = "btnUnread"
        Me.btnUnread.Size = New System.Drawing.Size(56, 22)
        Me.btnUnread.TabIndex = 23
        Me.btnUnread.Text = "Unread"
        '
        'btnDelete
        '
        Me.btnDelete.Enabled = False
        Me.btnDelete.Location = New System.Drawing.Point(425, 176)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(56, 22)
        Me.btnDelete.TabIndex = 21
        Me.btnDelete.Text = "Delete Message"
        '
        'btnUndelete
        '
        Me.btnUndelete.Enabled = False
        Me.btnUndelete.Location = New System.Drawing.Point(487, 176)
        Me.btnUndelete.Name = "btnUndelete"
        Me.btnUndelete.Size = New System.Drawing.Size(66, 22)
        Me.btnUndelete.TabIndex = 22
        Me.btnUndelete.Text = "Undelete"
        '
        'webMail
        '
        Me.webMail.Enabled = True
        Me.webMail.Location = New System.Drawing.Point(432, 208)
        Me.webMail.OcxState = CType(resources.GetObject("webMail.OcxState"), System.Windows.Forms.AxHost.State)
        Me.webMail.Size = New System.Drawing.Size(440, 184)
        Me.webMail.TabIndex = 32
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(884, 432)
        Me.Controls.Add(Me.webMail)
        Me.Controls.Add(Me.btnCopy)
        Me.Controls.Add(Me.btnUpload)
        Me.Controls.Add(Me.groupBox1)
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.lstMail)
        Me.Controls.Add(Me.pgBar)
        Me.Controls.Add(Me.btnMove)
        Me.Controls.Add(Me.trAccounts)
        Me.Controls.Add(Me.btnPure)
        Me.Controls.Add(Me.btnUnread)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnUndelete)
        Me.Name = "Form1"
        Me.ShowInTaskbar = False
        Me.Text = "Form1"
        Me.groupBox1.ResumeLayout(False)
        CType(Me.webMail, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

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

    Public Sub OnSendingDataStream(ByVal sender As Object, ByVal sent As Integer, ByVal total As Integer, ByRef cancel As Boolean)
        pgBar.Minimum = 0
        pgBar.Maximum = total
        pgBar.Value = sent
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

#Region "Parse and Display Email"
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

        If (oMail.IsEncrypted) Then
            Try
                'this email is encrypted, we decrypt it by user default certificate.
                ' you can also use specified certificate like this
                ' oCert = new Certificate()
                'oCert.Load("c:\test.pfx", "pfxpassword", Certificate.CertificateKeyLocation.CRYPT_USER_KEYSET)
                ' oMail = oMail.Decrypt( oCert )
                oMail = oMail.Decrypt(Nothing)
            Catch ep As Exception
                MessageBox.Show(ep.Message)
                oMail.Load(emlFile, False)
            End Try
        End If

        If (oMail.IsSigned) Then
            Try
                'this email is digital signed.
                Dim cert As EAGetMail.Certificate = oMail.VerifySignature()
                MessageBox.Show("This email contains a valid digital signature.")
                'you can add the certificate to your certificate storage like this
                'cert.AddToStore( Certificate.CertificateStoreLocation.CERT_SYSTEM_STORE_CURRENT_USER,"addressbook" );
                ' then you can use send the encrypted email back to this sender.
            Catch ep As Exception
                MessageBox.Show(ep.Message)
            End Try
        End If

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
                    hdr.Append(";")
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
                    hdr.Append(";")
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
            hdr.Insert(0, "<meta HTTP-EQUIV=""Content-Type"" Content=""text/html; charset=utf-8"">")
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
            webMail.Navigate(htmlName)
        Catch ep As Exception
            MessageBox.Show(ep.Message)
        End Try
    End Sub
#End Region

#Region "IMAP4/Exchange Folders Management"
    ' Refresh Folders
    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click
        Dim node As TreeNode = trAccounts.SelectedNode
        If node Is Nothing Then
            Return
        End If

        EnableIdle(False)
        Try
            While Not (node.Parent Is Nothing)
                node = node.Parent
            End While

            trAccounts.SelectedNode = Nothing
            lstMail.Items.Clear()

            ConnectServer(node)
            oClient.RefreshFolders()

            lblStatus.Text = "Refreshing Folders ..."
            Dim fds As Imap4Folder() = oClient.Imap4Folders
            ExpendFolders(fds, node)
            node.ExpandAll()
            lblStatus.Text = ""
        Catch ep As Exception
            MessageBox.Show(ep.Message)
            oClient.Close()
        End Try

        EnableIdle(True)
    End Sub

    ' Refresh mails
    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click
        Dim node As TreeNode = trAccounts.SelectedNode
        If node Is Nothing Then
            Return
        End If

        Try

            lstMail.Items.Clear()
            ConnectServer(node)
            oClient.RefreshMailInfos()
            lblStatus.Text = "Refreshing Mails ..."


            ShowNode()
        Catch ep As Exception
            MessageBox.Show(ep.Message)
            oClient.Close()
        End Try
    End Sub

    ' Add Folder
    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click
        Dim frm As New Form3
        If frm.ShowDialog() = DialogResult.OK Then
            Try
                Dim folder As String = frm.textFolder.Text
                Dim node As TreeNode = trAccounts.SelectedNode
                ConnectServer(node)
                Dim fd As Imap4Folder = Nothing
                If (node.Parent Is Nothing) Then
                    fd = oClient.CreateFolder(Nothing, folder)
                Else
                    fd = oClient.CreateFolder(node.Tag, folder)
                End If

                Dim newNode As TreeNode = node.Nodes.Add(fd.Name)

                newNode.Tag = fd
            Catch ep As Exception
                MessageBox.Show(ep.Message)
                oClient.Close()
            End Try
        End If
    End Sub

    ' Delete Folder
    Private Sub MenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem4.Click
        Dim node As TreeNode = trAccounts.SelectedNode
        If node Is Nothing Then
            Return
        End If

        If node.Parent Is Nothing Then
            Return
        End If

        EnableIdle(False)
        Try

            ConnectServer(node)
            oClient.DeleteFolder(node.Tag)
            Directory.Delete(GetFolderByNode(node), True)
            trAccounts.SelectedNode = Nothing
            node.Remove()

            lstMail.Items.Clear()
        Catch ep As Exception
            MessageBox.Show(ep.Message)
            oClient.Close()
        End Try

        EnableIdle(True)
    End Sub
    ' rename folder
    Private Sub MenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem5.Click
        Dim node As TreeNode = trAccounts.SelectedNode
        If (Not (node Is Nothing)) And (Not (node.Parent Is Nothing)) Then
            trAccounts.LabelEdit = True
            If Not node.IsEditing Then
                node.BeginEdit()
            End If
        End If
    End Sub
    Private Sub trAccounts_AfterLabelEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.NodeLabelEditEventArgs) Handles trAccounts.AfterLabelEdit
        If e.Label Is Nothing Then
            Return
        End If

        If e.Label.Length = 0 Then
            e.CancelEdit = True
            MessageBox.Show("Invalid folder name.The name cannot be blank")
            Return
        End If

        EnableIdle(False)
        Try
            Dim node As TreeNode = e.Node
            If Not (node.Tag Is Nothing) Then
                ' rename folder
                ConnectServer(node)
                Dim cur_localpath As String = GetFolderByNode(node)
                oClient.RenameFolder(node.Tag, e.Label)
                e.Node.EndEdit(False)
                Dim new_localpath As String = GetFolderByNode(node)


                Try
                    ' Try to rename local folder as well.
                    System.IO.Directory.Move(cur_localpath, new_localpath)

                Catch ex As Exception
                End Try
            End If
        Catch ep As Exception
            MessageBox.Show(ep.Message)
            oClient.Close()
            e.CancelEdit = True
        End Try

        EnableIdle(True)

    End Sub
#End Region


    Private Function GetFolderByNode(ByVal node As TreeNode) As String
        If (node.Parent Is Nothing) Then
            GetFolderByNode = ""
            Exit Function
        End If

        Dim fd As Imap4Folder = node.Tag
        While Not (node.Parent Is Nothing)
            node = node.Parent
        End While
        GetFolderByNode = String.Format("{0}\{1}\{2}", m_curpath, node.Text, fd.LocalPath)
    End Function

    ' Create local folder
    Private Sub _CreateFullFolder(ByVal folder As String)
        If (Directory.Exists(folder)) Then
            Exit Sub
        End If
        Dim pos As Integer = 0
        While ((pos = folder.IndexOf("\", pos)) <> -1)
            If (pos > 2) Then
                Dim s As String = folder.Substring(0, pos)
                If Not (Directory.Exists(s)) Then
                    Directory.CreateDirectory(s)
                End If
            End If
            pos = pos + 1
        End While

        If Not (Directory.Exists(folder)) Then
            Directory.CreateDirectory(folder)
        End If
    End Sub

    ' Connect server
    Private Sub ConnectServer(ByVal node As TreeNode)
        Dim cur_node As TreeNode = node
        While Not (node.Parent Is Nothing)
            node = node.Parent
        End While

        Dim server As MailServer = node.Tag
        Dim bReConnect As Boolean = False
        If oCurServer Is Nothing OrElse Not oClient.Connected Then
            bReConnect = True
        ElseIf Not (oCurServer Is Nothing) Then
            If [String].Compare(oCurServer.Server, server.Server, True) <> 0 OrElse [String].Compare(oCurServer.User, server.User, True) <> 0 Then
                bReConnect = True
            End If
        End If

        If bReConnect Then
            lblStatus.Text = "Connecting ... "
            oCurServer = server.Copy()
            m_bcancel = False
            oClient.Connect(server)
            If Not (cur_node.Parent Is Nothing) Then
                oClient.SelectFolder(cur_node.Tag)
            End If
        End If
    End Sub


    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        EnableIdle(False)

        lstAuthType.Items.Add("USER/LOGIN")
        lstAuthType.Items.Add("CRAM-MD5")
        lstAuthType.Items.Add("NTLM")
        lstAuthType.SelectedIndex = 0

        lstProtocol.Items.Add("IMAP4")
        lstProtocol.Items.Add("Exchange Web Service - 2007/2010")
        lstProtocol.Items.Add("Exchange WebDAV - Exchange 2000/2003")
        lstProtocol.SelectedIndex = 0

        Dim path As String = Application.ExecutablePath
        Dim pos As Integer = path.LastIndexOf("\"c)
        If pos <> -1 Then
            path = path.Substring(0, pos)
        End If

        m_curpath = path

        Dim empty As Object = System.Reflection.Missing.Value
        webMail.Navigate("about:blank")

        ' Catching the following events is not necessary, 
        ' just make the application more user friendly.
        ' If you use the object in asp.net/windows service or non-gui application, 
        ' You need not to catch the following events.
        ' To learn more detail, please refer to the code in EAGetMail EventHandler region
        AddHandler oClient.OnAuthorized, AddressOf OnAuthorized
        AddHandler oClient.OnConnected, AddressOf OnConnected
        AddHandler oClient.OnIdle, AddressOf OnIdle
        AddHandler oClient.OnSecuring, AddressOf OnSecuring
        AddHandler oClient.OnReceivingDataStream, AddressOf OnReceivingDataStream
        AddHandler oClient.OnSendingDataStream, AddressOf OnSendingDataStream

    End Sub

    Private Sub btnStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStart.Click
        EnableIdle(False)
        m_bcancel = False
        Dim server As String, user As String, password As String
        server = textServer.Text.Trim()
        user = textUser.Text.Trim()
        password = textPassword.Text.Trim()

        If server.Length = 0 OrElse user.Length = 0 OrElse password.Length = 0 Then
            MessageBox.Show("Please input server, user and password.")
            Return
        End If

        btnStart.Enabled = False

        trAccounts.Nodes.Clear()
        lstMail.Items.Clear()

        Dim authType As ServerAuthType = ServerAuthType.AuthLogin
        If lstAuthType.SelectedIndex = 1 Then
            authType = ServerAuthType.AuthCRAM5
        ElseIf lstAuthType.SelectedIndex = 2 Then
            authType = ServerAuthType.AuthNTLM
        End If

        Dim protocol As ServerProtocol = lstProtocol.SelectedIndex + 1

        Dim oServer As New MailServer(server, user, password, chkSSL.Checked, authType, protocol)

        Try
            btnCancel.Enabled = True
            ' Enable log file 
            ' oClient.LogFileName = "d:\imap.txt"
            oClient.Connect(oServer)

            Dim node As New TreeNode([String].Format("{0}\{1}", oServer.Server.ToLower(), oServer.User.ToLower()))
            node.Tag = oServer
            oCurServer = oServer.Copy()

            Dim nodes As TreeNodeCollection = trAccounts.Nodes
            nodes.Add(node)
            trAccounts.SelectedNode = node

            ShowNode()

            EnableIdle(True)
        Catch ep As Exception
            btnStart.Enabled = True
            lblStatus.Text = ep.Message
            MessageBox.Show(ep.Message)

            btnCancel.Enabled = False
        End Try
    End Sub

    ' quit and close current connection
    Private Sub btnQuit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnQuit.Click
        Try
            btnStart.Enabled = True
            trAccounts.Nodes.Clear()
            lstMail.Items.Clear()
            webMail.Navigate("about:blank")
            oClient.Logout()
            oClient.Close()
            lblStatus.Text = "Discconnected"
            trAccounts.SelectedNode = Nothing

            Application.DoEvents()
        Catch ep As Exception
        End Try

        EnableIdle(False)
    End Sub

    ' Terminate current operation
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        m_bcancel = True
    End Sub

    Private Sub ExpendFolders(ByVal fds() As Imap4Folder, ByVal node As TreeNode)
        node.Nodes.Clear()
        Dim count As Integer = fds.Length
        For i As Integer = 0 To count - 1
            Dim fd As Imap4Folder = fds(i)
            Dim subNode As TreeNode = node.Nodes.Add(fd.Name)
            subNode.Tag = fd
            ExpendFolders(fd.SubFolders, subNode)
        Next
    End Sub

    ' Clear local file that is not existed on server.
    Private Sub _ClearLocalMails(ByVal infos As MailInfo(), ByVal localfolder As String)
        Dim files As String() = Directory.GetFiles(localfolder, "*.eml")
        Dim count As Integer = files.Length
        For i As Integer = 0 To count - 1
            Dim s As String = files(i)
            Dim pos As Integer = s.LastIndexOf("\")
            If pos <> -1 Then
                s = s.Substring(pos + 1)
            End If

            Dim bfind As Boolean = False

            Dim item As UIDLItem = oUIDLManager.FindLocalFile(s)
            If Not (item Is Nothing) Then
                bfind = True
            End If

            If Not bfind Then
                Dim fileName As String = files(i)
                File.Delete(fileName)
                Dim p As Integer = fileName.LastIndexOf(".")
                Dim tempFolder As String = fileName.Substring(0, p)
                Dim htmlName As String = tempFolder & ".htm"
                If File.Exists(htmlName) Then
                    File.Delete(htmlName)
                End If

                If Directory.Exists(tempFolder) Then
                    Directory.Delete(tempFolder, True)
                End If
            End If
        Next
    End Sub

    ' Get email list from server and diplay it in treeview.
    Private Sub LoadServerMails(ByVal node As TreeNode, ByVal fd As Imap4Folder)
        lstMail.Items.Clear()
        lstMail.Sorting = SortOrder.Descending
        lstMail.ListViewItemSorter = Me

        ConnectServer(node)
        Dim localfolder As String = GetFolderByNode(node)
        _CreateFullFolder(localfolder)

        lblStatus.Text = "Refreshing email(s) ..."
        oClient.SelectFolder(fd)
        Dim infos As MailInfo() = oClient.GetMailInfos()

        ' UIDL is the identifier of every email on POP3/IMAP4/Exchange server, to avoid retrieve
        ' the same email from server more than once, we record the email UIDL retrieved every time
        ' UIDLManager wraps the function to write/read uidl record from a text file.
        oUIDLManager = New UIDLManager()

        ' Load existed uidl records to UIDLManager
        Dim uidlfile As String = [String].Format("{0}\uidl.txt", localfolder)
        oUIDLManager.Load(uidlfile)

        ' Remove the local uidl which is not existed on the server.
        oUIDLManager.SyncUIDL(oCurServer, infos)
        oUIDLManager.Update()

        ' Remove the email file on local disk that is not existed on server
        _ClearLocalMails(infos, localfolder)

        Try
            Dim count As Integer = infos.Length
            For i As Integer = 0 To count - 1
                lblStatus.Text = [String].Format("Retrieve summary {0}/{1} ...", i + 1, count)
                Dim info As MailInfo = infos(i)

                ' For evaluation usage, please use "TryIt" as the license code, otherwise the 
                ' "Invalid license code" exception will be thrown. However, the object will expire in 1-2 months, then
                ' "Trial version expired" exception will be thrown.
                Dim oMail As New Mail("TryIt")

                Dim d As System.DateTime = System.DateTime.Now
                Dim cur As New System.Globalization.CultureInfo("en-US")
                Dim sdate As String = d.ToString("yyyyMMddHHmmss", cur)
                Dim fileName As String = [String].Format("{0}{1}{2}.eml", sdate, d.Millisecond.ToString("d3"), i)
                Dim localfile As String = [String].Format("{0}\{1}", localfolder, fileName)

                ' Detect if current email has been downloaded before.
                Dim uidl_item As UIDLItem = oUIDLManager.FindUIDL(oCurServer, info.UIDL)
                If Not (uidl_item Is Nothing) Then
                    localfile = [String].Format("{0}\{1}", localfolder, uidl_item.FileName)
                End If

                If File.Exists(localfile) Then
                    ' This mail has been downloaded from server.
                    oMail.Load(localfile, True)
                Else
                    ' Get the mail header from server.
                    oMail.Load(oClient.GetMailHeader(info))
                    oMail.SaveAs(localfile, True)
                End If

                If uidl_item Is Nothing Then
                    ' Add the email UIDL and local file name to uidl file to avoid we retrieve it again. 
                    oUIDLManager.AddUIDL(oCurServer, info.UIDL, fileName)
                End If

                Dim item As New ListViewItem(oMail.From.ToString())
                item.SubItems.Add(oMail.Subject)
                item.SubItems.Add(oMail.ReceivedDate.ToString("yyyy-MM-dd HH:mm:ss"))
                item.Tag = info

                If info.Deleted Then
                    item.Font = New System.Drawing.Font(item.Font, FontStyle.Strikeout)
                ElseIf Not info.Read Then
                    item.Font = New System.Drawing.Font(item.Font, FontStyle.Bold)
                End If

                lstMail.Items.Add(item)
            Next

            ' Update the uidl list to local uidl file and then we can load it next time.
            oUIDLManager.Update()
            lblStatus.Text = [String].Format("Total {0} email(s)", count)
        Catch ep As Exception
            ' Update the uidl list to local uidl file and then we can load it next time.
            oUIDLManager.Update()
            lblStatus.Text = ep.Message
            Throw ep
        End Try
    End Sub

#Region "Email Management"
    ' Pure deleted email from server, only IMAP4 supports this operation.
    Private Sub btnPure_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPure.Click
        If oCurServer.Protocol = ServerProtocol.ExchangeEWS OrElse oCurServer.Protocol = ServerProtocol.ExchangeWebDAV Then
            ' EWS and WebDAV doesn't support this operating
            Return
        End If
        Dim node As TreeNode = trAccounts.SelectedNode
        If node Is Nothing Then
            Return
        End If

        If node.Parent Is Nothing Then
            Return
        End If

        EnableIdle(False)
        Try
            ConnectServer(node)
            oClient.Expunge()
            LoadServerMails(node, node.Tag)
        Catch ep As Exception
            oClient.Close()
            MessageBox.Show(ep.Message)
        End Try

        EnableIdle(True)

    End Sub

    ' Delete email from server.
    Private Sub btnDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim node As TreeNode = trAccounts.SelectedNode
        If node Is Nothing Then
            Return
        End If

        If node.Parent Is Nothing Then
            Return
        End If

        Dim items As ListView.SelectedListViewItemCollection = lstMail.SelectedItems
        If items Is Nothing Then
            Return
        End If

        If items.Count = 0 Then
            Return
        End If

        Dim count As Integer = items.Count

        EnableIdle(False)
        Try

            Dim ar As New ArrayList()
            ConnectServer(node)
            For i As Integer = 0 To count - 1
                Dim info As MailInfo = items(i).Tag
                oClient.Delete(info)
                items(i).Font = New System.Drawing.Font(items(i).Font, FontStyle.Strikeout)

                If oCurServer.Protocol = ServerProtocol.ExchangeEWS Or oCurServer.Protocol = ServerProtocol.ExchangeWebDAV Then
                    ar.Add(items(i))
                End If
            Next

            count = ar.Count
            For i As Integer = 0 To count - 1
                lstMail.Items.Remove(ar(i))

            Next
        Catch ep As Exception
            oClient.Close()
            MessageBox.Show(ep.Message)
        End Try

        EnableIdle(True)
    End Sub

    ' Undelete  email marked as deleted from server, only IMAP4 supports this operation.
    Private Sub btnUndelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUndelete.Click
        If oCurServer.Protocol = ServerProtocol.ExchangeEWS OrElse oCurServer.Protocol = ServerProtocol.ExchangeWebDAV Then
            ' EWS and WebDAV doesn't support this operating
            Return
        End If

        Dim node As TreeNode = trAccounts.SelectedNode
        If node Is Nothing Then
            Return
        End If

        If node.Parent Is Nothing Then
            Return
        End If

        Dim items As ListView.SelectedListViewItemCollection = lstMail.SelectedItems
        If items Is Nothing Then
            Return
        End If

        If items.Count = 0 Then
            Return
        End If

        Dim count As Integer = items.Count

        EnableIdle(False)
        Try
            ConnectServer(node)
            For i As Integer = 0 To count - 1
                Dim item As ListViewItem = items(i)
                Dim info As MailInfo = item.Tag
                oClient.Undelete(info)
                If Not info.Read Then
                    item.Font = New System.Drawing.Font(item.Font, FontStyle.Bold)
                Else
                    item.Font = New System.Drawing.Font(item.Font, FontStyle.Regular)
                End If
            Next
        Catch ep As Exception
            oClient.Close()
            MessageBox.Show(ep.Message)
        End Try

        EnableIdle(True)

    End Sub

    ' Mark email as unread.
    Private Sub btnUnread_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUnread.Click
        Dim node As TreeNode = trAccounts.SelectedNode
        If node Is Nothing Then
            Return
        End If

        If node.Parent Is Nothing Then
            Return
        End If

        Dim items As ListView.SelectedListViewItemCollection = lstMail.SelectedItems
        If items Is Nothing Then
            Return
        End If

        If items.Count = 0 Then
            Return
        End If

        Dim count As Integer = items.Count

        EnableIdle(False)
        Try
            ConnectServer(node)
             For i As Integer = 0 To count - 1
                Dim item As ListViewItem = items(i)
                Dim info As MailInfo = item.Tag
                oClient.MarkAsRead(info, False)

                If info.Deleted Then
                    item.Font = New System.Drawing.Font(item.Font, FontStyle.Strikeout)
                Else
                    item.Font = New System.Drawing.Font(item.Font, FontStyle.Bold)
                End If
            Next
        Catch ep As Exception
            oClient.Close()
            MessageBox.Show(ep.Message)
        End Try

        EnableIdle(True)
    End Sub

    ' Move email to another folder.
    Private Sub btnMove_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnMove.Click
        Dim node As TreeNode = trAccounts.SelectedNode
        If node Is Nothing Then
            Return
        End If

        If node.Parent Is Nothing Then
            Return
        End If

        Dim items As ListView.SelectedListViewItemCollection = lstMail.SelectedItems
        If items Is Nothing Then
            Return
        End If

        If items.Count = 0 Then
            Return
        End If

        Dim count As Integer = items.Count

        EnableIdle(False)
        Try

            ConnectServer(node)
            Dim fds As Imap4Folder() = oClient.Imap4Folders

            Dim frm As New Form4()
            Dim fnode As TreeNode = frm.trFolders.Nodes.Add("Root Folder")
            ExpendFolders(fds, fnode)
            fnode.ExpandAll()

            If frm.ShowDialog() <> DialogResult.OK Then
                EnableIdle(True)
                Return
            End If

            Dim curfd As Imap4Folder = node.Tag
            Dim fd As Imap4Folder = frm.trFolders.SelectedNode.Tag
            If [String].Compare(curfd.FullPath, fd.FullPath, True) = 0 Then
                EnableIdle(True)
                Return
            End If

            Dim ar As New ArrayList()
            For i As Integer = 0 To count - 1
                Dim item As ListViewItem = items(i)
                Dim info As MailInfo = item.Tag
                oClient.Copy(info, fd)
                oClient.Delete(info)
                item.Font = New System.Drawing.Font(item.Font, FontStyle.Strikeout)

                If oCurServer.Protocol = ServerProtocol.ExchangeEWS OrElse oCurServer.Protocol = ServerProtocol.ExchangeWebDAV Then
                    ar.Add(items(i))
                End If
            Next

            count = ar.Count
            For i As Integer = 0 To count - 1
                lstMail.Items.Remove(ar(i))
            Next
        Catch ep As Exception
            oClient.Close()
            MessageBox.Show(ep.Message)
        End Try

        EnableIdle(True)
    End Sub

    ' Download and display current email.
    Private Sub lstMail_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstMail.Click

        Dim items As ListView.SelectedListViewItemCollection = lstMail.SelectedItems
        If items Is Nothing Then
            Return
        End If

        If items.Count = 0 Then
            Return
        End If

        Dim node As TreeNode = trAccounts.SelectedNode
        If node Is Nothing Then
            Return
        End If

        EnableIdle(True)
        If items.Count > 1 Then
            Return
        End If

        Dim item As ListViewItem = items(0)
        Dim info As MailInfo = item.Tag
        Dim fd As Imap4Folder = node.Tag

        ' try
        If True Then
            EnableIdle(False)
            ConnectServer(node)


            Dim mailbox As String = GetFolderByNode(node)
            _CreateFullFolder(mailbox)

            ' Find current email record in UIDL file.
            Dim oUIDL As UIDLItem = oUIDLManager.FindUIDL(oCurServer, info.UIDL)
            If oUIDL Is Nothing Then
                ' show never happen except you delete the file from the folder manually.
                Throw New Exception("No email file found!")
            End If

            ' Get the  local file name for this email UIDL
            Dim emlFile As String = [String].Format("{0}\{1}", mailbox, oUIDL.FileName)

            Dim pos As Integer = emlFile.LastIndexOf(".")
            Dim mainName As String = emlFile.Substring(0, pos)
            Dim htmlName As String = mainName & ".htm"

            ' only mail header is retrieved, now retrieve full content of mail.
            If Not File.Exists(htmlName) Then
                Dim oMail As Mail = oClient.GetMail(info)
                oMail.SaveAs(emlFile, True)
                pgBar.Minimum = 0
                pgBar.Maximum = 100
                pgBar.Value = 0
            End If

            If Not info.Read Then
                oClient.MarkAsRead(info, True)
                If info.Deleted Then
                    item.Font = New System.Drawing.Font(item.Font, FontStyle.Strikeout)
                Else
                    item.Font = New System.Drawing.Font(item.Font, FontStyle.Regular)
                End If
            End If

            ShowMail(emlFile)
        End If
        ' catch (Exception ep)
        ' {
        '   oClient.Close();
        '   MessageBox.Show(ep.Message);
        '}

        EnableIdle(True)
    End Sub

    ' Copy emails from one folder to another.
    Private Sub btnCopy_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnCopy.Click
        Dim node As TreeNode = trAccounts.SelectedNode
        If node Is Nothing Then
            Return
        End If

        If node.Parent Is Nothing Then
            Return
        End If

        Dim items As ListView.SelectedListViewItemCollection = lstMail.SelectedItems
        If items Is Nothing Then
            Return
        End If

        If items.Count = 0 Then
            Return
        End If

        Dim count As Integer = items.Count

        EnableIdle(False)
        Try
            ConnectServer(node)
            Dim fds As Imap4Folder() = oClient.Imap4Folders

            Dim frm As New Form4()
            Dim fnode As TreeNode = frm.trFolders.Nodes.Add("Root Folder")
            ExpendFolders(fds, fnode)
            fnode.ExpandAll()

            If frm.ShowDialog() <> DialogResult.OK Then
                EnableIdle(True)
                Return
            End If

            Dim curfd As Imap4Folder = node.Tag
            Dim fd As Imap4Folder = frm.trFolders.SelectedNode.Tag
            If [String].Compare(curfd.FullPath, fd.FullPath, True) = 0 Then
                EnableIdle(True)
                Return
            End If

            For i As Integer = 0 To count - 1
                Dim item As ListViewItem = items(i)
                Dim info As MailInfo = item.Tag
                oClient.Copy(info, fd)
            Next
        Catch ep As Exception
            oClient.Close()
            MessageBox.Show(ep.Message)
        End Try

        EnableIdle(True)
    End Sub

    ' Upload EML file to specified folder
    Private Sub btnUpload_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnUpload.Click
        Dim node As TreeNode = trAccounts.SelectedNode
        If node Is Nothing Then
            Return
        End If

        If node.Parent Is Nothing Then
            Return
        End If

        openFileDialog1.Multiselect = False
        If openFileDialog1.ShowDialog() = DialogResult.OK Then
            EnableIdle(False)
            Try
                Dim oMail As New Mail("TryIt")
                oMail.Load(openFileDialog1.FileName, False)
                ConnectServer(node)

                oClient.Append(node.Tag, oMail.Content)
                oClient.RefreshMailInfos()
                ShowNode()
            Catch ep As Exception
                oClient.Close()
                MessageBox.Show(ep.Message)
            End Try

            EnableIdle(True)
        End If

        pgBar.Minimum = 0
        pgBar.Maximum = 100
        pgBar.Value = 0
    End Sub
#End Region

    ' Show folder list and email list
    Private Sub ShowNode()
        Dim node As TreeNode = trAccounts.SelectedNode
        If node Is Nothing Then
            lstMail.Items.Clear()
            Return
        End If

        Try
            If node.Parent Is Nothing Then
                ' Current node is root node, 
                ' So we get all folders from server and display it
                lstMail.Items.Clear()
                ConnectServer(node)

                lblStatus.Text = "Refreshing Folders ..."
                Dim fds As Imap4Folder() = oClient.Imap4Folders
                ExpendFolders(fds, node)

                node.ExpandAll()
                lblStatus.Text = ""
            Else
                Dim fd As Imap4Folder = node.Tag
                If Not fd.NoSelect Then
                    ' Display emails list in current folder
                    LoadServerMails(node, fd)
                Else
                    lblStatus.Text = ""
                    ' This is a folder without email storage
                    lstMail.Items.Clear()
                End If
            End If
        Catch ep As Exception
            oClient.Close()
            MessageBox.Show(ep.Message)
        End Try

    End Sub

    Private Sub trAccounts_AfterSelect(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles trAccounts.AfterSelect
        EnableIdle(False)
        ShowNode()
        EnableIdle(True)
    End Sub


    ' Enable/Disable control
    Private Sub EnableIdle(ByVal bIdle As Boolean)
        btnDelete.Enabled = bIdle
        btnUndelete.Enabled = bIdle
        btnUnread.Enabled = bIdle
        btnPure.Enabled = bIdle
        btnMove.Enabled = bIdle
        btnCopy.Enabled = bIdle
        btnUpload.Enabled = bIdle

        If lstMail.SelectedItems.Count = 0 Then
            btnDelete.Enabled = False
            btnUndelete.Enabled = False
            btnUnread.Enabled = False
            btnMove.Enabled = False
            btnCopy.Enabled = False
        End If

        If trAccounts.SelectedNode Is Nothing Then
            btnUpload.Enabled = False
        Else
            If trAccounts.SelectedNode.Parent Is Nothing Then
                btnUpload.Enabled = False

            End If
        End If

        btnCancel.Enabled = Not bIdle
        If btnStart.Enabled Then
            btnCancel.Enabled = False
        End If

        btnQuit.Enabled = bIdle

        trAccounts.Enabled = bIdle
        lstMail.Enabled = bIdle

        If Not (oCurServer Is Nothing) Then
            If oCurServer.Protocol = ServerProtocol.ExchangeWebDAV OrElse oCurServer.Protocol = ServerProtocol.ExchangeEWS Then
                ' Exchange WebDAV and EWS doesn't support this operating
                btnUndelete.Enabled = False
                btnPure.Enabled = False
            End If
        End If
    End Sub

    ' Resize control automatically
    Private Sub Form1_Resize(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Resize
        If Me.Width < 900 Then
            Me.Width = 900
        End If

        If Me.Height < 470 Then
            Me.Height = 470
        End If

        lstMail.Width = Me.Width - 455
        webMail.Width = lstMail.Width
        pgBar.Width = webMail.Width
        trAccounts.Height = Me.Height - 60
        webMail.Height = Me.Height - lstMail.Height - 120
        pgBar.Top = Me.Height - 60
    End Sub

    Private Sub lstProtocol_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles lstProtocol.SelectedIndexChanged
        ' By default, Exchange Web Service requires SSL connection.
        ' For other protocol, please set SSL connection based on your server setting manually

        If (lstProtocol.SelectedIndex + 1) = CInt(ServerProtocol.ExchangeEWS) Then
            chkSSL.Checked = True
        End If
    End Sub

  
    
End Class
