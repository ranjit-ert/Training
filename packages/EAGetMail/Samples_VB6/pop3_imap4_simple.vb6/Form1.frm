VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   11700
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser webMail 
      Height          =   2895
      Left            =   3720
      TabIndex        =   20
      Top             =   3360
      Width           =   7935
      ExtentX         =   13996
      ExtentY         =   5106
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.CommandButton btnDel 
      Caption         =   "Delete"
      Height          =   375
      Left            =   10320
      TabIndex        =   19
      Top             =   2880
      Width           =   1335
   End
   Begin MSComctlLib.ListView lstMail 
      Height          =   2535
      Left            =   3720
      TabIndex        =   18
      Top             =   240
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   4471
      SortKey         =   3
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "From"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Subject"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Date"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "SortValue"
         Object.Width           =   2
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Account Information"
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin MSComctlLib.ProgressBar pgBar 
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   4680
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton btnCancel 
         Caption         =   "Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   14
         Top             =   3480
         Width           =   1455
      End
      Begin VB.CommandButton btnStart 
         Caption         =   "Start"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   3480
         Width           =   1455
      End
      Begin VB.CheckBox chkLeaveCopy 
         Caption         =   "Leave a copy of message on server"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   3000
         Width           =   3015
      End
      Begin VB.ComboBox lstProtocol 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2580
         Width           =   2175
      End
      Begin VB.ComboBox lstAuthType 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2100
         Width           =   2175
      End
      Begin VB.CheckBox chkSSL 
         Caption         =   "SSL Connection"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox textPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1080
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1275
         Width           =   2175
      End
      Begin VB.TextBox textUser 
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Top             =   795
         Width           =   2175
      End
      Begin VB.TextBox textServer 
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Top             =   315
         Width           =   2175
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "Ready"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   4200
         Width           =   465
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Protocol"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   2640
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Auth Type"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "User"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Server"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   465
      End
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      Caption         =   "Total"
      Height          =   195
      Left            =   3720
      TabIndex        =   21
      Top             =   2940
      Width           =   360
   End
   Begin VB.Label Label6 
      Caption         =   "Warning: if ""leave a copy of message on server"" is not checked,  the emails on the server will be deleted !"
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   120
      TabIndex        =   17
      Top             =   5400
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  ===============================================================================
' |    THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF      |
' |    ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO    |
' |    THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A         |
' |    PARTICULAR PURPOSE.                                                    |
' |    Copyright (c)2010  ADMINSYSTEM SOFTWARE LIMITED                         |
' |
' |    Project: It demonstrates how to use EAGetMailObj to receive/parse email.
' |
' |
' |    File: Form1 : implementation file
' |
' |    Author: Ivan Lui ( ivan@emailarchitect.net )
'  ===============================================================================
Const MailServerConnectType_ConnectSSLAuto = 0
Const MailServerConnectType_ConnectSSL = 1
Const MailServerConnectType_ConnectTLS = 2

Const ProxyType_Socks4 = 0
Const ProxyType_Socks5 = 1
Const ProxyType_Http = 2

Const MailServerAuthLogin = 0
Const MailServerAuthCRAM5 = 1
Const MailServerAuthNTLM = 2

Const MailServerPop3 = 0
Const MailServerImap4 = 1
Const MailServerEWS = 2 'Exchange Web Service, Exchange 2007/2010
Const MailServerDAV = 3 'Exchange WebDAV, Exchange 2000/2003

Private WithEvents oClient As EAGetMailObjLib.MailClient
Attribute oClient.VB_VarHelpID = -1
Private oTools As EAGetMailObjLib.Tools

Private m_bCancel As Boolean
Private m_curpath As String
Private m_uidlfile As String
Private m_uidls As String

Private Sub Form_Load()

        webMail.Navigate "about:blank"
        Set oTools = New EAGetMailObjLib.Tools
        m_curpath = App.Path
        m_uidlfile = m_curpath
        m_uidlfile = m_uidlfile & "\uidl.txt"

        lstProtocol.AddItem "POP3"
        lstProtocol.AddItem "IMAP4"
        lstProtocol.AddItem "Exchange Web Service - Exchange 2007/2010"
        lstProtocol.AddItem "Exchange WebDAV - Exchange 2000/2003"
        lstProtocol.ListIndex = 0

        lstAuthType.AddItem "USER/LOGIN"
        lstAuthType.AddItem "APOP(CRAM-MD5)"
        lstAuthType.AddItem "NTLM"
        lstAuthType.ListIndex = 0
        m_bCancel = False
        lblStatus.Caption = ""
        pgBar.Min = 0
        pgBar.Max = 100
        pgBar.Value = 0
        LoadMails
        
        If lstMail.ListItems.Count > 0 Then
            lstMail.ListItems.Item(1).Selected = True
        End If
End Sub

Private Sub btnStart_Click()
        Dim Server, User, Password As String
        Server = Trim(textServer.Text)
        User = Trim(textUser.Text)
        Password = Trim(textPassword.Text)
        
        If Len(Server) = 0 Or Len(User) = 0 Or Len(Password) = 0 Then
            MsgBox "Please input server address, user and password!"
            Exit Sub
        End If
        
        On Error GoTo ErrorHandle
        
        Set oClient = New EAGetMailObjLib.MailClient
        'For evaluation usage, please use "TryIt" as the license code, otherwise the
        '"invalid license code" exception will be thrown. However, the object will expire in 1-2 months, then
        '"trial version expired" exception will be thrown.
        oClient.LicenseCode = "TryIt"
        
        Dim oServer As New EAGetMailObjLib.MailServer
        
        'oClient.LogFileName = "d:\pop3.txt" 'generate a log file
        
        m_bCancel = False
        oServer.Server = Server
        oServer.User = User
        oServer.Password = Password
        oServer.SSLConnection = chkSSL.Value
        oServer.AuthType = lstAuthType.ListIndex
        oServer.Protocol = lstProtocol.ListIndex
        
        If lstProtocol.ListIndex = MailServerImap4 Then
            If oServer.SSLConnection Then
                oServer.Port = 993 'SSL IMAP4
            Else
                oServer.Port = 143 'IMAP4 normal
            End If
        Else
            If oServer.SSLConnection Then
                oServer.Port = 995 'SSL POP3
            Else
                oServer.Port = 110 'POP3 normal
            End If
        End If
       
        btnStart.Enabled = False
        btnCancel.Enabled = True
        
        lblStatus.Caption = "Connecting server ..."
        oClient.Connect oServer
        
        ' uidl is the identifier of every email on POP3/IMAP4 server, to avoid retrieve
        ' the same email from server more than once, we record the email uidl retrieved every time
        ' if you delete the email from server every time and not to leave a copy of email on
        ' the server, then please remove all the function about uidl.
        Dim oUIDLManager As New UIDLManager
        oUIDLManager.Load m_uidlfile
        
        Dim arInfo
        arInfo = oClient.GetMailInfos()
        
        ' Remove the local uidl which is not existed on the server.
        oUIDLManager.SyncUIDL oServer, arInfo
        oUIDLManager.Update
        
        Dim mailFolder As String
        mailFolder = m_curpath & "\inbox"
        oTools.CreateFolder mailFolder
        Dim i, Count
        Count = UBound(arInfo)
        For i = LBound(arInfo) To Count
            Dim oInfo As EAGetMailObjLib.MailInfo
            Set oInfo = arInfo(i)
            lblStatus.Caption = "Retrieving " & i + 1 & "/" & Count + 1 & "..."
            
            Dim oUIDLItem As UIDLItem
            Set oUIDLItem = oUIDLManager.FindUIDL(oServer, oInfo.UIDL)
            
            'if this email has not retrieved before, then get it
            If oUIDLItem Is Nothing Then
                Dim oMail As EAGetMailObjLib.Mail
                Set oMail = oClient.GetMail(oInfo)
                
                Dim fileName As String
                'generate a random file name by current local datetime,
                'you can use your method to generate the filename if you do not like it
                Dim shortName As String
                shortName = oTools.GenFileName(i) & ".eml"
                fileName = mailFolder & "\" & shortName
               
                oMail.SaveAs fileName, True
                
                Dim oItem As ListItem
                Set oItem = lstMail.ListItems.Add()
                oItem.Bold = True
                oItem.Text = oMail.From.Name & "<" & oMail.From.Address & ">"
                oItem.Tag = fileName
                oItem.ListSubItems.Add , , oMail.Subject
                oItem.ListSubItems.Add , , oMail.ReceivedDate
                oItem.ListSubItems.Add , , CDbl(oMail.ReceivedDate)
                
                lblTotal.Caption = "Total " & lstMail.ListItems.Count & " email(s)"
                If chkLeaveCopy.Value = Checked Then
                    'add the email uidl to uidl file to avoid we retrieve it next time.
                    oUIDLManager.AddUIDL oServer, oInfo.UIDL, shortName
                End If
            End If
        Next
        
        If chkLeaveCopy.Value = Unchecked Then
            For i = LBound(arInfo) To Count
                lblStatus.Caption = "Deleting " & i & " ..."
                oClient.Delete arInfo(i)
                ' Remove UIDL from local uidl file.
                oUIDLManager.RemoveUIDL oServer, arInfo(i).UIDL
            Next
        End If
        
        ' Delete method just mark the email as deleted,
        ' Quit method pure the emails from server exactly.
        oClient.Quit
        Set oClient = Nothing
        lblStatus = "Completed"
        
        'update the uidl list to a text file and then we can load it next time.
        oUIDLManager.Update
        btnStart.Enabled = True
        btnCancel.Enabled = False

        Exit Sub
ErrorHandle:
        MsgBox Err.Description
        oUIDLManager.Update
        
        Set oClient = Nothing
        btnStart.Enabled = True
        btnCancel.Enabled = False
        lblStatus = Err.Description
End Sub

'cancel email retrieving
Private Sub btnCancel_Click()
    m_bCancel = True
    btnCancel.Enabled = False
End Sub

'delete selected email from local disk
Private Sub btnDel_Click()
    Dim oItem As ListItem
    Set oItem = lstMail.SelectedItem
    If oItem Is Nothing Then
        Exit Sub
    End If
    
    If MsgBox("Are you sure to delete selected emails?", vbYesNo, "Delete Email") <> vbYes Then
        Exit Sub
    End If
    
    Dim fileName As String
    fileName = oItem.Tag
    Dim pos
    pos = InStrRev(fileName, ".")
    Dim mainName
    Dim htmlName
    mainName = Mid(fileName, 1, pos - 1)
    htmlName = mainName & ".htm"

    Dim tempFolder As String
    tempFolder = mainName
    Dim nIndex
    nIndex = oItem.Index
    
    lstMail.ListItems.Remove (oItem.Index)
    
    On Error GoTo ErrorDel
    
    oTools.RemoveFile fileName
    oTools.RemoveFile htmlName
    oTools.RemoveFolder tempFolder, True
 
    
    lblTotal.Caption = "Total " & lstMail.ListItems.Count & " email(s)"

    
    If lstMail.ListItems.Count > 0 Then
        If nIndex > lstMail.ListItems.Count Then
            nIndex = lstMail.ListItems.Count
        End If
        
        lstMail.ListItems.Item(nIndex).Selected = True
        ShowMail lstMail.ListItems.Item(nIndex).Tag
    Else
        webMail.Navigate "about:blank"
    End If
    Exit Sub
ErrorDel:
    MsgBox Err.Description
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    m_bCancel = True
    End
End Sub

Private Sub Form_Resize()
On Error Resume Next
    If Me.Width < 12000 Then
        Me.Width = 12000
    End If
    
    If Me.Height < 6900 Then
        Me.Height = 6900
    End If
    
    lstMail.Width = Me.Width - 3900
    webMail.Width = lstMail.Width
    btnDel.Left = Me.Width - 1600
    webMail.Height = Me.Height - (lstMail.Height + btnDel.Height + 1200)
End Sub

Private Sub lstMail_ItemClick(ByVal Item As MSComctlLib.ListItem)
    ShowMail Item.Tag
    Item.Bold = False
End Sub

'=========================================================================================================
'Region "Parse and Display Mails"
'=========================================================================================================
Private Sub LoadMails()
        lstMail.ListItems.Clear
        Dim mailFolder As String
        mailFolder = m_curpath & "\inbox"
        oTools.CreateFolder mailFolder
        
        Dim arFile
        arFile = oTools.GetFiles(mailFolder & "\*.eml")
        
        Dim i, n
        For i = LBound(arFile) To UBound(arFile)
            Dim fullname As String
            fullname = arFile(i)
            
            Dim oMail As New EAGetMailObjLib.Mail
            'For evaluation usage, please use "TryIt" as the license code, otherwise the
            '"invalid license code" exception will be thrown. However, the object will expire in 1-2 months, then
            '"trial version expired" exception will be thrown.
            oMail.LicenseCode = "TryIt"
            oMail.LoadFile fullname, True
            
            Dim oItem As ListItem
            Set oItem = lstMail.ListItems.Add()
            oItem.Text = oMail.From.Name & "<" & oMail.From.Address & ">"
            oItem.Tag = fullname
            oItem.ListSubItems.Add , , oMail.Subject
            oItem.ListSubItems.Add , , oMail.ReceivedDate
            oItem.ListSubItems.Add , , CDbl(oMail.ReceivedDate)
            n = n + 1
            

            Dim pos
            pos = InStrRev(fullname, ".")
            Dim mainName
            Dim htmlName
            mainName = Mid(fullname, 1, pos - 1)
            htmlName = mainName & ".htm"

            If Not oTools.ExistFile(htmlName) Then
                ' this email is unread, we set the font style to bold.
                oItem.Bold = True
            Else
                oItem.Bold = False
            End If
            oMail.Clear
        Next
        
        lblTotal.Caption = "Total " & n & " email(s)"
End Sub

Private Sub ShowMail(ByVal fileName As String)
    Dim pos
    pos = InStrRev(fileName, ".")
    Dim mainName
    Dim htmlName
    mainName = Mid(fileName, 1, pos - 1)
    htmlName = mainName & ".htm"

    Dim tempFolder As String
    tempFolder = mainName
    If Not (oTools.ExistFile(htmlName)) Then
        'we haven't generate the html for this email, generate it now.
        GenerateHtmlForEmail htmlName, fileName, tempFolder
    End If
    webMail.Navigate htmlName
End Sub

' we generate a html + attachment folder for every email, once the html is create,
' next time we don't need to parse the email again.
Private Sub GenerateHtmlForEmail(ByVal htmlName As String, ByVal emlFile As String, ByVal tempFolder As String)
    
    On Error GoTo ErrorGenHtml

    Dim oMail As New EAGetMailObjLib.Mail
    'For evaluation usage, please use "TryIt" as the license code, otherwise the
    '"invalid license code" exception will be thrown. However, the object will expire in 1-2 months, then
    '"trial version expired" exception will be thrown.
    oMail.LicenseCode = "TryIt"
        
    oMail.LoadFile emlFile, False
    If Err.Number <> 0 Then
        MsgBox Err.Description
        Exit Sub
    End If
    
    On Error Resume Next
    If oMail.IsEncrypted Then
        Set oMail = oMail.Decrypt(Nothing)
        If Err.Number <> 0 Then
            MsgBox Err.Description
        End If
    End If
    
    If oMail.IsSigned Then
        oMail.VerifySignature
        If Err.Number <> 0 Then
            MsgBox Err.Description
        End If
    End If
    
    On Error GoTo ErrorGenHtml
    
    Dim html As String
    html = oMail.HtmlBody
    Dim hdr As String
    hdr = hdr & "<font face=""Courier New,Arial"" size=2>"
    hdr = hdr & "<b>From:</b> " + FormatHtmlTag(oMail.From.Name & "<" & oMail.From.Address & ">") + "<br>"
    
    
    Dim addrs
    addrs = oMail.To
    Dim i, Count
    
    i = LBound(addrs)
    Count = UBound(addrs)
    If (Count >= i) Then
        hdr = hdr & "<b>To:</b> "
        For i = LBound(addrs) To Count
            hdr = hdr & FormatHtmlTag(addrs(i).Name & "<" & addrs(i).Address & ">")
            If (i < Count) Then
                hdr = hdr & ";"
            End If
        Next
        hdr = hdr & "<br>"
    End If
    
    addrs = oMail.Cc
    i = LBound(addrs)
    Count = UBound(addrs)
    If (Count >= i) Then
        hdr = hdr & "<b>Cc:</b> "
        For i = LBound(addrs) To Count
            hdr = hdr & FormatHtmlTag(addrs(i).Name & "<" & addrs(i).Address & ">")
            If (i < Count) Then
                hdr = hdr & ";"
            End If
        Next
        hdr = hdr & "<br>"
    End If
        
    hdr = hdr & "<b>Subject:</b>" & FormatHtmlTag(oMail.Subject) & "<br>" & vbCrLf
    Dim atts
    atts = oMail.Attachments
    i = LBound(atts)
    Count = UBound(atts)
    If (Count >= i) Then
        If Not oTools.ExistFile(tempFolder) Then
            oTools.CreateFolder (tempFolder)
        End If
        
        hdr = hdr & "<b>Attachments:</b>"
        For i = LBound(atts) To Count
            Dim att As Attachment
            Set att = atts(i)
            If LCase(att.Name) = "winmail.dat" Then
                Dim tatts
                tatts = oMail.ParseTNEF(att.Content, True)
                For x = LBound(tatts) To UBound(tatts)
                        Dim tatt As Attachment
                        Set tatt = tatts(x)
                        Dim tattname As String
                        tattname = tempFolder & "\" & tatt.Name
                        tatt.SaveAs tattname, True
                        hdr = hdr & "<a href=""" & tattname & """ target=""_blank"">" & tatt.Name & "</a> "
                Next
                
            Else
                Dim attname
                attname = tempFolder & "\" & att.Name
                att.SaveAs attname, True
                hdr = hdr & "<a href=""" & attname & """ target=""_blank"">" & att.Name & "</a> "
                If Len(att.ContentID) > 0 Then
                    'show embedded image.
                    html = Replace(html, "cid:" + att.ContentID, attname)
                ElseIf InStr(1, att.ContentType, "image/", vbTextCompare) = 1 Then
                    'show attached image.
                    html = html & "<hr><img src=""" & attname & """>"
                End If
            End If
        Next
    End If
    
    Dim reg
    Set reg = CreateObject("VBScript.RegExp")
    reg.Global = True
    reg.IgnoreCase = True
    reg.Pattern = "(<meta[^>]*charset[ \t]*=[ \t""]*)([^<> \r\n""]*)"
    html = reg.Replace(html, "$1utf-8")
    
    hdr = "<meta HTTP-EQUIV=""Content-Type"" Content=""text/html; charset=utf-8"">" & hdr
    html = hdr & "<hr>" & html
    oTools.WriteTextFile htmlName, html, 65001
    oMail.Clear
    Exit Sub
    
ErrorGenHtml:
    MsgBox "Failed to generate html file for the email; " & Err.Description
    
End Sub
Private Function FormatHtmlTag(ByVal src As String) As String
    src = Replace(src, ">", "&gt;")
    src = Replace(src, "<", "&lt;")
    FormatHtmlTag = src
End Function
'=============================================================================================
' End of Parse and Display Mails
'=============================================================================================



'================================================================================================
'"EAGetMail Event Handler"
' you don't have to use EAGetMail Event, but using Event make your application more user friendly
'================================================================================================
Private Sub oClient_OnAuthorized(ByVal oSender As Object, Cancel As Boolean)
        lblStatus.Caption = "Authorized"
        Cancel = m_bCancel
End Sub

Private Sub oClient_OnConnected(ByVal oSender As Object, Cancel As Boolean)
        lblStatus.Caption = "Connected"
        Cancel = m_bCancel
End Sub

Private Sub oClient_OnIdle(ByVal oSender As Object, Cancel As Boolean)
        DoEvents
        Cancel = m_bCancel
End Sub

Private Sub oClient_OnQuit(ByVal oSender As Object, Cancel As Boolean)
        lblStatus.Caption = "Quiting ... "
End Sub

Private Sub oClient_OnReceivingDataStream(ByVal oSender As Object, _
ByVal oInfo As Object, ByVal Received As Long, ByVal Total As Long, Cancel As Boolean)
        pgBar.Min = 0
        pgBar.Max = 100
        pgBar.Value = (Received / Total) * 100
        DoEvents
        Cancel = m_bCancel
End Sub

Private Sub oClient_OnSecuring(ByVal oSender As Object, Cancel As Boolean)
        lblStatus.Caption = "Securing ..."
End Sub
'================================================================================================
'"EAGetMail Event Handler" end
'================================================================================================


' By default, Exchange Web Service requires SSL connection.
'For other protocol, please set SSL connection based on your server setting manually
Private Sub lstProtocol_Click()
    If lstProtocol.ListIndex = MailServerEWS Then
        MsgBox "By default, Exchange Web Service requires SSL!"
        chkSSL.Value = Checked 'By default, Exchange Web Service requires SSL
    End If
End Sub
