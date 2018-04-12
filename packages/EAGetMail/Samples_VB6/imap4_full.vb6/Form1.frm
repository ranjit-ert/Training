VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7980
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   13560
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   13560
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog FileDlg 
      Left            =   2400
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "*.eml"
      Filter          =   "Email file (*.EML)|*.EML"
   End
   Begin VB.CommandButton btnUpload 
      Caption         =   "Upload"
      Height          =   375
      Left            =   12480
      TabIndex        =   17
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton btnMove 
      Caption         =   "Move"
      Height          =   375
      Left            =   11400
      TabIndex        =   16
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton btnCopy 
      Caption         =   "Copy"
      Height          =   375
      Left            =   10320
      TabIndex        =   15
      Top             =   3000
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Account Information"
      Height          =   5895
      Left            =   0
      TabIndex        =   21
      Top             =   120
      Width           =   3495
      Begin VB.CommandButton btnQuit 
         Caption         =   "Quit"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   3840
         Width           =   3135
      End
      Begin VB.TextBox textServer 
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Top             =   315
         Width           =   2175
      End
      Begin VB.TextBox textUser 
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Top             =   795
         Width           =   2175
      End
      Begin VB.TextBox textPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1080
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1275
         Width           =   2175
      End
      Begin VB.CheckBox chkSSL 
         Caption         =   "SSL Connection"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   2895
      End
      Begin VB.ComboBox lstAuthType 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2100
         Width           =   2175
      End
      Begin VB.ComboBox lstProtocol 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2580
         Width           =   2175
      End
      Begin VB.CommandButton btnStart 
         Caption         =   "Start"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   3240
         Width           =   3135
      End
      Begin VB.CommandButton btnCancel 
         Caption         =   "Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   4440
         Width           =   3135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Server"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "User"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   840
         Width           =   330
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Auth Type"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Protocol"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   2640
         Width           =   585
      End
   End
   Begin MSComctlLib.ProgressBar pgBar 
      Height          =   255
      Left            =   6240
      TabIndex        =   20
      Top             =   7200
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton btnPure 
      Caption         =   "Pure"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9360
      TabIndex        =   14
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton btnUnread 
      Caption         =   "Unread"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8400
      TabIndex        =   13
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton btnUndelete 
      Caption         =   "Undelete"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7320
      TabIndex        =   12
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton btnDelete 
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6240
      TabIndex        =   11
      Top             =   3000
      Width           =   975
   End
   Begin MSComctlLib.StatusBar statusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   19
      Top             =   7605
      Width           =   13560
      _ExtentX        =   23918
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser webMail 
      Height          =   3495
      Left            =   6240
      TabIndex        =   18
      Top             =   3600
      Width           =   7215
      ExtentX         =   12726
      ExtentY         =   6165
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
      Location        =   "http:///"
   End
   Begin MSComctlLib.ListView lstMail 
      Height          =   2655
      Left            =   6240
      TabIndex        =   10
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   4683
      SortKey         =   3
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
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
   Begin MSComctlLib.TreeView lstAccount 
      Height          =   7335
      Left            =   3600
      TabIndex        =   9
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   12938
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Menu mnuFolders 
      Caption         =   "Folders"
      Visible         =   0   'False
      Begin VB.Menu mnuRefreshFolder 
         Caption         =   "Refresh Folder"
      End
      Begin VB.Menu menuRefreshMails 
         Caption         =   "Refresh Mails"
      End
      Begin VB.Menu mnuAddFolder 
         Caption         =   "AddFolder"
      End
      Begin VB.Menu mnuDeleteFolder 
         Caption         =   "Delete Folder"
      End
      Begin VB.Menu mnuRenameFolder 
         Caption         =   "Rename Folder"
      End
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
' |    Copyright (c)2010-2012  ADMINSYSTEM SOFTWARE LIMITED                         |
' |
' |    Project: It demonstrates how to use EAGetMailObj to  access IMAP4/Exchange Server folders/emails.
' |
' |
' |    File: Form1 : implementation file
' |
' |    Author: Ivan Lui ( ivan@emailarchitect.net )
'  ==============================================================================='

Option Explicit

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
Const MailServerEWS = 2 'Exchange Web Service - Exchange 2007/2010
Const MailServerDAV = 3 'Exchange WebDAV - Exchange 2000/2003

Private WithEvents oClient As EAGetMailObjLib.MailClient
Attribute oClient.VB_VarHelpID = -1
Private oTools As EAGetMailObjLib.Tools
Private oCurServer As EAGetMailObjLib.MailServer
Private oUIDLManager As EAGetMailObjLib.UIDLManager

Private m_bCancel As Boolean
Private m_curpath As String
Private m_uidlfile As String
Private m_uidls As String

Private Sub EnableButton(IsEnabled As Boolean)
    btnDelete.Enabled = IsEnabled
    btnUndelete.Enabled = IsEnabled
    btnUnread.Enabled = IsEnabled
    btnPure.Enabled = IsEnabled
    btnMove.Enabled = IsEnabled
    btnCopy.Enabled = IsEnabled
    btnUpload.Enabled = IsEnabled
    
    If SelectedCounts() = 0 Then
        btnDelete.Enabled = False
        btnUndelete.Enabled = False
        btnUnread.Enabled = False
        btnCopy.Enabled = False
        btnMove.Enabled = False
    End If

    Dim oNode
    Set oNode = lstAccount.SelectedItem
    If oNode Is Nothing Then
        btnUpload.Enabled = False
    Else
        If oNode.Parent Is Nothing Then
            btnUpload.Enabled = False
        End If
    End If
    
    btnCancel.Enabled = (Not IsEnabled)

    If btnStart.Enabled Then
        btnCancel.Enabled = False
    End If

    btnQuit.Enabled = IsEnabled
    lstAccount.Enabled = IsEnabled
    lstMail.Enabled = IsEnabled

    

    If Not (oCurServer Is Nothing) Then
    
        If (oCurServer.Protocol = MailServerEWS Or _
            oCurServer.Protocol = MailServerDAV) Then
        ' Exchange WebDAV and EWS doesn't support this operating
            btnUndelete.Enabled = False
            btnPure.Enabled = False
        End If
    End If
End Sub

Private Sub btnCancel_Click()
    m_bCancel = True
    btnCancel.Enabled = False
End Sub

Private Sub ConnectServer()
    EnableButton False
    
    If oClient.Connected Then
        Exit Sub
    End If

    statusBar.SimpleText = "Connecting server ..."
    m_bCancel = False
    oClient.Connect oCurServer
    statusBar.SimpleText = "Online"
    
    Dim oNode As Node
    Set oNode = lstAccount.SelectedItem
   
    If oNode Is Nothing Then
        Exit Sub
    End If
    
    If oNode.Parent Is Nothing Then
        Exit Sub
    End If
    ' select current folder
    oClient.SelectFolder oNode.Tag
   
End Sub

' Copy email to another folder
Private Sub btnCopy_Click()

    Dim oNode As Node
 
    Set oNode = lstAccount.SelectedItem
    If oNode Is Nothing Then
        Exit Sub
    End If
    
    If oNode.Parent Is Nothing Then
        Exit Sub
    End If
    
    If SelectedCounts() < 1 Then
        Exit Sub
    End If
    
    On Error GoTo ErrorHandle
    
    EnableButton False
   
    Dim oFolder As EAGetMailObjLib.Imap4Folder
    Set oFolder = oNode.Tag
    
    ConnectServer 'connect server on demand
   
    
    Dialog1.lstFolder.Nodes.Clear
    Set oNode = Dialog1.lstFolder.Nodes.Add(, , "Root")
    oNode.Text = "Root Folder"
    
    Dialog1.lstFolder.SelectedItem = oNode
    ConnectServer 'connect server on demand
    
    ExpendFoldersEx oClient.Imap4Folders, oNode.Key
    oNode.Expanded = True
    Dialog1.Show 1, Me
    
    Dim oDest As Node
    Set oDest = Dialog1.lstFolder.SelectedItem
    If oDest Is Nothing Then
        EnableButton True
        Exit Sub
    End If
    
    
    Dim i As Integer
    i = 1
    Do While (i <= lstMail.ListItems.Count)
        Dim Item As ListItem
        Set Item = lstMail.ListItems(i)
        If Item.Selected Then
            Dim oInfo As EAGetMailObjLib.MailInfo
            Set oInfo = Item.Tag
            oClient.Copy oInfo, oDest.Tag
        End If
        i = i + 1
    Loop
        
    EnableButton True
    Exit Sub
ErrorHandle:
    oClient.Close
    MsgBox Err.Description
    EnableButton True
End Sub


' Move email to another folder
Private Sub btnMove_Click()
    Dim oNode As Node
 
    Set oNode = lstAccount.SelectedItem
    If oNode Is Nothing Then
        Exit Sub
    End If
    
    If oNode.Parent Is Nothing Then
        Exit Sub
    End If
    
    If SelectedCounts() < 1 Then
        Exit Sub
    End If
    
    On Error GoTo ErrorHandle
    
    EnableButton False
    Dim oFolder As EAGetMailObjLib.Imap4Folder
    Set oFolder = oNode.Tag
    
    ConnectServer 'connect server on demand
  
    Dialog1.lstFolder.Nodes.Clear
    Set oNode = Dialog1.lstFolder.Nodes.Add(, , "Root")
    oNode.Text = "Root Folder"
    
    Dialog1.lstFolder.SelectedItem = oNode
    ConnectServer 'connect server on demand
    
    ExpendFoldersEx oClient.Imap4Folders, oNode.Key
    oNode.Expanded = True
    Dialog1.Show 1, Me
    
    Dim oDest As Node
    Set oDest = Dialog1.lstFolder.SelectedItem
    If oDest Is Nothing Then
        EnableButton True
        Exit Sub
    End If
    
    
     Dim i As Integer
    i = 1
    Do While (i <= lstMail.ListItems.Count)
        Dim Item As ListItem
        Set Item = lstMail.ListItems(i)
        If Item.Selected Then
            Dim oInfo As EAGetMailObjLib.MailInfo
            Set oInfo = Item.Tag
            oClient.Move oInfo, oDest.Tag
            Item.Bold = False
            Item.ForeColor = &H80000010
            
            If oCurServer.Protocol <> MailServerImap4 Then
                lstMail.ListItems.Remove i
            Else
                i = i + 1
            End If
        Else
            i = i + 1
        End If
    Loop
        
    EnableButton True
    Exit Sub
ErrorHandle:
    oClient.Close  'error, reconnect server next time
    MsgBox Err.Description
    EnableButton True
End Sub

' Pure deleted emails from server, only IMAP4 support it.
Private Sub btnPure_Click()
    Dim oNode As Node
 
    Set oNode = lstAccount.SelectedItem
    If oNode Is Nothing Then
        Exit Sub
    End If
    
    If oNode.Parent Is Nothing Then
        Exit Sub
    End If
    
On Error GoTo ErrorHandle
    EnableButton False
    
    Dim oFolder As EAGetMailObjLib.Imap4Folder
    Set oFolder = oNode.Tag
    
    ConnectServer 'connect server on demand
    oClient.SelectFolder oFolder
    oClient.Expunge
    
    LoadServerMails oNode, oFolder
    EnableButton True
    Exit Sub
ErrorHandle:
    oClient.Close 'error, reconnect server next time
    MsgBox Err.Description
    EnableButton True
End Sub

Private Sub btnQuit_Click()
    On Error GoTo ErrorHandle
    
    oClient.Logout
    webMail.Navigate ("about:blank")
ErrorHandle:
    lstMail.ListItems.Clear
    lstAccount.Nodes.Clear
    btnStart.Enabled = True
    oClient.Close
    EnableButton False
End Sub

' Mark email as unread
Private Sub btnUnread_Click()
    Dim oNode As Node
 
    Set oNode = lstAccount.SelectedItem
    If oNode Is Nothing Then
        Exit Sub
    End If
    
    If oNode.Parent Is Nothing Then
        Exit Sub
    End If

    Dim oFolder As EAGetMailObjLib.Imap4Folder
    Set oFolder = oNode.Tag
    
 
    If SelectedCounts() < 1 Then
        Exit Sub
    End If

    
On Error GoTo ErrorHandle
    
    EnableButton False
    ConnectServer 'connect server on demand
     
    Dim i As Integer
    i = 1
    Do While (i <= lstMail.ListItems.Count)
        Dim Item As ListItem
        Set Item = lstMail.ListItems(i)
        If Item.Selected Then
            Dim oInfo As EAGetMailObjLib.MailInfo
            Set oInfo = Item.Tag
            If oInfo.Read Then
                oClient.MarkAsRead Item.Tag, False
                Item.Bold = True
                Item.ForeColor = &H80000008
            End If
        End If
        i = i + 1
    Loop
        
    EnableButton True
    Exit Sub
ErrorHandle:
    oClient.Close  'error, reconnect server next time
    MsgBox Err.Description
    EnableButton True
End Sub

Private Sub btnUndelete_Click()
    Dim oNode As Node
    Set oNode = lstAccount.SelectedItem
    If oNode Is Nothing Then
        Exit Sub
    End If
    
    If oNode.Parent Is Nothing Then
        Exit Sub
    End If
    
    Dim oFolder As EAGetMailObjLib.Imap4Folder
    Set oFolder = oNode.Tag
    

    If SelectedCounts() < 1 Then
        Exit Sub
    End If
    
On Error GoTo ErrorHandle

    EnableButton False
    ConnectServer 'connect server on demand
    
    Dim i As Integer
    i = 1
    Do While (i <= lstMail.ListItems.Count)
        Dim Item As ListItem
        Set Item = lstMail.ListItems(i)
        If Item.Selected Then
            Dim oInfo As EAGetMailObjLib.MailInfo
            Set oInfo = Item.Tag
            If oInfo.Deleted Then
                oClient.Undelete Item.Tag
                Item.Bold = False
                Item.ForeColor = &H80000008
                If Not oInfo.Read Then
                    Item.Bold = True
                End If
            End If
        End If
        i = i + 1
    Loop

    EnableButton True
    Exit Sub
ErrorHandle:
    oClient.Close  'error, reconnect server next time
    MsgBox Err.Description
    EnableButton True
End Sub


'Delete email
Private Sub btnDelete_Click()
    Dim oNode As Node
    Set oNode = lstAccount.SelectedItem
    If oNode Is Nothing Then
        Exit Sub
    End If
    
    If oNode.Parent Is Nothing Then
        Exit Sub
    End If
    
    If SelectedCounts() < 1 Then
        Exit Sub
    End If
    
    Dim oFolder As EAGetMailObjLib.Imap4Folder
    Set oFolder = oNode.Tag
    
    
On Error GoTo ErrorHandle
    EnableButton False
    
    ConnectServer 'connect server on demand
    
    Dim i As Integer
    i = 1
    Do While (i <= lstMail.ListItems.Count)
        Dim Item As ListItem
        Set Item = lstMail.ListItems(i)
        If Item.Selected Then
            oClient.Delete Item.Tag
            Item.Bold = False
            Item.ForeColor = &H80000010
            If oCurServer.Protocol <> MailServerImap4 Then
                ' For Exchange Web Service and WebDAV, there is no undelete
                ' so remove it here.
                lstMail.ListItems.Remove i
            Else
                i = i + 1
            End If
        Else
             i = i + 1
        End If
    Loop

    EnableButton True
    Exit Sub
ErrorHandle:
    oClient.Close  'error, reconnect server next time
    MsgBox Err.Description
    EnableButton True
End Sub

Private Sub btnStart_Click()

    lstMail.ListItems.Clear
    lstAccount.Nodes.Clear
    
    EnableButton False
    m_bCancel = False
    btnCancel.Enabled = False
    
    If textServer.Text = "" Or textUser.Text = "" Or textPassword.Text = "" Then
        MsgBox "Please input server, user, password!"
        Exit Sub
    End If
    
    btnStart.Enabled = False
    
    Dim oServer As New EAGetMailObjLib.MailServer
    oServer.Server = textServer.Text
    oServer.AuthType = lstAuthType.ListIndex
    oServer.User = textUser.Text
    oServer.Password = textPassword.Text
    oServer.Protocol = lstProtocol.ListIndex + 1
    oServer.SSLConnection = chkSSL.Value
    
    If oServer.SSLConnection And oServer.Protocol = MailServerImap4 Then
        oServer.Port = 993
    Else
        oServer.Port = 143
    End If
    
On Error GoTo ErrorHandle
    Set oClient = New EAGetMailObjLib.MailClient
    oClient.LicenseCode = "TryIt"
    Set oCurServer = oServer
    btnCancel.Enabled = True
    
    m_bCancel = False
    
    oClient.Connect oCurServer
    
    ShowFolders
    EnableButton True
    Exit Sub
ErrorHandle:
    oClient.Close 'error, reconnect server next time
    MsgBox Err.Description
    lstAccount.Nodes.Clear
    
    btnStart.Enabled = True
    btnCancel.Enabled = False
    
End Sub

Private Sub ShowFolders()

On Error GoTo ErrorHandle
    lstAccount.Nodes.Clear
    
    Dim oNode As Node
    Set oNode = lstAccount.Nodes.Add(, , "Root")
    oNode.Text = oCurServer.Server & "\" & oCurServer.User
    
    lstAccount.SelectedItem = oNode
    ConnectServer 'connect server on demand
    
    ExpendFolders oClient.Imap4Folders, oNode.Key
    oNode.Expanded = True
    EnableButton True
    Exit Sub
ErrorHandle:
    oClient.Close  'error, reconnect server next time
    MsgBox Err.Description
    EnableButton True
End Sub

Sub ExpendFoldersEx(oFolders, NodeName)
    Dim i
    For i = LBound(oFolders) To UBound(oFolders)
        Dim oFolder As EAGetMailObjLib.Imap4Folder
        Set oFolder = oFolders(i)
        Dim oNode As Node
        Set oNode = Dialog1.lstFolder.Nodes.Add(NodeName, tvwChild, oFolder.ServerPath)
        oNode.Text = oFolder.Name
        Set oNode.Tag = oFolder
        oNode.Key = oFolder.FullPath
        
        ExpendFoldersEx oFolder.SubFolders, oNode.Key
        oNode.Expanded = True
    Next
End Sub

Sub ExpendFolders(oFolders, NodeName)
    
    Dim i
    For i = LBound(oFolders) To UBound(oFolders)
        Dim oFolder As EAGetMailObjLib.Imap4Folder
        Set oFolder = oFolders(i)
        Dim oNode As Node
        Set oNode = lstAccount.Nodes.Add(NodeName, tvwChild, oFolder.ServerPath)
        oNode.Text = oFolder.Name
        Set oNode.Tag = oFolder
        oNode.Key = oFolder.FullPath
        
        ExpendFolders oFolder.SubFolders, oNode.Key
        oNode.Expanded = True
    Next
End Sub

' Upload email file to server.
Private Sub btnUpload_Click()
    Dim oNode As Node
 
    Set oNode = lstAccount.SelectedItem
    If oNode Is Nothing Then
        Exit Sub
    End If
    
    If oNode.Parent Is Nothing Then
        Exit Sub
    End If
    
    FileDlg.ShowOpen
    
    If FileDlg.fileName = vbNullString Or FileDlg.fileName = "" Then
        Exit Sub
    End If
    
    Dim fileName As String
    fileName = FileDlg.fileName
    
    Dim oFolder As EAGetMailObjLib.Imap4Folder
    Set oFolder = oNode.Tag
    
    
On Error GoTo ErrorHandle
    EnableButton False
    
    ConnectServer 'connect server on demand
    
    Dim oMail As New EAGetMailObjLib.Mail
    oMail.LicenseCode = "TryIt"
    oMail.LoadFile fileName, False
    oClient.Append oFolder, oMail.Content
    
    oClient.RefreshMailInfos
    LoadServerMails oNode, oFolder
    
    EnableButton True
    Exit Sub
ErrorHandle:
    oClient.Close  'error, reconnect server next time
    MsgBox Err.Description
    EnableButton True
    
End Sub

Private Sub Form_Load()
    webMail.Navigate "about:blank"
    Set oTools = New EAGetMailObjLib.Tools
    Set oUIDLManager = New UIDLManager
    m_curpath = App.Path

    lstAuthType.AddItem "USER/LOGIN"
    lstAuthType.AddItem "APOP(CRAM-MD5)"
    lstAuthType.AddItem "NTLM"
    lstAuthType.ListIndex = 0
    
    lstProtocol.AddItem "IMAP4"
    lstProtocol.AddItem "Exchange Web Service - 2007/2010"
    lstProtocol.AddItem "Exchange WebDAV - 2000/2003"
    lstProtocol.ListIndex = 0
    
    m_bCancel = False
    statusBar.SimpleText = ""
    lstAccount.Nodes.Clear
    m_bCancel = False

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'm_bCancel = True
    End
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.Width < 13800 Then
        Me.Width = 13800
    End If
    
    If Me.Height < 8400 Then
        Me.Height = 8400
    End If
    
   lstMail.Width = Me.Width - 6500
    webMail.Width = lstMail.Width
    pgBar.Width = webMail.Width
  
    webMail.Height = Me.Height - (lstMail.Height + btnDelete.Height + 2000)
    pgBar.Top = Me.Height - 1200
    lstAccount.Height = Me.Height - 1000
End Sub

' rename folder
Private Sub lstAccount_AfterLabelEdit(Cancel As Integer, NewString As String)
    Dim oNode As Node
    Set oNode = lstAccount.SelectedItem

    If oNode.Text = NewString Then
        Cancel = 1
        Exit Sub
    End If
On Error GoTo ErrorHandle
    EnableButton False
    ConnectServer 'connect server on demand
    oClient.RenameFolder oNode.Tag, NewString
    oNode.Text = NewString
    Cancel = 0
    EnableButton True
    Exit Sub
    
ErrorHandle:
    oClient.Close  'error, reconnect server next time
    MsgBox Err.Description
    EnableButton True
    Cancel = 1
End Sub

' Show folder context menu
Private Sub lstAccount_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> vbRightButton Then
        Exit Sub
    End If
    
    If lstAccount.SelectedItem Is Nothing Then
        Exit Sub
    End If
    
    PopupMenu mnuFolders
End Sub

Private Sub lstAccount_NodeClick(ByVal Node As MSComctlLib.Node)
    lstMail.ListItems.Clear
    If Node.Parent Is Nothing Then
        Exit Sub
    End If
    
    EnableButton False

On Error GoTo ErrorHandle
    LoadServerMails Node, Node.Tag
    EnableButton True
    Exit Sub
ErrorHandle:
    oClient.Close  'error, reconnect server next time
    MsgBox Err.Description
    EnableButton True

End Sub

Private Sub LoadServerMails(oNode As MSComctlLib.Node, oFolder As Imap4Folder)
    
    lstMail.ListItems.Clear
    
    If oFolder.NoSelect Then
        oNode.Expanded = True
        Exit Sub
    End If
    
On Error GoTo ErrorHandle
    
    
    Dim folder As String
    folder = m_curpath & "\" & oCurServer.Server & "\" & oCurServer.User & "\" & oFolder.LocalPath
    CreateFullFolder folder
    ConnectServer 'connect server on demand
    
    statusBar.SimpleText = "Refreshing email(s) ..."
    oClient.SelectFolder oFolder
    oUIDLManager.Load folder & "\uidl.txt"
    Dim infos
    infos = oClient.GetMailInfos()
  
    ' Remove the local uidl which is not existed on the server.
    oUIDLManager.SyncUIDL oCurServer, infos
    oUIDLManager.Update
    
    'remove the mail on local disk which is not existed on server
    ClearLocalMails infos, folder
    
    Dim i, Count
    Count = UBound(infos)
    For i = LBound(infos) To UBound(infos)
        statusBar.SimpleText = "Retrieve summary " & (i + 1) & "/" & (Count + 1) & "..."
        Dim oInfo As EAGetMailObjLib.MailInfo
        Set oInfo = infos(i)
        Dim localfile
        Dim fileName
        fileName = oTools.GenFileName(i) & ".eml"
        localfile = folder & "\" & fileName
        
        Dim uidl_item As UIDLItem
        Set uidl_item = oUIDLManager.FindUIDL(oCurServer, oInfo.UIDL)
        If Not (uidl_item Is Nothing) Then
                localfile = folder & "\" & uidl_item.fileName
        End If
            
        Dim oMail As New EAGetMailObjLib.Mail
        oMail.LicenseCode = "TryIt"
        
        If oTools.ExistFile(localfile) Then
            oMail.LoadFile localfile, True
        Else
            oMail.Load oClient.GetMailHeader(oInfo)
            oMail.SaveAs localfile, True
        End If
        
        Dim oItem As ListItem
        Set oItem = lstMail.ListItems.Add()
        oItem.Text = oMail.From.Name & "<" & oMail.From.Address & ">"
        Set oItem.Tag = oInfo
        oItem.ListSubItems.Add , , oMail.Subject
        oItem.ListSubItems.Add , , oMail.ReceivedDate
        oItem.ListSubItems.Add , , CDbl(oMail.ReceivedDate)
        
       
        If oInfo.Deleted Then
            oItem.ForeColor = &H80000010
        ElseIf Not oInfo.Read Then
            oItem.Bold = True
        End If
        
        If uidl_item Is Nothing Then
               oUIDLManager.AddUIDL oCurServer, oInfo.UIDL, fileName
        End If
        
    Next
    
    oUIDLManager.Update
    statusBar.SimpleText = "Total " & lstMail.ListItems.Count & " email(s)"
    Unselectall

    Exit Sub
ErrorHandle:
    Unselectall
    oUIDLManager.Update
    statusBar.SimpleText = "Total " & lstMail.ListItems.Count & " email(s)"
    Err.Raise Err.Number, Err.Source, Err.Description
    
   
End Sub

Private Sub Unselectall()
    Dim i As Integer
    For i = 1 To lstMail.ListItems.Count
        lstMail.ListItems(i).Selected = False
    Next

End Sub

Private Sub ClearLocalMails(infos, folder)

    Dim files
    files = oTools.GetFiles(folder & "\*.eml")
    Dim i, x
    For i = LBound(files) To UBound(files)
        Dim fileName
        fileName = files(i)
        Dim s
        s = fileName
        Dim pos
        pos = InStrRev(s, "\")
        If pos > 0 Then
            s = Mid(s, pos + 1)
        End If
        Dim bFind
        bFind = False
        If Not (oUIDLManager.FindLocalFile(s) Is Nothing) Then
            bFind = True
        End If
        
        'this email has not existed on server, so delete the local file
        If Not bFind Then
            oTools.RemoveFile fileName
            pos = InStrRev(fileName, ".")
            Dim tempFolder, htmlName
            tempFolder = Mid(fileName, 1, pos - 1)
            htmlName = tempFolder + ".htm"
            If oTools.ExistFile(htmlName) Then
                oTools.RemoveFile htmlName
            End If
            If oTools.ExistFile(tempFolder) Then
                oTools.RemoveFolder tempFolder, True
            End If
        End If
    Next
End Sub

Private Sub CreateFullFolder(folder As String)
    If oTools.ExistFile(folder) Then
        Exit Sub
    End If
    Dim pos As Integer
    pos = 1
        
    Dim s As String
    Do While True
        pos = InStr(pos, folder, "\")
        If pos > 3 Then
            s = Mid(folder, 1, pos - 1)
            If Not oTools.ExistFile(s) Then
                oTools.CreateFolder s
            End If
        ElseIf pos < 1 Then
            Exit Do
        End If
        pos = pos + 1
    Loop
    
    If Not oTools.ExistFile(folder) Then
        oTools.CreateFolder folder
    End If
End Sub

Private Function SelectedCounts()
    Dim i, nSelected As Integer
    nSelected = 0
    For i = 1 To lstMail.ListItems.Count
        If lstMail.ListItems(i).Selected = True Then
            nSelected = nSelected + 1
        End If
    Next
    
    SelectedCounts = nSelected
End Function

Private Sub lstMail_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    Dim oNode As Node
    Set oNode = lstAccount.SelectedItem
    If oNode Is Nothing Then
        Exit Sub
    End If
    
    If oNode.Parent Is Nothing Then
        Exit Sub
    End If
    
    If SelectedCounts() <> 1 Then
        Exit Sub
    End If
    
    Dim OriginalStatus As String
    OriginalStatus = statusBar.SimpleText
On Error GoTo ErrorHandle

    EnableButton False

    Dim oFolder As EAGetMailObjLib.Imap4Folder
    Set oFolder = oNode.Tag
    
    Dim oInfo As EAGetMailObjLib.MailInfo
    Set oInfo = Item.Tag
    
    ConnectServer 'connect server on demand
          
    Dim folder As String
    folder = m_curpath & "\" & oCurServer.Server & "\" & oCurServer.User & "\" & oFolder.LocalPath
    CreateFullFolder folder
    
    ' Find current email record in UIDL file.
    Dim uidl_item As UIDLItem
    Set uidl_item = oUIDLManager.FindUIDL(oCurServer, oInfo.UIDL)
    If (uidl_item Is Nothing) Then
        ' show never happen except you delete the file from the folder manually.
        MsgBox ("No email file found!")
        'EnableIdle(TRUE);
        Exit Sub
    End If
    Dim fileName As String
    ' Get the  local file name for this email UIDL
    fileName = folder & "\" & uidl_item.fileName
    
    Dim pos As Integer
    pos = InStrRev(fileName, ".")
    Dim tempFolder, htmlName
    tempFolder = Mid(fileName, 1, pos - 1)
    htmlName = tempFolder + ".htm"
    
    ' only mail header is retrieved, now retrieve full content of mail.
    If Not oTools.ExistFile(htmlName) Then
        Dim oMail As EAGetMailObjLib.Mail
        pgBar.Min = 0
        pgBar.Max = 100
        pgBar.Value = 0
        statusBar.SimpleText = "Downloading Email " & oInfo.Index & "..."
        Set oMail = oClient.GetMail(oInfo)
        oMail.SaveAs fileName, True
    

    End If
    
    If Not oInfo.Read Then
        oClient.MarkAsRead oInfo, True
        Item.Bold = False
    End If
        
    If oInfo.Deleted Then
        Item.ForeColor = &H80000010
    End If
    
    statusBar.SimpleText = OriginalStatus
    ShowMail fileName
    EnableButton True
    Exit Sub
ErrorHandle:
    oClient.Close  'error, reconnect server next time
    MsgBox Err.Description
    EnableButton True
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
    Dim i, Count, x
    
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

'========================================================
' fnTrim
'========================================================
Sub fnTrim(ByRef src, trimer)
    Dim i, nCount, ch
    nCount = Len(src)
    For i = 1 To nCount
        ch = Mid(src, i, 1)
        If InStr(1, trimer, ch) < 1 Then
            Exit For
        End If
    Next
    
    src = Mid(src, i)
    nCount = Len(src)
    For i = nCount To 1 Step -1
        ch = Mid(src, i, 1)
        If InStr(1, trimer, ch) < 1 Then
            Exit For
        End If
    Next
    src = Mid(src, 1, i)
End Sub
'=============================================================================================
' End of Parse and Display Mails
'=============================================================================================
' referesh emails
Private Sub menuRefreshMails_Click()
    Dim oNode As Node
 
    Set oNode = lstAccount.SelectedItem
    If oNode Is Nothing Then
        Exit Sub
    End If
    
    If oNode.Parent Is Nothing Then
        Exit Sub
    End If
    
On Error GoTo ErrorHandle
    EnableButton False
    Dim oFolder As EAGetMailObjLib.Imap4Folder
    Set oFolder = oNode.Tag
    
    ConnectServer 'connect server on demand
    oClient.SelectFolder oFolder
    oClient.RefreshMailInfos
    
    LoadServerMails oNode, oFolder
    EnableButton True
    Exit Sub
ErrorHandle:
    oClient.Close  'error, reconnect server next time
    MsgBox Err.Description
    EnableButton True
End Sub
' add folder
Private Sub mnuAddFolder_Click()
    Dialog.textFolder.Text = ""
    Dialog.Show 1, Me
    Dim folderName As String
    folderName = Dialog.textFolder
    If folderName = "" Then
        Exit Sub
    End If
    Dim oNode As Node
    Set oNode = lstAccount.SelectedItem
    Dim oFolder As EAGetMailObjLib.Imap4Folder
    If oNode Is Nothing Then
        Set oNode = lstAccount.Nodes.Item(1)
    End If
    
On Error GoTo ErrorHandle

    EnableButton False
    ConnectServer 'connect server on demand

    If oNode.Parent Is Nothing Then
        Set oFolder = oClient.CreateFolder(Nothing, folderName)
    Else
        Set oFolder = oClient.CreateFolder(oNode.Tag, folderName)
    End If
    
    Dim oNewNode As Node
    Set oNewNode = lstAccount.Nodes.Add(oNode.Key, tvwChild, oFolder.ServerPath, oFolder.Name)
    Set oNewNode.Tag = oFolder
    
    EnableButton True
    Exit Sub
ErrorHandle:
    oClient.Close  'error, reconnect server next time
    MsgBox Err.Description
    EnableButton True
End Sub
'Delete folder
Private Sub mnuDeleteFolder_Click()
    Dim oNode As Node
    Set oNode = lstAccount.SelectedItem
    Dim oFolder As EAGetMailObjLib.Imap4Folder
    If oNode Is Nothing Then
        Exit Sub
    End If

    If oNode.Parent Is Nothing Then
        Exit Sub
    End If
    
On Error GoTo ErrorHandle

    EnableButton False

    ConnectServer
    oClient.DeleteFolder oNode.Tag
    lstAccount.Nodes.Remove oNode.Index
    EnableButton True
    Exit Sub
ErrorHandle:
    EnableButton True
    oClient.Close 'error, reconnect server next time
    MsgBox Err.Description
End Sub
'refresh folder
Private Sub mnuRefreshFolder_Click()
On Error GoTo ErrorHandle
    EnableButton False
    lstMail.ListItems.Clear
    oClient.RefreshFolders
    ShowFolders
    EnableButton True
    Exit Sub
ErrorHandle:
    oClient.Close  'error, reconnect server next time
    MsgBox Err.Description
    EnableButton True
    
End Sub

Private Sub mnuRenameFolder_Click()
    Dim oNode As Node
    Set oNode = lstAccount.SelectedItem
    Dim oFolder As EAGetMailObjLib.Imap4Folder
    If oNode Is Nothing Then
        Exit Sub
    End If

    If oNode.Parent Is Nothing Then
        Exit Sub
    End If
    lstAccount.StartLabelEdit
End Sub

'================================================================================================
'"EAGetMail Event Handler"
' you don't have to use EAGetMail Event, but using Event make your application more user friendly
'================================================================================================
Private Sub oClient_OnAuthorized(ByVal oSender As Object, Cancel As Boolean)
        statusBar.SimpleText = "Authorized"
        Cancel = m_bCancel
End Sub

Private Sub oClient_OnConnected(ByVal oSender As Object, Cancel As Boolean)
        statusBar.SimpleText = "Connected"
        Cancel = m_bCancel
End Sub

Private Sub oClient_OnIdle(ByVal oSender As Object, Cancel As Boolean)
        DoEvents
        Cancel = m_bCancel
End Sub

Private Sub oClient_OnQuit(ByVal oSender As Object, Cancel As Boolean)
        statusBar.SimpleText = "Quiting ... "
End Sub

Private Sub oClient_OnReceivingDataStream(ByVal oSender As Object, _
ByVal oInfo As Object, ByVal Received As Long, ByVal Total As Long, Cancel As Boolean)
        pgBar.Min = 0
        pgBar.Max = 100
        pgBar.Value = (Received / Total) * 100
        DoEvents
        Cancel = m_bCancel
End Sub

Private Sub oClient_OnSendingDataStream(ByVal oSender As Object, ByVal Sent As Long, ByVal Total As Long, Cancel As Boolean)
        pgBar.Min = 0
        pgBar.Max = 100
        pgBar.Value = (Sent / Total) * 100
        DoEvents
        Cancel = m_bCancel
End Sub

Private Sub oClient_OnSecuring(ByVal oSender As Object, Cancel As Boolean)
        statusBar.SimpleText = "Securing ..."
        Cancel = m_bCancel
End Sub
'================================================================================================
'"EAGetMail Event Handler" end
'================================================================================================



' By default, Exchange Web Service requires SSL connection.
'For other protocol, please set SSL connection based on your server setting manually
Private Sub lstProtocol_Click()
    If lstProtocol.ListIndex + 1 = MailServerEWS Then
        MsgBox "By default, Exchange Web Service requires SSL!"
        chkSSL.Value = Checked 'By default, Exchange Web Service requires SSL
    End If
End Sub

