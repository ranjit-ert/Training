<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Public Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    private mainMenu1 As System.Windows.Forms.MainMenu

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.mainMenu1 = New System.Windows.Forms.MainMenu
        Me.textServer = New System.Windows.Forms.TextBox
        Me.textUser = New System.Windows.Forms.TextBox
        Me.textPassword = New System.Windows.Forms.TextBox
        Me.chkSSL = New System.Windows.Forms.CheckBox
        Me.lstAuthType = New System.Windows.Forms.ComboBox
        Me.lstProtocol = New System.Windows.Forms.ComboBox
        Me.chkLeaveCopy = New System.Windows.Forms.CheckBox
        Me.btnStart = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.pgBar = New System.Windows.Forms.ProgressBar
        Me.Label1 = New System.Windows.Forms.Label
        Me.label2 = New System.Windows.Forms.Label
        Me.label3 = New System.Windows.Forms.Label
        Me.label4 = New System.Windows.Forms.Label
        Me.label5 = New System.Windows.Forms.Label
        Me.lblStatus = New System.Windows.Forms.Label
        Me.lstMail = New System.Windows.Forms.ListView
        Me.colFrom = New System.Windows.Forms.ColumnHeader
        Me.colSubject = New System.Windows.Forms.ColumnHeader
        Me.colDate = New System.Windows.Forms.ColumnHeader
        Me.btnDel = New System.Windows.Forms.Button
        Me.label6 = New System.Windows.Forms.Label
        Me.webMail = New System.Windows.Forms.WebBrowser
        Me.SuspendLayout()
        '
        'textServer
        '
        Me.textServer.Location = New System.Drawing.Point(71, 14)
        Me.textServer.Name = "textServer"
        Me.textServer.Size = New System.Drawing.Size(103, 21)
        Me.textServer.TabIndex = 0
        '
        'textUser
        '
        Me.textUser.Location = New System.Drawing.Point(71, 55)
        Me.textUser.Name = "textUser"
        Me.textUser.Size = New System.Drawing.Size(103, 21)
        Me.textUser.TabIndex = 1
        '
        'textPassword
        '
        Me.textPassword.Location = New System.Drawing.Point(71, 95)
        Me.textPassword.Name = "textPassword"
        Me.textPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.textPassword.Size = New System.Drawing.Size(103, 21)
        Me.textPassword.TabIndex = 2
        '
        'chkSSL
        '
        Me.chkSSL.Location = New System.Drawing.Point(71, 137)
        Me.chkSSL.Name = "chkSSL"
        Me.chkSSL.Size = New System.Drawing.Size(128, 20)
        Me.chkSSL.TabIndex = 3
        Me.chkSSL.Text = "SSL Connection"
        '
        'lstAuthType
        '
        Me.lstAuthType.Location = New System.Drawing.Point(71, 163)
        Me.lstAuthType.Name = "lstAuthType"
        Me.lstAuthType.Size = New System.Drawing.Size(103, 22)
        Me.lstAuthType.TabIndex = 4
        '
        'lstProtocol
        '
        Me.lstProtocol.Location = New System.Drawing.Point(71, 205)
        Me.lstProtocol.Name = "lstProtocol"
        Me.lstProtocol.Size = New System.Drawing.Size(103, 22)
        Me.lstProtocol.TabIndex = 5
        '
        'chkLeaveCopy
        '
        Me.chkLeaveCopy.Location = New System.Drawing.Point(71, 247)
        Me.chkLeaveCopy.Name = "chkLeaveCopy"
        Me.chkLeaveCopy.Size = New System.Drawing.Size(248, 20)
        Me.chkLeaveCopy.TabIndex = 6
        Me.chkLeaveCopy.Text = "Leave a copy of message on server"
        '
        'btnStart
        '
        Me.btnStart.Location = New System.Drawing.Point(71, 284)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(75, 20)
        Me.btnStart.TabIndex = 7
        Me.btnStart.Text = "Start"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(160, 284)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 20)
        Me.btnCancel.TabIndex = 8
        Me.btnCancel.Text = "Cancel"
        '
        'pgBar
        '
        Me.pgBar.Location = New System.Drawing.Point(71, 340)
        Me.pgBar.Name = "pgBar"
        Me.pgBar.Size = New System.Drawing.Size(167, 10)
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(3, 14)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(62, 20)
        Me.Label1.Text = "Server"
        '
        'label2
        '
        Me.label2.Location = New System.Drawing.Point(3, 58)
        Me.label2.Name = "label2"
        Me.label2.Size = New System.Drawing.Size(62, 20)
        Me.label2.Text = "User"
        '
        'label3
        '
        Me.label3.Location = New System.Drawing.Point(3, 98)
        Me.label3.Name = "label3"
        Me.label3.Size = New System.Drawing.Size(62, 20)
        Me.label3.Text = "Password"
        '
        'label4
        '
        Me.label4.Location = New System.Drawing.Point(3, 163)
        Me.label4.Name = "label4"
        Me.label4.Size = New System.Drawing.Size(68, 20)
        Me.label4.Text = "Auth Type"
        '
        'label5
        '
        Me.label5.Location = New System.Drawing.Point(3, 205)
        Me.label5.Name = "label5"
        Me.label5.Size = New System.Drawing.Size(62, 20)
        Me.label5.Text = "Protocol"
        '
        'lblStatus
        '
        Me.lblStatus.Location = New System.Drawing.Point(71, 307)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(248, 20)
        '
        'lstMail
        '
        Me.lstMail.Columns.Add(Me.colFrom)
        Me.lstMail.Columns.Add(Me.colSubject)
        Me.lstMail.Columns.Add(Me.colDate)
        Me.lstMail.FullRowSelect = True
        Me.lstMail.Location = New System.Drawing.Point(317, 14)
        Me.lstMail.Name = "lstMail"
        Me.lstMail.Size = New System.Drawing.Size(387, 190)
        Me.lstMail.TabIndex = 21
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
        Me.btnDel.Location = New System.Drawing.Point(630, 210)
        Me.btnDel.Name = "btnDel"
        Me.btnDel.Size = New System.Drawing.Size(72, 20)
        Me.btnDel.TabIndex = 22
        Me.btnDel.Text = "Delete"
        '
        'label6
        '
        Me.label6.ForeColor = System.Drawing.SystemColors.Highlight
        Me.label6.Location = New System.Drawing.Point(14, 367)
        Me.label6.Name = "label6"
        Me.label6.Size = New System.Drawing.Size(291, 54)
        Me.label6.Text = "Warning: if ""leave a copy of message on server"" is not checked,  the emails on th" & _
            "e server will be deleted !"
        '
        'webMail
        '
        Me.webMail.Location = New System.Drawing.Point(317, 236)
        Me.webMail.Name = "webMail"
        Me.webMail.Size = New System.Drawing.Size(387, 200)
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(722, 448)
        Me.Controls.Add(Me.webMail)
        Me.Controls.Add(Me.label6)
        Me.Controls.Add(Me.btnDel)
        Me.Controls.Add(Me.lstMail)
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.label5)
        Me.Controls.Add(Me.label4)
        Me.Controls.Add(Me.label3)
        Me.Controls.Add(Me.label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.pgBar)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnStart)
        Me.Controls.Add(Me.chkLeaveCopy)
        Me.Controls.Add(Me.lstProtocol)
        Me.Controls.Add(Me.lstAuthType)
        Me.Controls.Add(Me.chkSSL)
        Me.Controls.Add(Me.textPassword)
        Me.Controls.Add(Me.textUser)
        Me.Controls.Add(Me.textServer)
        Me.Menu = Me.mainMenu1
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents textServer As System.Windows.Forms.TextBox
    Friend WithEvents textUser As System.Windows.Forms.TextBox
    Friend WithEvents textPassword As System.Windows.Forms.TextBox
    Friend WithEvents chkSSL As System.Windows.Forms.CheckBox
    Friend WithEvents lstAuthType As System.Windows.Forms.ComboBox
    Friend WithEvents lstProtocol As System.Windows.Forms.ComboBox
    Friend WithEvents chkLeaveCopy As System.Windows.Forms.CheckBox
    Friend WithEvents btnStart As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents pgBar As System.Windows.Forms.ProgressBar
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents label2 As System.Windows.Forms.Label
    Friend WithEvents label3 As System.Windows.Forms.Label
    Friend WithEvents label4 As System.Windows.Forms.Label
    Friend WithEvents label5 As System.Windows.Forms.Label
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents lstMail As System.Windows.Forms.ListView
    Friend WithEvents btnDel As System.Windows.Forms.Button
    Friend WithEvents label6 As System.Windows.Forms.Label
    Friend WithEvents colFrom As System.Windows.Forms.ColumnHeader
    Friend WithEvents colSubject As System.Windows.Forms.ColumnHeader
    Friend WithEvents colDate As System.Windows.Forms.ColumnHeader
    Friend webMail As System.Windows.Forms.WebBrowser

End Class
