//  ===============================================================================
// |    THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF      |
// |    ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO    |
// |    THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A         |
// |    PARTICULAR PURPOSE.                                                    |
// |    Copyright (c)2006-2012  ADMINSYSTEM SOFTWARE LIMITED                         |
// |
// |    Project: It demonstrates how to use EAGetMail to receive/parse email.
// |        
// |
// |    File: Form1 : implementation file
// |
// |    Author: Ivan Lui ( ivan@emailarchitect.net )
//  ===============================================================================

using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using EAGetMail;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace imap4.csharp
{
    /// <summary>
    /// Summary description for Form1.
    /// </summary>
    public class Form1 : System.Windows.Forms.Form, IComparer
    {
        private System.Windows.Forms.TreeView trAccounts;
        private System.Windows.Forms.ListView lstMail;
        private System.Windows.Forms.ColumnHeader colFrom;
        private System.Windows.Forms.ColumnHeader colSubject;
        private System.Windows.Forms.ColumnHeader colDate;
        private System.Windows.Forms.Button btnDelete;
        private System.Windows.Forms.Button btnUndelete;
        private System.Windows.Forms.Button btnUnread;
        private System.Windows.Forms.Button btnPure;
        private System.Windows.Forms.Button btnMove;
        private System.Windows.Forms.ProgressBar pgBar;
        private System.Windows.Forms.Label lblStatus;
        private IContainer components;
        
        // For evaluation usage, please use "TryIt" as the license code, otherwise the 
        // "Invalid license code" exception will be thrown. However, the object will expire in 1-2 months, then
        // "Trial version expired" exception will be thrown.
        private MailClient oClient = new MailClient("TryIt");
        private MailServer oCurServer = null;
        private UIDLManager oUIDLManager = null;

        private bool m_bcancel = false;
        private string m_curpath = "";
        private WebBrowser webMail;
        private GroupBox groupBox1;
        private Label label1;
        private Button btnCancel;
        private Button btnStart;
        private ComboBox lstProtocol;
        private Label label5;
        private ComboBox lstAuthType;
        private Label label4;
        private CheckBox chkSSL;
        private TextBox textPassword;
        private TextBox textUser;
        private TextBox textServer;
        private Label label3;
        private Label label2;
        private Label label6;
        private Button btnQuit;
        private ContextMenuStrip contextMenuFolder;
        private ToolStripMenuItem refreshFoldersToolStripMenuItem;
        private ToolStripMenuItem refreshMailsToolStripMenuItem;
        private ToolStripMenuItem addFolderToolStripMenuItem;
        private ToolStripMenuItem deleteFolderToolStripMenuItem;
        private ToolStripMenuItem renameFolderToolStripMenuItem;
        private Button btnUpload;
        private Button btnCopy;
        private OpenFileDialog openFileDialog1;
        public Form1()
        {
            //
            // Required for Windows Form Designer support
            //
            InitializeComponent();
            //
            // TODO: Add any constructor code after InitializeComponent call
            //
        }

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.trAccounts = new System.Windows.Forms.TreeView();
            this.contextMenuFolder = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.refreshFoldersToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.refreshMailsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.addFolderToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.deleteFolderToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.renameFolderToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.lstMail = new System.Windows.Forms.ListView();
            this.colFrom = new System.Windows.Forms.ColumnHeader();
            this.colSubject = new System.Windows.Forms.ColumnHeader();
            this.colDate = new System.Windows.Forms.ColumnHeader();
            this.btnDelete = new System.Windows.Forms.Button();
            this.btnUndelete = new System.Windows.Forms.Button();
            this.btnUnread = new System.Windows.Forms.Button();
            this.btnPure = new System.Windows.Forms.Button();
            this.btnMove = new System.Windows.Forms.Button();
            this.pgBar = new System.Windows.Forms.ProgressBar();
            this.lblStatus = new System.Windows.Forms.Label();
            this.webMail = new System.Windows.Forms.WebBrowser();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnQuit = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnStart = new System.Windows.Forms.Button();
            this.lstProtocol = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.lstAuthType = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.chkSSL = new System.Windows.Forms.CheckBox();
            this.textPassword = new System.Windows.Forms.TextBox();
            this.textUser = new System.Windows.Forms.TextBox();
            this.textServer = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.btnUpload = new System.Windows.Forms.Button();
            this.btnCopy = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.contextMenuFolder.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // trAccounts
            // 
            this.trAccounts.ContextMenuStrip = this.contextMenuFolder;
            this.trAccounts.HideSelection = false;
            this.trAccounts.Location = new System.Drawing.Point(259, 12);
            this.trAccounts.Name = "trAccounts";
            this.trAccounts.Size = new System.Drawing.Size(162, 408);
            this.trAccounts.TabIndex = 0;
            this.trAccounts.AfterLabelEdit += new System.Windows.Forms.NodeLabelEditEventHandler(this.trAccounts_AfterLabelEdit);
            this.trAccounts.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.trAccounts_AfterSelect);
            // 
            // contextMenuFolder
            // 
            this.contextMenuFolder.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.refreshFoldersToolStripMenuItem,
            this.refreshMailsToolStripMenuItem,
            this.addFolderToolStripMenuItem,
            this.deleteFolderToolStripMenuItem,
            this.renameFolderToolStripMenuItem});
            this.contextMenuFolder.Name = "contextMenuFolder";
            this.contextMenuFolder.Size = new System.Drawing.Size(155, 114);
            // 
            // refreshFoldersToolStripMenuItem
            // 
            this.refreshFoldersToolStripMenuItem.Name = "refreshFoldersToolStripMenuItem";
            this.refreshFoldersToolStripMenuItem.Size = new System.Drawing.Size(154, 22);
            this.refreshFoldersToolStripMenuItem.Text = "Refresh Folders";
            this.refreshFoldersToolStripMenuItem.Click += new System.EventHandler(this.refreshFoldersToolStripMenuItem_Click);
            // 
            // refreshMailsToolStripMenuItem
            // 
            this.refreshMailsToolStripMenuItem.Name = "refreshMailsToolStripMenuItem";
            this.refreshMailsToolStripMenuItem.Size = new System.Drawing.Size(154, 22);
            this.refreshMailsToolStripMenuItem.Text = "Refresh Mails";
            this.refreshMailsToolStripMenuItem.Click += new System.EventHandler(this.refreshMailsToolStripMenuItem_Click);
            // 
            // addFolderToolStripMenuItem
            // 
            this.addFolderToolStripMenuItem.Name = "addFolderToolStripMenuItem";
            this.addFolderToolStripMenuItem.Size = new System.Drawing.Size(154, 22);
            this.addFolderToolStripMenuItem.Text = "Add Folder";
            this.addFolderToolStripMenuItem.Click += new System.EventHandler(this.addFolderToolStripMenuItem_Click);
            // 
            // deleteFolderToolStripMenuItem
            // 
            this.deleteFolderToolStripMenuItem.Name = "deleteFolderToolStripMenuItem";
            this.deleteFolderToolStripMenuItem.Size = new System.Drawing.Size(154, 22);
            this.deleteFolderToolStripMenuItem.Text = "Delete Folder";
            this.deleteFolderToolStripMenuItem.Click += new System.EventHandler(this.deleteFolderToolStripMenuItem_Click);
            // 
            // renameFolderToolStripMenuItem
            // 
            this.renameFolderToolStripMenuItem.Name = "renameFolderToolStripMenuItem";
            this.renameFolderToolStripMenuItem.Size = new System.Drawing.Size(154, 22);
            this.renameFolderToolStripMenuItem.Text = "Rename Folder";
            this.renameFolderToolStripMenuItem.Click += new System.EventHandler(this.renameFolderToolStripMenuItem_Click);
            // 
            // lstMail
            // 
            this.lstMail.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.colFrom,
            this.colSubject,
            this.colDate});
            this.lstMail.FullRowSelect = true;
            this.lstMail.HideSelection = false;
            this.lstMail.Location = new System.Drawing.Point(427, 12);
            this.lstMail.Name = "lstMail";
            this.lstMail.Size = new System.Drawing.Size(446, 161);
            this.lstMail.TabIndex = 1;
            this.lstMail.UseCompatibleStateImageBehavior = false;
            this.lstMail.View = System.Windows.Forms.View.Details;
            this.lstMail.Click += new System.EventHandler(this.lstMail_Click);
            // 
            // colFrom
            // 
            this.colFrom.Text = "From";
            this.colFrom.Width = 100;
            // 
            // colSubject
            // 
            this.colSubject.Text = "Subject";
            this.colSubject.Width = 200;
            // 
            // colDate
            // 
            this.colDate.Text = "Date";
            this.colDate.Width = 100;
            // 
            // btnDelete
            // 
            this.btnDelete.Enabled = false;
            this.btnDelete.Location = new System.Drawing.Point(427, 180);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(56, 22);
            this.btnDelete.TabIndex = 8;
            this.btnDelete.Text = "Delete Message";
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // btnUndelete
            // 
            this.btnUndelete.Enabled = false;
            this.btnUndelete.Location = new System.Drawing.Point(489, 180);
            this.btnUndelete.Name = "btnUndelete";
            this.btnUndelete.Size = new System.Drawing.Size(66, 22);
            this.btnUndelete.TabIndex = 9;
            this.btnUndelete.Text = "Undelete";
            this.btnUndelete.Click += new System.EventHandler(this.btnUndelete_Click);
            // 
            // btnUnread
            // 
            this.btnUnread.Enabled = false;
            this.btnUnread.Location = new System.Drawing.Point(561, 180);
            this.btnUnread.Name = "btnUnread";
            this.btnUnread.Size = new System.Drawing.Size(56, 22);
            this.btnUnread.TabIndex = 10;
            this.btnUnread.Text = "Unread";
            this.btnUnread.Click += new System.EventHandler(this.btnUnread_Click);
            // 
            // btnPure
            // 
            this.btnPure.Enabled = false;
            this.btnPure.Location = new System.Drawing.Point(623, 180);
            this.btnPure.Name = "btnPure";
            this.btnPure.Size = new System.Drawing.Size(56, 22);
            this.btnPure.TabIndex = 11;
            this.btnPure.Text = "Pure";
            this.btnPure.Click += new System.EventHandler(this.btnPure_Click);
            // 
            // btnMove
            // 
            this.btnMove.Enabled = false;
            this.btnMove.Location = new System.Drawing.Point(750, 180);
            this.btnMove.Name = "btnMove";
            this.btnMove.Size = new System.Drawing.Size(56, 22);
            this.btnMove.TabIndex = 12;
            this.btnMove.Text = "Move";
            this.btnMove.Click += new System.EventHandler(this.btnMove_Click);
            // 
            // pgBar
            // 
            this.pgBar.Location = new System.Drawing.Point(427, 412);
            this.pgBar.Name = "pgBar";
            this.pgBar.Size = new System.Drawing.Size(445, 10);
            this.pgBar.TabIndex = 13;
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(12, 351);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(38, 13);
            this.lblStatus.TabIndex = 14;
            this.lblStatus.Text = "Ready";
            // 
            // webMail
            // 
            this.webMail.Location = new System.Drawing.Point(427, 211);
            this.webMail.MinimumSize = new System.Drawing.Size(20, 20);
            this.webMail.Name = "webMail";
            this.webMail.Size = new System.Drawing.Size(446, 195);
            this.webMail.TabIndex = 15;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnQuit);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.btnCancel);
            this.groupBox1.Controls.Add(this.btnStart);
            this.groupBox1.Controls.Add(this.lstProtocol);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.lstAuthType);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.chkSSL);
            this.groupBox1.Controls.Add(this.textPassword);
            this.groupBox1.Controls.Add(this.textUser);
            this.groupBox1.Controls.Add(this.textServer);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Location = new System.Drawing.Point(12, 9);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(232, 328);
            this.groupBox1.TabIndex = 16;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Account Information";
            // 
            // btnQuit
            // 
            this.btnQuit.Location = new System.Drawing.Point(12, 252);
            this.btnQuit.Name = "btnQuit";
            this.btnQuit.Size = new System.Drawing.Size(208, 24);
            this.btnQuit.TabIndex = 16;
            this.btnQuit.Text = "Quit";
            this.btnQuit.UseVisualStyleBackColor = true;
            this.btnQuit.Click += new System.EventHandler(this.btnQuit_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(8, 272);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(0, 13);
            this.label1.TabIndex = 14;
            // 
            // btnCancel
            // 
            this.btnCancel.Enabled = false;
            this.btnCancel.Location = new System.Drawing.Point(12, 281);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(208, 24);
            this.btnCancel.TabIndex = 13;
            this.btnCancel.Text = "Cancel Task";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnStart
            // 
            this.btnStart.Location = new System.Drawing.Point(11, 222);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(208, 24);
            this.btnStart.TabIndex = 12;
            this.btnStart.Text = "Start";
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // lstProtocol
            // 
            this.lstProtocol.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.lstProtocol.Location = new System.Drawing.Point(80, 181);
            this.lstProtocol.Name = "lstProtocol";
            this.lstProtocol.Size = new System.Drawing.Size(136, 21);
            this.lstProtocol.TabIndex = 10;
            this.lstProtocol.SelectedIndexChanged += new System.EventHandler(this.lstProtocol_SelectedIndexChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(8, 183);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(46, 13);
            this.label5.TabIndex = 9;
            this.label5.Text = "Protocol";
            // 
            // lstAuthType
            // 
            this.lstAuthType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.lstAuthType.Location = new System.Drawing.Point(80, 149);
            this.lstAuthType.Name = "lstAuthType";
            this.lstAuthType.Size = new System.Drawing.Size(136, 21);
            this.lstAuthType.TabIndex = 8;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(8, 151);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(56, 13);
            this.label4.TabIndex = 7;
            this.label4.Text = "Auth Type";
            // 
            // chkSSL
            // 
            this.chkSSL.Location = new System.Drawing.Point(11, 122);
            this.chkSSL.Name = "chkSSL";
            this.chkSSL.Size = new System.Drawing.Size(208, 16);
            this.chkSSL.TabIndex = 6;
            this.chkSSL.Text = "SSL connection";
            // 
            // textPassword
            // 
            this.textPassword.Location = new System.Drawing.Point(80, 86);
            this.textPassword.Name = "textPassword";
            this.textPassword.PasswordChar = '*';
            this.textPassword.Size = new System.Drawing.Size(136, 20);
            this.textPassword.TabIndex = 5;
            // 
            // textUser
            // 
            this.textUser.Location = new System.Drawing.Point(80, 54);
            this.textUser.Name = "textUser";
            this.textUser.Size = new System.Drawing.Size(136, 20);
            this.textUser.TabIndex = 4;
            // 
            // textServer
            // 
            this.textServer.Location = new System.Drawing.Point(80, 22);
            this.textServer.Name = "textServer";
            this.textServer.Size = new System.Drawing.Size(136, 20);
            this.textServer.TabIndex = 3;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(8, 88);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 13);
            this.label3.TabIndex = 2;
            this.label3.Text = "Password";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(8, 56);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(29, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "User";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(8, 24);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(38, 13);
            this.label6.TabIndex = 0;
            this.label6.Text = "Server";
            // 
            // btnUpload
            // 
            this.btnUpload.Enabled = false;
            this.btnUpload.Location = new System.Drawing.Point(812, 180);
            this.btnUpload.Name = "btnUpload";
            this.btnUpload.Size = new System.Drawing.Size(56, 22);
            this.btnUpload.TabIndex = 17;
            this.btnUpload.Text = "Upload";
            this.btnUpload.UseVisualStyleBackColor = true;
            this.btnUpload.Click += new System.EventHandler(this.btnUpload_Click);
            // 
            // btnCopy
            // 
            this.btnCopy.Enabled = false;
            this.btnCopy.Location = new System.Drawing.Point(687, 180);
            this.btnCopy.Name = "btnCopy";
            this.btnCopy.Size = new System.Drawing.Size(56, 22);
            this.btnCopy.TabIndex = 18;
            this.btnCopy.Text = "Copy";
            this.btnCopy.UseVisualStyleBackColor = true;
            this.btnCopy.Click += new System.EventHandler(this.btnCopy_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            this.openFileDialog1.Filter = "Email File | *.EML";
            // 
            // Form1
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(884, 432);
            this.Controls.Add(this.btnCopy);
            this.Controls.Add(this.btnUpload);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.webMail);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.pgBar);
            this.Controls.Add(this.lstMail);
            this.Controls.Add(this.btnMove);
            this.Controls.Add(this.trAccounts);
            this.Controls.Add(this.btnPure);
            this.Controls.Add(this.btnUnread);
            this.Controls.Add(this.btnDelete);
            this.Controls.Add(this.btnUndelete);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.Resize += new System.EventHandler(this.Form1_Resize);
            this.contextMenuFolder.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        #region EAGetMail Event Handler
        public void OnConnected(object sender, ref bool cancel)
        {
            lblStatus.Text = "Connected";
            cancel = m_bcancel;
            Application.DoEvents();
        }

        public void OnQuit(object sender, ref bool cancel)
        {
            lblStatus.Text = "Quit";
            cancel = m_bcancel;
            Application.DoEvents();
        }

        public void OnReceivingDataStream(object sender, MailInfo info, int received, int total, ref bool cancel)
        {
            pgBar.Minimum = 0;
            pgBar.Maximum = total;
            pgBar.Value = received;
            cancel = m_bcancel;
            Application.DoEvents();
        }

        public void OnSendingDataStream(object sender,  int sent, int total, ref bool cancel)
        {
            pgBar.Minimum = 0;
            pgBar.Maximum = total;
            pgBar.Value = sent;
            cancel = m_bcancel;
            Application.DoEvents();
        }


        public void OnIdle(object sender, ref bool cancel)
        {
            cancel = m_bcancel;
            Application.DoEvents();
        }

        public void OnAuthorized(object sender, ref bool cancel)
        {
            lblStatus.Text = "Authorized";
            cancel = m_bcancel;
            Application.DoEvents();
        }

        public void OnSecuring(object sender, ref bool cancel)
        {
            lblStatus.Text = "Securing";
            cancel = m_bcancel;
            Application.DoEvents();
        }
        #endregion

        #region Parse and Display Mails
        private void ShowMail(string fileName)
        {
            try
            {
                int pos = fileName.LastIndexOf(".");
                string mainName = fileName.Substring(0, pos);
                string htmlName = mainName + ".htm";

                string tempFolder = mainName;
                if (!File.Exists(htmlName))
                {
                    _GenerateHtmlForEmail(htmlName, fileName, tempFolder);
                }

                object empty = System.Reflection.Missing.Value;
                webMail.Navigate(htmlName);
            }
            catch (Exception ep)
            {
                MessageBox.Show(ep.Message);
            }
        }
        private string _FormatHtmlTag(string src)
        {
            src = src.Replace(">", "&gt;");
            src = src.Replace("<", "&lt;");
            return src;
        }

        // We generate a html + attachment folder for every email, once the html is create,
        // next time we don't need to parse the email again.
        private void _GenerateHtmlForEmail(string htmlName, string emlFile, string tempFolder)
        {
            // For evaluation usage, please use "TryIt" as the license code, otherwise the 
            // "Invalid license code" exception will be thrown. However, the object will expire in 1-2 months, then
            // "Trial version expired" exception will be thrown.
            Mail oMail = new Mail("TryIt");
            oMail.Load(emlFile, false);

            if (oMail.IsEncrypted)
            {
                try
                {
                    // This email is encrypted, we decrypt it by user default certificate.
                    // you can also use specified certificate like this
                    // oCert = new Certificate();
                    // oCert.Load("c:\test.pfx", "pfxpassword", Certificate.CertificateKeyLocation.CRYPT_USER_KEYSET)
                    // oMail = oMail.Decrypt( oCert );
                    oMail = oMail.Decrypt(null);
                }
                catch (Exception ep)
                {
                    MessageBox.Show(ep.Message);
                    oMail.Load(emlFile, false);
                }
            }

            if (oMail.IsSigned)
            {
                try
                {
                    //this email is digital signed.
                    EAGetMail.Certificate cert = oMail.VerifySignature();
                    MessageBox.Show("This email contains a valid digital signature.");
                    //you can add the certificate to your certificate storage like this
                    //cert.AddToStore( Certificate.CertificateStoreLocation.CERT_SYSTEM_STORE_CURRENT_USER,
                    //	"addressbook" );
                    // then you can use send the encrypted email back to this sender.
                }
                catch (Exception ep)
                {
                    MessageBox.Show(ep.Message);
                }
            }

            string html = oMail.HtmlBody;
            StringBuilder hdr = new StringBuilder();

            hdr.Append("<font face=\"Courier New,Arial\" size=2>");
            hdr.Append("<b>From:</b> " + _FormatHtmlTag(oMail.From.ToString()) + "<br>");
            MailAddress[] addrs = oMail.To;
            int count = addrs.Length;
            if (count > 0)
            {
                hdr.Append("<b>To:</b> ");
                for (int i = 0; i < count; i++)
                {
                    hdr.Append(_FormatHtmlTag(addrs[i].ToString()));
                    if (i < count - 1)
                    {
                        hdr.Append(";");
                    }
                }
                hdr.Append("<br>");
            }

            addrs = oMail.Cc;

            count = addrs.Length;
            if (count > 0)
            {
                hdr.Append("<b>Cc:</b> ");
                for (int i = 0; i < count; i++)
                {
                    hdr.Append(_FormatHtmlTag(addrs[i].ToString()));
                    if (i < count - 1)
                    {
                        hdr.Append(";");
                    }
                }
                hdr.Append("<br>");
            }

            hdr.Append(String.Format("<b>Subject:</b>{0}<br>\r\n", _FormatHtmlTag(oMail.Subject)));

            Attachment[] atts = oMail.Attachments;
            count = atts.Length;
            if (count > 0)
            {
                if (!Directory.Exists(tempFolder))
                    Directory.CreateDirectory(tempFolder);

                hdr.Append("<b>Attachments:</b>");
                for (int i = 0; i < count; i++)
                {
                    Attachment att = atts[i];
                    //this attachment is in OUTLOOK RTF format, decode it here.
                    if (String.Compare(att.Name, "winmail.dat") == 0)
                    {
                        Attachment[] tatts = null;
                        try
                        {
                            tatts = Mail.ParseTNEF(att.Content, true);
                        }
                        catch (Exception ep)
                        {
                            MessageBox.Show(ep.Message);
                            continue;
                        }
                        int y = tatts.Length;
                        for (int x = 0; x < y; x++)
                        {
                            Attachment tatt = tatts[x];
                            string tattname = String.Format("{0}\\{1}", tempFolder, tatt.Name);
                            tatt.SaveAs(tattname, true);
                            hdr.Append(String.Format("<a href=\"{0}\" target=\"_blank\">{1}</a> ", tattname, tatt.Name));
                        }
                        continue;
                    }
                    string attname = String.Format("{0}\\{1}", tempFolder, att.Name);
                    att.SaveAs(attname, true);
                    hdr.Append(String.Format("<a href=\"{0}\" target=\"_blank\">{1}</a> ", attname, att.Name));
                    if (att.ContentID.Length > 0)
                    {
                        //show embedded image.
                        html = html.Replace("cid:" + att.ContentID, attname);
                    }
                    else if (String.Compare(att.ContentType, 0, "image/", 0, "image/".Length, true) == 0)
                    {
                        //show attached image.
                        html = html + String.Format("<hr><img src=\"{0}\">", attname);
                    }
                }
            }

            Regex reg = new Regex("(<meta[^>]*charset[ \t]*=[ \t\"]*)([^<> \r\n\"]*)", RegexOptions.Multiline | RegexOptions.IgnoreCase);
            html = reg.Replace(html, "$1utf-8");
            if (!reg.IsMatch(html))
            {
                hdr.Insert(0, "<meta HTTP-EQUIV=\"Content-Type\" Content=\"text/html; charset=utf-8\">");
            }

            html = hdr.ToString() + "<hr>" + html;
            FileStream fs = new FileStream(htmlName, FileMode.Create, FileAccess.Write, FileShare.None);
            byte[] data = System.Text.UTF8Encoding.UTF8.GetBytes(html);
            fs.Write(data, 0, data.Length);
            fs.Close();
            oMail.Clear();
        }
        #endregion

        #region IMAP4/Exchange Folders Management
        // Add Folder
        private void addFolderToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form3 frm = new Form3();
            if (frm.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    string folder = frm.textFolder.Text;
                    TreeNode node = trAccounts.SelectedNode;
                    ConnectServer(node);
                    Imap4Folder fd = (node.Parent != null) ?
                        oClient.CreateFolder(node.Tag as Imap4Folder, folder) :
                        oClient.CreateFolder(null, folder);

                    TreeNode newNode = node.Nodes.Add(fd.Name);
                    newNode.Tag = fd;

                }
                catch (Exception ep)
                {
                    MessageBox.Show(ep.Message);
                    oClient.Close();
                }
            }
        }

        // Refresh Folders
        private void refreshFoldersToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TreeNode node = trAccounts.SelectedNode;
            if (node == null)
                return;

            EnableIdle(false);
            try
            {
                while (node.Parent != null)
                    node = node.Parent;

                trAccounts.SelectedNode = null;
                lstMail.Items.Clear();

                ConnectServer(node);
                oClient.RefreshFolders();

                lblStatus.Text = "Refreshing Folders ...";
                Imap4Folder[] fds = oClient.Imap4Folders;
                ExpendFolders(fds, node);
                node.ExpandAll();
                lblStatus.Text = "";
            }
            catch (Exception ep)
            {
                MessageBox.Show(ep.Message);
                oClient.Close();
            }

            EnableIdle(true);
        }

        // Delete folder
        private void deleteFolderToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TreeNode node = trAccounts.SelectedNode;
            if (node == null)
                return;

            if (node.Parent == null)
                return;

            EnableIdle(false);
            try
            {
                ConnectServer(node);
                oClient.DeleteFolder(node.Tag as Imap4Folder);
                Directory.Delete(GetFolderByNode(node), true);
                trAccounts.SelectedNode = null;
                node.Remove();
                lstMail.Items.Clear();

            }
            catch (Exception ep)
            {
                MessageBox.Show(ep.Message);
                oClient.Close();
            }

            EnableIdle(true);
        }

        // Begin rename folder
        private void renameFolderToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TreeNode node = trAccounts.SelectedNode;
            if (node != null && node.Parent != null)
            {
                trAccounts.LabelEdit = true;
                if (!node.IsEditing)
                {
                    node.BeginEdit();
                }
            }
        }

        // Rename folder
        private void trAccounts_AfterLabelEdit(object sender, NodeLabelEditEventArgs e)
        {
            if (e.Label == null)
                return;

            if (e.Label.Length == 0)
            {
                e.CancelEdit = true;
                MessageBox.Show("Invalid folder name.The name cannot be blank");
                return;
            }

            EnableIdle(false);
            try
            {
                TreeNode node = e.Node;
                if (node.Tag != null) // rename folder
                {
                    ConnectServer(node);
                    string cur_localpath = GetFolderByNode(node);
                    oClient.RenameFolder(node.Tag as Imap4Folder, e.Label);
                    e.Node.EndEdit(false);
                    string new_localpath = GetFolderByNode(node);


                    try
                    {
                        // Try to rename local folder as well.
                        System.IO.Directory.Move(cur_localpath, new_localpath);
                    }
                    catch (Exception ex)
                    {
                        
                    }
                }
            }
            catch (Exception ep)
            {
                MessageBox.Show(ep.Message);
                oClient.Close();
                e.CancelEdit = true;
            }

            EnableIdle(true);

        }

        // Refresh mails
        private void refreshMailsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TreeNode node = trAccounts.SelectedNode;
            if (node == null)
                return;

            try
            {

                lstMail.Items.Clear();
                ConnectServer(node);
                oClient.RefreshMailInfos();
                lblStatus.Text = "Refreshing Mails ...";
                
                ShowNode();

            }
            catch (Exception ep)
            {
                MessageBox.Show(ep.Message);
                oClient.Close();
            }
        }

        #endregion

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.Run(new Form1());
        }


        private void Form1_Load(object sender, System.EventArgs e)
        {
            EnableIdle(false);

            lstAuthType.Items.Add("USER/LOGIN");
            lstAuthType.Items.Add("CRAM-MD5");
            lstAuthType.Items.Add("NTLM");
            lstAuthType.SelectedIndex = 0;

            lstProtocol.Items.Add("IMAP4");
            lstProtocol.Items.Add("Exchange Web Service - 2007/2010");
            lstProtocol.Items.Add("Exchange WebDAV - Exchange 2000/2003");
            lstProtocol.SelectedIndex = 0;

            string path = Application.ExecutablePath;
            int pos = path.LastIndexOf('\\');
            if (pos != -1)
                path = path.Substring(0, pos);

            m_curpath = path;

            object empty = System.Reflection.Missing.Value;
            webMail.Navigate("about:blank");

            // Catching the following events is not necessary, 
            // just make the application more user friendly.
            // If you use the object in asp.net/windows service or non-gui application, 
            // You need not to catch the following events.
            // To learn more detail, please refer to the code in EAGetMail EventHandler region
            oClient.OnAuthorized += new MailClient.OnAuthorizedEventHandler(OnAuthorized);
            oClient.OnConnected += new MailClient.OnConnectedEventHandler(OnConnected);
            oClient.OnIdle += new MailClient.OnIdleEventHandler(OnIdle);
            oClient.OnSecuring += new MailClient.OnSecuringEventHandler(OnSecuring);
            oClient.OnReceivingDataStream += new MailClient.OnReceivingDataStreamEventHandler(OnReceivingDataStream);
            oClient.OnSendingDataStream += new MailClient.OnSendingDataStreamEventHandler(OnSendingDataStream);

        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            EnableIdle(false);
            m_bcancel = false;
            string server, user, password;
            server = textServer.Text.Trim();
            user = textUser.Text.Trim();
            password = textPassword.Text.Trim();

            if (server.Length == 0 || user.Length == 0 || password.Length == 0)
            {
                MessageBox.Show("Please input server, user and password.");
                return;
            }

            btnStart.Enabled = false;

            trAccounts.Nodes.Clear();
            lstMail.Items.Clear();

            ServerAuthType authType = ServerAuthType.AuthLogin;
            if (lstAuthType.SelectedIndex == 1)
                authType = ServerAuthType.AuthCRAM5;
            else if (lstAuthType.SelectedIndex == 2)
                authType = ServerAuthType.AuthNTLM;

            ServerProtocol protocol = (ServerProtocol)(lstProtocol.SelectedIndex + 1);

            MailServer oServer = new MailServer(server, user, password,
                chkSSL.Checked, authType, protocol);

            try
            {
                btnCancel.Enabled = true;
                // Enable log file 
                // oClient.LogFileName = "d:\\imap.txt";
                lblStatus.Text = "Connecting ...";
                oClient.Connect(oServer);

                TreeNode node = new TreeNode(String.Format("{0}\\{1}", oServer.Server.ToLower(), oServer.User.ToLower()));
                node.Tag = oServer;
                oCurServer = oServer.Copy();

                TreeNodeCollection nodes = trAccounts.Nodes;
                nodes.Add(node);
                trAccounts.SelectedNode = node;

                ShowNode();
                EnableIdle(true);

            }
            catch (Exception ep)
            {
                btnStart.Enabled = true;
                lblStatus.Text = ep.Message;
                MessageBox.Show(ep.Message);
                btnCancel.Enabled = false;

            }
        }

        private MailServer GetServerByNode(TreeNode node)
        {
            while (node.Parent != null)
                node = node.Parent;

            return node.Tag as MailServer;
        }

        private string GetFolderByNode(TreeNode node)
        {
            if (node.Parent == null)
                return "";

            Imap4Folder fd = node.Tag as Imap4Folder;
            while (node.Parent != null)
                node = node.Parent;

            return String.Format("{0}\\{1}\\{2}", m_curpath, node.Text, fd.LocalPath);
        }

        // Create local folder
        private void _CreateFullFolder(string folder)
        {
            if (Directory.Exists(folder))
                return;

            int pos = 0;
            while ((pos = folder.IndexOf('\\', pos)) != -1)
            {
                if (pos > 2)
                {
                    string s = folder.Substring(0, pos);
                    if (!Directory.Exists(s))
                        Directory.CreateDirectory(s);
                }
                pos++;
            }

            if (!Directory.Exists(folder))
                Directory.CreateDirectory(folder);
        }

        // Connect server
        private void ConnectServer(TreeNode node)
        {
            TreeNode cur_node = node;
            while (node.Parent != null)
                node = node.Parent;

            MailServer server = node.Tag as MailServer;
            bool bReConnect = false;
            if (oCurServer == null || !oClient.Connected)
            {
                bReConnect = true;
            }
            else if (oCurServer != null)
            {
                if (String.Compare(oCurServer.Server, server.Server, true) != 0 ||
                    String.Compare(oCurServer.User, server.User, true) != 0)
                {
                    bReConnect = true;
                }
            }

            if (bReConnect)
            {
                lblStatus.Text = "Connecting ... ";
                oCurServer = server.Copy();
                m_bcancel = false;
                oClient.Connect(server);
                if (cur_node.Parent != null)
                {
                    oClient.SelectFolder(cur_node.Tag as Imap4Folder);
                }
            }
        }


        // Show folder list and email list
        private void ShowNode()
        {
            TreeNode node = trAccounts.SelectedNode;
            if (node == null)
            {
                lstMail.Items.Clear();
                return;
            }

            try
            {
                if (node.Parent == null)
                {
                    // Current node is root node, 
                    // So we get all folders from server and display it
                    lstMail.Items.Clear();
                    ConnectServer(node);
                   
                    lblStatus.Text = "Refreshing Folders ...";
                    Imap4Folder[] fds = oClient.Imap4Folders;
                    ExpendFolders(fds, node);
                    
                    node.ExpandAll();
                    lblStatus.Text = "";
                }
                else
                {
                    Imap4Folder fd = node.Tag as Imap4Folder;
                    if (!fd.NoSelect)
                    {
                        // Display emails list in current folder
                        LoadServerMails(node, fd);
                    }
                    else
                    {
                        lblStatus.Text = "";
                        lstMail.Items.Clear(); // This is a folder without email storage
                    }
                }
            }
            catch (Exception ep)
            {
                oClient.Close();
                  MessageBox.Show(ep.Message);
            }

        }

        private void trAccounts_AfterSelect(object sender, System.Windows.Forms.TreeViewEventArgs e)
        {
            EnableIdle(false);
            ShowNode();
            EnableIdle(true);
        }

        // Clear local file that is not existed on server.
        private void _ClearLocalMails(
            MailInfo[] infos,
            string localfolder)
        {
            string[] files = Directory.GetFiles(localfolder, "*.eml");
            int count = files.Length;
            for (int i = 0; i < count; i++)
            {
                string s = files[i];
                int pos = s.LastIndexOf("\\");
                if (pos != -1)
                    s = s.Substring(pos + 1);

                bool bfind = false;

                UIDLItem item = oUIDLManager.FindLocalFile(s);
                if (item != null)
                {
                    bfind = true;
                }

                if (!bfind)
                {
                    string fileName = files[i];
                    File.Delete(fileName);
                    int p = fileName.LastIndexOf(".");
                    string tempFolder = fileName.Substring(0, p);
                    string htmlName = tempFolder + ".htm";
                    if (File.Exists(htmlName))
                        File.Delete(htmlName);

                    if (Directory.Exists(tempFolder))
                    {
                        Directory.Delete(tempFolder, true);
                    }
                }
            }
        }

        // Get email list from server and diplay it in treeview.
        private void LoadServerMails(TreeNode node, Imap4Folder fd)
        {
            lstMail.Items.Clear();
            lstMail.Sorting = SortOrder.Descending;
            lstMail.ListViewItemSorter = this;

            ConnectServer(node);
            string localfolder = GetFolderByNode(node);
            _CreateFullFolder(localfolder);

            lblStatus.Text = "Refreshing email(s) ...";
            oClient.SelectFolder(fd);
            MailInfo[] infos = oClient.GetMailInfos();

            // UIDL is the identifier of every email on POP3/IMAP4/Exchange server, to avoid retrieve
            // the same email from server more than once, we record the email UIDL retrieved every time
            // UIDLManager wraps the function to write/read uidl record from a text file.
            oUIDLManager = new UIDLManager();

            // Load existed uidl records to UIDLManager
            string uidlfile = String.Format("{0}\\uidl.txt", localfolder);
            oUIDLManager.Load(uidlfile);

            // Remove the local uidl which is not existed on the server.
            oUIDLManager.SyncUIDL(oCurServer, infos);
            oUIDLManager.Update();

            // Remove the email file on local disk that is not existed on server
            _ClearLocalMails( infos, localfolder);

            try
            {
                int count = infos.Length;
                for (int i = 0; i < count; i++)
                {
                    lblStatus.Text = String.Format("Retrieve summary {0}/{1} ...", i + 1, count);
                    MailInfo info = infos[i];

                    // For evaluation usage, please use "TryIt" as the license code, otherwise the 
                    // "Invalid license code" exception will be thrown. However, the object will expire in 1-2 months, then
                    // "Trial version expired" exception will be thrown.
                    Mail oMail = new Mail("TryIt");
                    
                    System.DateTime d = System.DateTime.Now;
                    System.Globalization.CultureInfo cur = new System.Globalization.CultureInfo("en-US");
                    string sdate = d.ToString("yyyyMMddHHmmss", cur);
                    string fileName = String.Format( "{0}{1}{2}.eml",  sdate, d.Millisecond.ToString("d3"), i);
                    string localfile = String.Format("{0}\\{1}", localfolder, fileName);

                    // Detect if current email has been downloaded before.
                    UIDLItem uidl_item = oUIDLManager.FindUIDL(oCurServer, info.UIDL);
                    if (uidl_item != null)
                        localfile = String.Format("{0}\\{1}", localfolder, uidl_item.FileName);

                    if (File.Exists(localfile))
                    {
                        oMail.Load(localfile, true);
                        // This mail has been downloaded from server.
                    }
                    else
                    {
                        // Get the mail header from server.
                        oMail.Load(oClient.GetMailHeader(info));
                        oMail.SaveAs(localfile, true);
                    }

                    if (uidl_item == null)
                    {
                        // Add the email UIDL and local file name to uidl file to avoid we retrieve it again. 
                        oUIDLManager.AddUIDL(oCurServer, info.UIDL, fileName);
                    }

                    ListViewItem item = new ListViewItem(oMail.From.ToString());
                    item.SubItems.Add(oMail.Subject);
                    item.SubItems.Add(oMail.ReceivedDate.ToString("yyyy-MM-dd HH:mm:ss"));
                    item.Tag = info;

                    if (info.Deleted)
                    {
                        item.Font = new System.Drawing.Font(item.Font, FontStyle.Strikeout);
                    }
                    else if (!info.Read)
                    {
                        item.Font = new System.Drawing.Font(item.Font, FontStyle.Bold);
                    }

                    lstMail.Items.Add(item);
                }

                // Update the uidl list to local uidl file and then we can load it next time.
                oUIDLManager.Update();
                lblStatus.Text = String.Format("Total {0} email(s)", count);
            }
            catch (Exception ep)
            {
                // Update the uidl list to local uidl file and then we can load it next time.
                oUIDLManager.Update();
                lblStatus.Text = ep.Message;
                throw ep;
            }
        }

        private void ExpendFolders(Imap4Folder[] fds, TreeNode node)
        {
            node.Nodes.Clear();
            int count = fds.Length;
            for (int i = 0; i < count; i++)
            {
                Imap4Folder fd = fds[i];
                TreeNode subNode = node.Nodes.Add(fd.Name);
                subNode.Tag = fd;
                ExpendFolders(fd.SubFolders, subNode);
            }
        }


        #region IComparer Members

        public int Compare(object x, object y)
        {
            ListViewItem itemx = x as ListViewItem;
            ListViewItem itemy = y as ListViewItem;

            string sx = itemx.SubItems[2].Text;
            string sy = itemy.SubItems[2].Text;
            if (sx.Length != sy.Length)
                return -1; //should never occured.

            int count = sx.Length;
            for (int i = 0; i < count; i++)
            {
                if (sx[i] > sy[i])
                    return -1;
                else if (sx[i] < sy[i])
                    return 1;
            }

            return 0;
        }

        #endregion

#region Email Management
        // Pure deleted email from server, only IMAP4 supports this operation.
        private void btnPure_Click(object sender, System.EventArgs e)
        {
            if (oCurServer.Protocol == ServerProtocol.ExchangeEWS ||
                oCurServer.Protocol == ServerProtocol.ExchangeWebDAV)
            {
                // EWS and WebDAV doesn't support this operating
                return;
            }
            TreeNode node = trAccounts.SelectedNode;
            if (node == null)
                return;

            if (node.Parent == null)
                return;

            EnableIdle(false);
            try
            {
                ConnectServer(node);
                oClient.Expunge();
                LoadServerMails(node, node.Tag as Imap4Folder);
            }
            catch (Exception ep)
            {
                oClient.Close();
                MessageBox.Show(ep.Message);
            }

            EnableIdle(true) ;

        }

        // Delete email from server.
        private void btnDelete_Click(object sender, System.EventArgs e)
        {
            TreeNode node = trAccounts.SelectedNode;
            if (node == null)
                return;

            if (node.Parent == null)
                return;

            ListView.SelectedListViewItemCollection items = lstMail.SelectedItems;
            if (items == null)
                return;

            if (items.Count == 0)
                return;

            int count = items.Count;

            EnableIdle(false);
            try
            {
                
                ArrayList ar = new ArrayList();
                ConnectServer(node);
                for (int i = 0; i < count; i++)
                {
                    MailInfo info = items[i].Tag as MailInfo;
                    oClient.Delete(info);
                    items[i].Font = new System.Drawing.Font(items[i].Font, FontStyle.Strikeout);

                    if (oCurServer.Protocol == ServerProtocol.ExchangeEWS ||
                        oCurServer.Protocol == ServerProtocol.ExchangeWebDAV)
                    {
                        ar.Add(items[i]);
                    }
                }

                count = ar.Count;
                for (int i = 0; i < count; i++)
                {
                    lstMail.Items.Remove(ar[i] as ListViewItem);
                }

            }
            catch (Exception ep)
            {
                oClient.Close();
                MessageBox.Show(ep.Message);
            }

            EnableIdle(true);
        }

        // Undelete  email marked as deleted from server, only IMAP4 supports this operation.
        private void btnUndelete_Click(object sender, System.EventArgs e)
        {
            if (oCurServer.Protocol == ServerProtocol.ExchangeEWS ||
                oCurServer.Protocol == ServerProtocol.ExchangeWebDAV)
            {
                // EWS and WebDAV doesn't support this operating
                return;
            }

            TreeNode node = trAccounts.SelectedNode;
            if (node == null)
                return;

            if (node.Parent == null)
                return;

            ListView.SelectedListViewItemCollection items = lstMail.SelectedItems;
            if (items == null)
                return;

            if (items.Count == 0)
                return;

            int count = items.Count;

            EnableIdle(false);
            try
            {
                ConnectServer(node);
                for (int i = 0; i < count; i++)
                {
                    ListViewItem item = items[i];
                    MailInfo info = item.Tag as MailInfo;
                    oClient.Undelete(info);
                    if (!info.Read)
                    {
                        item.Font = new System.Drawing.Font(item.Font, FontStyle.Bold);
                    }
                    else
                    {
                        item.Font = new System.Drawing.Font(item.Font, FontStyle.Regular);
                    }
                }
            }
            catch (Exception ep)
            {
                oClient.Close();
                MessageBox.Show(ep.Message);
            }

            EnableIdle(true);

        }

        // Mark email as unread.
        private void btnUnread_Click(object sender, System.EventArgs e)
        {
            TreeNode node = trAccounts.SelectedNode;
            if (node == null)
                return;

            if (node.Parent == null)
                return;

            ListView.SelectedListViewItemCollection items = lstMail.SelectedItems;
            if (items == null)
                return;

            if (items.Count == 0)
                return;

            int count = items.Count;

            EnableIdle(false);
            try
            {
                ConnectServer(node);
                for (int i = 0; i < count; i++)
                {
                    ListViewItem item = items[i];
                    MailInfo info = item.Tag as MailInfo;
                    oClient.MarkAsRead(info, false);
                   
                    if (info.Deleted )
                    {
                        item.Font = new System.Drawing.Font(item.Font, FontStyle.Strikeout);
                    }
                    else
                    {
                        item.Font = new System.Drawing.Font(item.Font, FontStyle.Bold);
                    }
                }
            }
            catch (Exception ep)
            {
                oClient.Close();
                MessageBox.Show(ep.Message);
            }

            EnableIdle(true);
        }

        // Move email to another folder.
        private void btnMove_Click(object sender, System.EventArgs e)
        {
            TreeNode node = trAccounts.SelectedNode;
            if (node == null)
                return;

            if (node.Parent == null)
                return;

            ListView.SelectedListViewItemCollection items = lstMail.SelectedItems;
            if (items == null)
                return;

            if (items.Count == 0)
                return;

            int count = items.Count;

            EnableIdle(false);
            try
            {

                ConnectServer(node);
                Imap4Folder[] fds = oClient.Imap4Folders;

                Form4 frm = new Form4();
                TreeNode fnode = frm.trFolders.Nodes.Add("Root Folder");
                ExpendFolders(fds, fnode);
                fnode.ExpandAll();

                if (frm.ShowDialog() != DialogResult.OK)
                {
                    EnableIdle(true);
                    return;
                }

                Imap4Folder curfd = node.Tag as Imap4Folder;
                Imap4Folder fd = frm.trFolders.SelectedNode.Tag as Imap4Folder;
                if (String.Compare(curfd.FullPath, fd.FullPath, true) == 0)
                {
                    EnableIdle(true);
                    return;
                }

                ArrayList ar = new ArrayList();
                for (int i = 0; i < count; i++)
                {
                    ListViewItem item = items[i];
                    MailInfo info = item.Tag as MailInfo;
                    oClient.Copy(info, fd);
                    oClient.Delete(info);
                    item.Font = new System.Drawing.Font(item.Font, FontStyle.Strikeout);
                    
                    if (oCurServer.Protocol == ServerProtocol.ExchangeEWS ||
                             oCurServer.Protocol == ServerProtocol.ExchangeWebDAV)
                    {
                        ar.Add(items[i]);
                    }
                }

                count = ar.Count;
                for (int i = 0; i < count; i++)
                {
                    lstMail.Items.Remove(ar[i] as ListViewItem);
                }
            }
            catch (Exception ep)
            {
                oClient.Close();
                MessageBox.Show(ep.Message);
            }

            EnableIdle(true);
        }

        // Download and display current email.
        private void lstMail_Click(object sender, System.EventArgs e)
        {

            ListView.SelectedListViewItemCollection items = lstMail.SelectedItems;
            if (items == null)
                return;

            if (items.Count == 0)
                return;

            TreeNode node = trAccounts.SelectedNode;
            if (node == null)
                return;

            EnableIdle(true);
            if (items.Count > 1)
            {
                return;
            }

            ListViewItem item = items[0];
            MailInfo info = item.Tag as MailInfo;
            Imap4Folder fd = node.Tag as Imap4Folder;

            try
            {
                EnableIdle(false);
                ConnectServer(node);

                string mailbox = GetFolderByNode(node);
                _CreateFullFolder(mailbox);

                // Find current email record in UIDL file.
                UIDLItem oUIDL = oUIDLManager.FindUIDL(oCurServer, info.UIDL);
                if (oUIDL == null)
                {
                    // show never happen except you delete the file from the folder manually.
                    throw new Exception("No email file found!");
                }

                // Get the  local file name for this email UIDL
                string emlFile = String.Format("{0}\\{1}", mailbox, oUIDL.FileName);

                int pos = emlFile.LastIndexOf(".");
                string mainName = emlFile.Substring(0, pos);
                string htmlName = mainName + ".htm";

                // only mail header is retrieved, now retrieve full content of mail.
                if (!File.Exists(htmlName))
                {
                    Mail oMail = oClient.GetMail(info);
                    oMail.SaveAs(emlFile, true);
                    pgBar.Minimum = 0;
                    pgBar.Maximum = 100;
                    pgBar.Value = 0;
                }

                if (!info.Read)
                {
                    oClient.MarkAsRead(info, true);
                    if (info.Deleted)
                    {
                        item.Font = new System.Drawing.Font(item.Font, FontStyle.Strikeout);
                    }
                    else
                    {
                        item.Font = new System.Drawing.Font(item.Font, FontStyle.Regular);
                    }
                }

                ShowMail(emlFile);
            }
            catch (Exception ep)
            {
                oClient.Close();
                MessageBox.Show(ep.Message);
            }

            EnableIdle(true);
        }

        // Copy emails from one folder to another.
        private void btnCopy_Click(object sender, EventArgs e)
        {
            TreeNode node = trAccounts.SelectedNode;
            if (node == null)
                return;

            if (node.Parent == null)
                return;

            ListView.SelectedListViewItemCollection items = lstMail.SelectedItems;
            if (items == null)
                return;

            if (items.Count == 0)
                return;

            int count = items.Count;

            EnableIdle(false);
            try
            {
                ConnectServer(node);
                Imap4Folder[] fds = oClient.Imap4Folders;

                Form4 frm = new Form4();
                TreeNode fnode = frm.trFolders.Nodes.Add("Root Folder");
                ExpendFolders(fds, fnode);
                fnode.ExpandAll();

                if (frm.ShowDialog() != DialogResult.OK)
                {
                    EnableIdle(true);
                    return;
                }

                Imap4Folder curfd = node.Tag as Imap4Folder;
                Imap4Folder fd = frm.trFolders.SelectedNode.Tag as Imap4Folder;
                if (String.Compare(curfd.FullPath, fd.FullPath, true) == 0)
                {
                    EnableIdle(true);
                    return;
                }

                for (int i = 0; i < count; i++)
                {
                    ListViewItem item = items[i];
                    MailInfo info = item.Tag as MailInfo;
                    oClient.Copy(info, fd);
                }
            }
            catch (Exception ep)
            {
                oClient.Close();
                MessageBox.Show(ep.Message);
            }

            EnableIdle(true);
        }

        // Upload EML file to specified folder
        private void btnUpload_Click(object sender, EventArgs e)
        {
            TreeNode node = trAccounts.SelectedNode;
            if (node == null)
                return;

            if (node.Parent == null)
                return;

            openFileDialog1.Multiselect = false;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                EnableIdle(false);
                try
                {
                    Mail oMail = new Mail("TryIt");
                    oMail.Load(openFileDialog1.FileName, false);
                    ConnectServer(node);

                    oClient.Append(node.Tag as Imap4Folder, oMail.Content);
                    oClient.RefreshMailInfos();
                    ShowNode();
                }
                catch (Exception ep)
                {
                    oClient.Close();
                    MessageBox.Show(ep.Message);
                }

                EnableIdle(true);
            }

            pgBar.Minimum = 0;
            pgBar.Maximum = 100;
            pgBar.Value = 0;
        }
#endregion

        // quit and close current connection.
        private void btnQuit_Click(object sender, EventArgs e)
        {
            try
            {
                btnStart.Enabled = true;
                trAccounts.Nodes.Clear();
                lstMail.Items.Clear();
                webMail.Navigate("about:blank");
                oClient.Logout();
                oClient.Close();
                lblStatus.Text = "Discconnected";
                trAccounts.SelectedNode = null;

                Application.DoEvents();
            }
            catch (Exception ep)
            { }

            EnableIdle(false);


        }

        // Terminate current operation
        private void btnCancel_Click(object sender, EventArgs e)
        {
            m_bcancel = true;
        }

        // Enable/Disable control
        private void EnableIdle(bool bIdle)
        {
            btnDelete.Enabled = bIdle;
            btnUndelete.Enabled = bIdle;
            btnUnread.Enabled = bIdle;
            btnPure.Enabled = bIdle;
            btnMove.Enabled = bIdle;
            btnCopy.Enabled = bIdle;
            btnUpload.Enabled = bIdle;

            if (lstMail.SelectedItems.Count == 0)
            {
                btnDelete.Enabled = false;
                btnUndelete.Enabled = false;
                btnUnread.Enabled = false;
                btnMove.Enabled = false;
                btnCopy.Enabled = false;
            }

            if (trAccounts.SelectedNode == null)
            {
                btnUpload.Enabled = false;
            }
            else
            {
                if (trAccounts.SelectedNode.Parent == null)
                {
                    btnUpload.Enabled = false;
                }
            
            }

            btnCancel.Enabled = !bIdle;
            if (btnStart.Enabled)
                btnCancel.Enabled = false;

            btnQuit.Enabled = bIdle;
            contextMenuFolder.Enabled = bIdle;
            trAccounts.Enabled = bIdle;
            lstMail.Enabled = bIdle;

            if (oCurServer != null)
            {
                if (oCurServer.Protocol == ServerProtocol.ExchangeWebDAV ||
                    oCurServer.Protocol == ServerProtocol.ExchangeEWS)
                {// Exchange WebDAV and EWS doesn't support this operating
                    btnUndelete.Enabled = false;
                    btnPure.Enabled = false;
                }
            }
        }

        // Resize control automatically
        private void Form1_Resize(object sender, EventArgs e)
        {
            if (this.Width < 900)
                this.Width = 900;

            if (this.Height < 470)
                this.Height = 470;

            lstMail.Width = this.Width - 455;
            webMail.Width = lstMail.Width;
            pgBar.Width = webMail.Width;
            trAccounts.Height = this.Height - 60;
            webMail.Height = this.Height - lstMail.Height - 120;
            pgBar.Top = this.Height - 60;
        }

        private void lstProtocol_SelectedIndexChanged(object sender, EventArgs e)
        {
            // By default, Exchange Web Service requires SSL connection.
            // For other protocol, please set SSL connection based on your server setting manually

            if ((lstProtocol.SelectedIndex + 1) == (int)ServerProtocol.ExchangeEWS)
            {
                chkSSL.Checked = true;
            }
        }

    }
}
