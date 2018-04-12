//  ===============================================================================
// |    THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF      |
// |    ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO    |
// |    THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A         |
// |    PARTICULAR PURPOSE.                                                    |
// |    Copyright (c)2006-2012  ADMINSYSTEM SOFTWARE LIMITED                         |
// |
// |    Project: It demonstrates how to use EAGetMail to receive/parse non-delivery report.
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
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using EAGetMail;

namespace pop3.csharp
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class Form1 : System.Windows.Forms.Form, IComparer
	{
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.TextBox textServer;
		private System.Windows.Forms.TextBox textUser;
		private System.Windows.Forms.TextBox textPassword;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.ComboBox lstAuthType;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.ComboBox lstProtocol;
		private System.Windows.Forms.CheckBox chkLeaveCopy;
		private System.Windows.Forms.Button btnStart;
		private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.Label lblStatus;
		private System.Windows.Forms.ProgressBar pgBar;
		private System.Windows.Forms.CheckBox chkSSL;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		private bool m_bcancel = false;
		private string m_uidlfile = "uidl.txt";
		private string m_curpath = "";
		
		private System.Windows.Forms.ListView lstMail;
		private System.Windows.Forms.ColumnHeader colFrom;
		private System.Windows.Forms.ColumnHeader colSubject;
        private System.Windows.Forms.ColumnHeader colDate;
		private System.Windows.Forms.Button btnDel;
		private System.Windows.Forms.Label lblTotal;
        private System.Windows.Forms.Label label6;
        private TextBox textReport;

		private ArrayList m_arUidl = new ArrayList();
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
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.pgBar = new System.Windows.Forms.ProgressBar();
            this.lblStatus = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnStart = new System.Windows.Forms.Button();
            this.chkLeaveCopy = new System.Windows.Forms.CheckBox();
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
            this.label1 = new System.Windows.Forms.Label();
            this.lstMail = new System.Windows.Forms.ListView();
            this.colFrom = new System.Windows.Forms.ColumnHeader();
            this.colSubject = new System.Windows.Forms.ColumnHeader();
            this.colDate = new System.Windows.Forms.ColumnHeader();
            this.btnDel = new System.Windows.Forms.Button();
            this.lblTotal = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.textReport = new System.Windows.Forms.TextBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.pgBar);
            this.groupBox1.Controls.Add(this.lblStatus);
            this.groupBox1.Controls.Add(this.btnCancel);
            this.groupBox1.Controls.Add(this.btnStart);
            this.groupBox1.Controls.Add(this.chkLeaveCopy);
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
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(8, 8);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(232, 328);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Account Information";
            // 
            // pgBar
            // 
            this.pgBar.Location = new System.Drawing.Point(8, 312);
            this.pgBar.Name = "pgBar";
            this.pgBar.Size = new System.Drawing.Size(216, 8);
            this.pgBar.TabIndex = 15;
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(8, 272);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(0, 13);
            this.lblStatus.TabIndex = 14;
            // 
            // btnCancel
            // 
            this.btnCancel.Enabled = false;
            this.btnCancel.Location = new System.Drawing.Point(128, 240);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(88, 24);
            this.btnCancel.TabIndex = 13;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnStart
            // 
            this.btnStart.Location = new System.Drawing.Point(32, 240);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(88, 24);
            this.btnStart.TabIndex = 12;
            this.btnStart.Text = "Start";
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // chkLeaveCopy
            // 
            this.chkLeaveCopy.Location = new System.Drawing.Point(8, 208);
            this.chkLeaveCopy.Name = "chkLeaveCopy";
            this.chkLeaveCopy.Size = new System.Drawing.Size(208, 16);
            this.chkLeaveCopy.TabIndex = 11;
            this.chkLeaveCopy.Text = "Leave a copy of message on server";
            // 
            // lstProtocol
            // 
            this.lstProtocol.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.lstProtocol.Location = new System.Drawing.Point(80, 176);
            this.lstProtocol.Name = "lstProtocol";
            this.lstProtocol.Size = new System.Drawing.Size(136, 21);
            this.lstProtocol.TabIndex = 10;
            this.lstProtocol.SelectedIndexChanged += new System.EventHandler(this.lstProtocol_SelectedIndexChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(8, 178);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(46, 13);
            this.label5.TabIndex = 9;
            this.label5.Text = "Protocol";
            // 
            // lstAuthType
            // 
            this.lstAuthType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.lstAuthType.Location = new System.Drawing.Point(80, 144);
            this.lstAuthType.Name = "lstAuthType";
            this.lstAuthType.Size = new System.Drawing.Size(136, 21);
            this.lstAuthType.TabIndex = 8;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(8, 146);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(56, 13);
            this.label4.TabIndex = 7;
            this.label4.Text = "Auth Type";
            // 
            // chkSSL
            // 
            this.chkSSL.Location = new System.Drawing.Point(8, 120);
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
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(8, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(38, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Server";
            // 
            // lstMail
            // 
            this.lstMail.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.colFrom,
            this.colSubject,
            this.colDate});
            this.lstMail.FullRowSelect = true;
            this.lstMail.HideSelection = false;
            this.lstMail.Location = new System.Drawing.Point(248, 16);
            this.lstMail.Name = "lstMail";
            this.lstMail.Size = new System.Drawing.Size(474, 168);
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
            this.colDate.Width = 150;
            // 
            // btnDel
            // 
            this.btnDel.Location = new System.Drawing.Point(650, 186);
            this.btnDel.Name = "btnDel";
            this.btnDel.Size = new System.Drawing.Size(72, 24);
            this.btnDel.TabIndex = 3;
            this.btnDel.Text = "Delete";
            this.btnDel.Click += new System.EventHandler(this.btnDel_Click);
            // 
            // lblTotal
            // 
            this.lblTotal.AutoSize = true;
            this.lblTotal.Location = new System.Drawing.Point(256, 192);
            this.lblTotal.Name = "lblTotal";
            this.lblTotal.Size = new System.Drawing.Size(0, 13);
            this.lblTotal.TabIndex = 4;
            // 
            // label6
            // 
            this.label6.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label6.Location = new System.Drawing.Point(16, 352);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(216, 48);
            this.label6.TabIndex = 5;
            this.label6.Text = "Warning: if \"leave a copy of message on server\" is not checked,  the emails on th" +
                "e server will be deleted !";
            // 
            // textReport
            // 
            this.textReport.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textReport.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.textReport.Location = new System.Drawing.Point(248, 215);
            this.textReport.Multiline = true;
            this.textReport.Name = "textReport";
            this.textReport.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textReport.Size = new System.Drawing.Size(473, 184);
            this.textReport.TabIndex = 6;
            // 
            // Form1
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(734, 412);
            this.Controls.Add(this.textReport);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.lblTotal);
            this.Controls.Add(this.btnDel);
            this.Controls.Add(this.lstMail);
            this.Controls.Add(this.groupBox1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.Resize += new System.EventHandler(this.Form1_Resize);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

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

		#region EAGetMail Event Handler
		public void OnConnected(object sender, ref bool cancel )
		{
			lblStatus.Text = "Connected ...";
			cancel = m_bcancel;
			Application.DoEvents();
		}

		public void OnQuit(object sender, ref bool cancel )
		{
			lblStatus.Text = "Quit ...";
			cancel = m_bcancel;
			Application.DoEvents();
		}

		public void OnReceivingDataStream(object sender, MailInfo info, int received, int total, ref bool cancel )
		{
			pgBar.Minimum =  0;
			pgBar.Maximum = total;
			pgBar.Value = received;
			cancel = m_bcancel;
			Application.DoEvents();
		}

		public void OnIdle( object sender, ref bool cancel )
		{
			cancel = m_bcancel;
			Application.DoEvents();
		}

		public void OnAuthorized( object sender, ref bool cancel )
		{
			lblStatus.Text =  "Authorized ...";
			cancel = m_bcancel;
			Application.DoEvents();
		}

		public void OnSecuring( object sender, ref bool cancel )
		{
			lblStatus.Text = "Securing ..." ;
			cancel = m_bcancel;
			Application.DoEvents();
		}
		#endregion

		#region Parse and Display Report
		private void LoadMails()
		{
			lstMail.Items.Clear();
			string mailFolder = String.Format( "{0}\\inbox", m_curpath );
			if( !Directory.Exists( mailFolder ))
				Directory.CreateDirectory( mailFolder );

			string [] files = Directory.GetFiles( mailFolder, "*.eml" );
			int count = files.Length;
			for( int i = 0; i < count; i++ )
			{
				string fullname = files[i];
				//For evaluation usage, please use "TryIt" as the license code, otherwise the 
				//"invalid license code" exception will be thrown. However, the object will expire in 1-2 months, then
				//"trial version expired" exception will be thrown.
				Mail oMail = new Mail( "TryIt" );

				// Load( file, true ) only load the email header to Mail object to save the CPU and memory
				// the Mail object will load the whole email file later automatically if bodytext or attachment is required..
				oMail.Load( fullname, true );

				ListViewItem item = new ListViewItem(oMail.From.ToString());
				item.SubItems.Add( oMail.Subject );
				item.SubItems.Add( oMail.ReceivedDate.ToString( "yyyy-MM-dd HH:mm:ss"));
				item.Tag = fullname;
				lstMail.Items.Add(item);
				oMail.Clear();
			}
		}

        // Report non-delivery report
        private void ShowReport(string fileName)
        {
            try
            {
                Mail oMail = new Mail("TryIt");
                oMail.Load(fileName, false);

                if( !oMail.IsReport )
                {
                    MessageBox.Show( "This is not a report" );
                    return;
                }
      
                string  s = "";
                MailReport oReport = oMail.GetReport();
                switch( oReport.ReportType )
                {
                    case DeliveryReportType.DeliveryReceipt:
                       s = "This is a deliver receipt.\r\n\r\n";
                       break;
                    case DeliveryReportType.ReadReceipt:
                        s = "This is a read receipt.\r\n\r\n";
                        break;
                    default:
                        s = "This is a failure report.\r\n\r\n";
                        break;
                }
                
                s += String.Format( "Original Sender: {0}\r\n", oReport.OriginalSender );
                s += String.Format( "Original Recipient: {0}\r\n", oReport.OriginalRecipient );
                s += String.Format( "Original Message-ID: {0}\r\n\r\n", oReport.OriginalMessageID );

                if( oReport.ReportType == DeliveryReportType.FailureReport )
                {
                    s += String.Format( "Original Subject: {0}\r\n", oReport.OriginalSubject );
                    s += String.Format( "Report MTA: {0}\r\n", oReport.ReportMTA );
                    s += String.Format( "Error Code: {0}\r\n", oReport.ErrCode );
                    s += String.Format( "Error Description: {0}\r\n\r\n", oReport.ErrDescription );

                    s += "---- Original Message Header ----\r\n\r\n";
                    HeaderCollection oHeaders = oReport.OriginalHeaders;
                    int count = oHeaders.Count;
                    for( int i = 0; i < count; i++ )
                    {
                        HeaderItem oHeader = oHeaders[i] as HeaderItem;
                        s += String.Format( "{0}: {1}\r\n", oHeader.HeaderKey, oHeader.HeaderValue );
                    }
                }

                textReport.Text = s;

            }
            catch (Exception ep)
            {
                MessageBox.Show(ep.Message);
            }
        }
	
		#endregion

		private void btnStart_Click(object sender, System.EventArgs e)
		{
			string server, user, password;
			server = textServer.Text.Trim();
			user = textUser.Text.Trim();
			password = textPassword.Text.Trim();

			if( server.Length == 0 || user.Length == 0 || password.Length == 0 )
			{
				MessageBox.Show("Please input server, user and password." );
				return;
			}

			btnStart.Enabled = false;
			btnCancel.Enabled = true;

			ServerAuthType authType = ServerAuthType.AuthLogin;
			if( lstAuthType.SelectedIndex == 1 )
				authType = ServerAuthType.AuthCRAM5;
			else if( lstAuthType.SelectedIndex == 2 )
				authType = ServerAuthType.AuthNTLM;

            ServerProtocol protocol = (ServerProtocol)lstProtocol.SelectedIndex;

			MailServer oServer = new MailServer( server, user, password,
				chkSSL.Checked, authType, protocol );

            //For evaluation usage, please use "TryIt" as the license code, otherwise the 
            //"invalid license code" exception will be thrown. However, the object will expire in 1-2 months, then
            //"trial version expired" exception will be thrown.
			MailClient oClient = new MailClient("TryIt");

			//Catching the following events is not necessary, 
            //just make the application more user friendly.
            //If you use the object in asp.net/windows service or non-gui application, 
            //You need not to catch the following events.
            //To learn more detail, please refer to the code in EAGetMail EventHandler region
			oClient.OnAuthorized += new MailClient.OnAuthorizedEventHandler( OnAuthorized );
			oClient.OnConnected += new MailClient.OnConnectedEventHandler( OnConnected );
			oClient.OnIdle += new MailClient.OnIdleEventHandler( OnIdle );
			oClient.OnSecuring += new MailClient.OnSecuringEventHandler( OnSecuring );
			oClient.OnReceivingDataStream += new MailClient.OnReceivingDataStreamEventHandler( OnReceivingDataStream );	

			bool bLeaveCopy = chkLeaveCopy.Checked;

            // UIDL is the identifier of every email on POP3/IMAP4/Exchange server, to avoid retrieve
            // the same email from server more than once, we record the email UIDL retrieved every time
            // if you delete the email from server every time and not to leave a copy of email on
            // the server, then please remove all the function about uidl.
            // UIDLManager wraps the function to write/read uidl record from a text file.
            UIDLManager oUIDLManager  = new UIDLManager();

			try
			{
			    // load existed uidl records to UIDLManager
                string uidlfile = String.Format("{0}\\{1}", m_curpath, m_uidlfile);
                oUIDLManager.Load(uidlfile);

				string mailFolder = String.Format( "{0}\\inbox", m_curpath );
				if( !Directory.Exists( mailFolder ))
					Directory.CreateDirectory( mailFolder );

				m_bcancel = false;
                lblStatus.Text = "Connecting ...";
				oClient.Connect( oServer );
				MailInfo[] infos = oClient.GetMailInfos();
				lblStatus.Text = String.Format( "Total {0} email(s)" , infos.Length );
				
			    // remove the local uidl which is not existed on the server.
                oUIDLManager.SyncUIDL(oServer, infos);
                oUIDLManager.Update();

				int count = infos.Length;

                ArrayList arReport = new ArrayList();
				for( int i = 0; i < count; i++ )
				{
					MailInfo info = infos[i];
					if(oUIDLManager.FindUIDL( oServer, info.UIDL ) != null )
					{
						// This email has been downloaded or checked before.
						continue;
					}

					lblStatus.Text = String.Format( "Retrieving mail header {0}/{1}...", info.Index, count );

                    System.DateTime d = System.DateTime.Now;
                    System.Globalization.CultureInfo cur = new System.Globalization.CultureInfo("en-US");
                    string sdate = d.ToString("yyyyMMddHHmmss", cur);
                    string fileName = String.Format("{0}\\{1}{2}{3}.eml", mailFolder, sdate, d.Millisecond.ToString("d3"), i);

                    Mail oMail = new Mail("TryIt");
                    oMail.Load(oClient.GetMailHeader(info));

                    if (!oMail.IsReport)
                    {
                        // Not a report, continue
                        // Add the email uidl to uidl file to avoid we check it next time. 
                        oUIDLManager.AddUIDL(oServer, info.UIDL, fileName);
                        continue;
                    }

                    // This is a report, get the entire email.
                    oMail = oClient.GetMail(info);
					oMail.SaveAs( fileName, true );

					ListViewItem item = new ListViewItem(oMail.From.ToString());
					item.SubItems.Add( oMail.Subject );
					item.SubItems.Add( oMail.ReceivedDate.ToString( "yyyy-MM-dd HH:mm:ss"));
					item.Font = new System.Drawing.Font(item.Font, FontStyle.Bold);
					item.Tag = fileName;
					lstMail.Items.Insert(0, item);
					oMail.Clear();

					lblTotal.Text = String.Format( "Total {0} email(s)", lstMail.Items.Count );

                    arReport.Add(info); // Add the report mail info to arraylist, 
                    // then we can delete it later.
					if( bLeaveCopy )
					{
						// Add the email uidl to uidl file to avoid we retrieve it next time. 
						oUIDLManager.AddUIDL( oServer, info.UIDL, fileName );
					}
				}

				if( !bLeaveCopy )
				{
					lblStatus.Text = "Deleting ...";
                    count = arReport.Count;
					for( int i = 0; i < count; i++ )
                    {
                        MailInfo info = arReport[i] as MailInfo;
                        oClient.Delete(info);
                        // Remove UIDL from local uidl file.
                        oUIDLManager.RemoveUIDL( oServer, info.UIDL );
                    }
				}
				// Delete method just mark the email as deleted, 
				// Quit method pure the emails from server exactly.
				oClient.Quit();
			
			}
			catch( Exception ep )
			{
				MessageBox.Show( ep.Message );
			}
				
			// Update the uidl list to local uidl file and then we can load it next time.
			oUIDLManager.Update();

			lblStatus.Text = "Completed";
			pgBar.Maximum = 100;
			pgBar.Minimum = 0;
			pgBar.Value = 0;
			btnStart.Enabled = true;
			btnCancel.Enabled = false;
		}

		private void Form1_Load(object sender, System.EventArgs e)
		{
			
			lstProtocol.Items.Add( "POP3" );
			lstProtocol.Items.Add( "IMAP4" );
            lstProtocol.Items.Add("Exchange Web Service - 2007/2010");
            lstProtocol.Items.Add("Exchange WebDAV - Exchange 2000/2003");
 
			lstProtocol.SelectedIndex = 0;

			lstAuthType.Items.Add( "USER/LOGIN" );
			lstAuthType.Items.Add( "APOP" );
			lstAuthType.Items.Add( "NTLM" );
			lstAuthType.SelectedIndex = 0;

			string path = Application.ExecutablePath;
			int pos = path.LastIndexOf( '\\' );
			if( pos != -1 )
				path = path.Substring( 0, pos );

			m_curpath = path;

			lstMail.Sorting = SortOrder.Descending;
			lstMail.ListViewItemSorter = this;

			LoadMails();
			lblTotal.Text = String.Format( "Total {0} report(s)", lstMail.Items.Count );
		}

		private void btnCancel_Click(object sender, System.EventArgs e)
		{
			m_bcancel = true;
		}

		private void lstMail_Click(object sender, System.EventArgs e)
		{
			ListView.SelectedListViewItemCollection items = lstMail.SelectedItems;
			if( items.Count == 0 )
				return;

			ListViewItem item = items[0] as ListViewItem;
			ShowReport( item.Tag as string );
			item.Font = new System.Drawing.Font(item.Font, FontStyle.Regular);
		}

		private void btnDel_Click(object sender, System.EventArgs e)
		{
			ListView.SelectedListViewItemCollection items = lstMail.SelectedItems;
			if( items.Count == 0 )
				return;

			 if( MessageBox.Show("Do you want to delete all selected emails", 
                             "", 
                             MessageBoxButtons.YesNo) == DialogResult.No )
				 return;

			while( items.Count > 0 )
			{
				try
				{
					string fileName = items[0].Tag as string;
					File.Delete( fileName );
					int pos = fileName.LastIndexOf(".");
					string tempFolder = fileName.Substring(0, pos);
					string htmlName = tempFolder + ".htm";
					if( File.Exists( htmlName ))
						File.Delete( htmlName );
					
					if( Directory.Exists( tempFolder ))
					{
						Directory.Delete( tempFolder, true );
					}

					lstMail.Items.Remove(items[0]);
				}
				catch(Exception ep )
				{
					MessageBox.Show( ep.Message );
					break;
				}
			}

			lblTotal.Text = String.Format( "Total {0} report(s)", lstMail.Items.Count );
		}

		#region IComparer Members
		// sort the email as received data.
		public int Compare(object x, object y)
		{
			ListViewItem itemx = x as ListViewItem;
			ListViewItem itemy = y as ListViewItem;

			string sx = itemx.SubItems[2].Text;
			string sy = itemy.SubItems[2].Text;
			if( sx.Length != sy.Length )
				return -1; //should never occured.

			int count = sx.Length;
			for( int i = 0; i < count; i++ )
			{
				if( sx[i] > sy[i] )
					return -1;
				else if( sx[i] < sy[i] )
					return 1;
			}

			return 0;
		}

		#endregion

        private void Form1_Resize(object sender, EventArgs e)
        {
            if (this.Width < 750)
            {
                this.Width = 750;
            }

            if (this.Height < 450)
            {
                this.Height = 450;
            }

            lstMail.Width = this.Width - 270;
            textReport.Width = lstMail.Width;
            btnDel.Left = this.Width - (btnDel.Width + 20);
            textReport.Height = this.Height - (lstMail.Height + 100);
        }

        private void lstProtocol_SelectedIndexChanged(object sender, EventArgs e)
        {
            // By default, Exchange Web Service requires SSL connection.
            // For other protocol, please set SSL connection based on your server setting manually
            if (lstProtocol.SelectedIndex == (int)ServerProtocol.ExchangeEWS)
            {
                chkSSL.Checked = true;
            }
         
        }



	}
}
