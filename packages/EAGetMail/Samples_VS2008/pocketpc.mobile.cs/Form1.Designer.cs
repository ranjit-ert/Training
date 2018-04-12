namespace pocketpc.mobile.cs
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.MainMenu mainMenu1;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
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
            this.mainMenu1 = new System.Windows.Forms.MainMenu();
            this.textServer = new System.Windows.Forms.TextBox();
            this.textUser = new System.Windows.Forms.TextBox();
            this.textPassword = new System.Windows.Forms.TextBox();
            this.chkSSL = new System.Windows.Forms.CheckBox();
            this.lstAuthType = new System.Windows.Forms.ComboBox();
            this.lstProtocol = new System.Windows.Forms.ComboBox();
            this.chkLeaveCopy = new System.Windows.Forms.CheckBox();
            this.btnStart = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.pgBar = new System.Windows.Forms.ProgressBar();
            this.Label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.lblStatus = new System.Windows.Forms.Label();
            this.lstMail = new System.Windows.Forms.ListView();
            this.btnDel = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.colFrom = new System.Windows.Forms.ColumnHeader();
            this.colSubject = new System.Windows.Forms.ColumnHeader();
            this.colDate = new System.Windows.Forms.ColumnHeader();
            this.webMail = new System.Windows.Forms.WebBrowser();
            this.SuspendLayout();
            // 
            // textServer
            // 
            this.textServer.Location = new System.Drawing.Point(71, 14);
            this.textServer.Name = "textServer";
            this.textServer.Size = new System.Drawing.Size(103, 23);
            this.textServer.TabIndex = 0;
            // 
            // textUser
            // 
            this.textUser.Location = new System.Drawing.Point(71, 55);
            this.textUser.Name = "textUser";
            this.textUser.Size = new System.Drawing.Size(103, 23);
            this.textUser.TabIndex = 1;
            // 
            // textPassword
            // 
            this.textPassword.Location = new System.Drawing.Point(71, 95);
            this.textPassword.Name = "textPassword";
            this.textPassword.PasswordChar = '*';
            this.textPassword.Size = new System.Drawing.Size(103, 23);
            this.textPassword.TabIndex = 2;
            // 
            // chkSSL
            // 
            this.chkSSL.Location = new System.Drawing.Point(71, 137);
            this.chkSSL.Name = "chkSSL";
            this.chkSSL.Size = new System.Drawing.Size(128, 20);
            this.chkSSL.TabIndex = 3;
            this.chkSSL.Text = "SSL Connection";
            // 
            // lstAuthType
            // 
            this.lstAuthType.Location = new System.Drawing.Point(71, 163);
            this.lstAuthType.Name = "lstAuthType";
            this.lstAuthType.Size = new System.Drawing.Size(103, 23);
            this.lstAuthType.TabIndex = 4;
            // 
            // lstProtocol
            // 
            this.lstProtocol.Location = new System.Drawing.Point(71, 205);
            this.lstProtocol.Name = "lstProtocol";
            this.lstProtocol.Size = new System.Drawing.Size(103, 23);
            this.lstProtocol.TabIndex = 5;
            // 
            // chkLeaveCopy
            // 
            this.chkLeaveCopy.Location = new System.Drawing.Point(71, 247);
            this.chkLeaveCopy.Name = "chkLeaveCopy";
            this.chkLeaveCopy.Size = new System.Drawing.Size(248, 20);
            this.chkLeaveCopy.TabIndex = 6;
            this.chkLeaveCopy.Text = "Leave a copy of message on server";
            // 
            // btnStart
            // 
            this.btnStart.Location = new System.Drawing.Point(71, 284);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(75, 20);
            this.btnStart.TabIndex = 7;
            this.btnStart.Text = "Start";
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(160, 284);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 20);
            this.btnCancel.TabIndex = 8;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // pgBar
            // 
            this.pgBar.Location = new System.Drawing.Point(71, 340);
            this.pgBar.Name = "pgBar";
            this.pgBar.Size = new System.Drawing.Size(167, 10);
            // 
            // Label1
            // 
            this.Label1.Location = new System.Drawing.Point(3, 14);
            this.Label1.Name = "Label1";
            this.Label1.Size = new System.Drawing.Size(62, 20);
            this.Label1.Text = "Server";
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(3, 58);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(62, 20);
            this.label2.Text = "User";
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(3, 98);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(62, 20);
            this.label3.Text = "Password";
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(3, 163);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(68, 20);
            this.label4.Text = "Auth Type";
            // 
            // label5
            // 
            this.label5.Location = new System.Drawing.Point(3, 205);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(62, 20);
            this.label5.Text = "Protocol";
            // 
            // lblStatus
            // 
            this.lblStatus.Location = new System.Drawing.Point(71, 307);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(248, 20);
            // 
            // lstMail
            // 
            this.lstMail.Columns.Add(this.colFrom);
            this.lstMail.Columns.Add(this.colSubject);
            this.lstMail.Columns.Add(this.colDate);
            this.lstMail.FullRowSelect = true;
            this.lstMail.Location = new System.Drawing.Point(317, 14);
            this.lstMail.Name = "lstMail";
            this.lstMail.Size = new System.Drawing.Size(387, 190);
            this.lstMail.TabIndex = 21;
            this.lstMail.View = System.Windows.Forms.View.Details;
            this.lstMail.SelectedIndexChanged += new System.EventHandler(this.lstMail_SelectedIndexChanged);
            // 
            // btnDel
            // 
            this.btnDel.Location = new System.Drawing.Point(630, 210);
            this.btnDel.Name = "btnDel";
            this.btnDel.Size = new System.Drawing.Size(72, 20);
            this.btnDel.TabIndex = 22;
            this.btnDel.Text = "Delete";
            this.btnDel.Click += new System.EventHandler(this.btnDel_Click);
            // 
            // label6
            // 
            this.label6.ForeColor = System.Drawing.SystemColors.Highlight;
            this.label6.Location = new System.Drawing.Point(14, 367);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(291, 54);
            this.label6.Text = "Warning: if \"leave a copy of message on server\" is not checked,  the emails on th" +
                "e server will be deleted !";
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
            // webMail
            // 
            this.webMail.Location = new System.Drawing.Point(317, 236);
            this.webMail.Name = "webMail";
            this.webMail.Size = new System.Drawing.Size(387, 200);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.AutoScroll = true;
            this.ClientSize = new System.Drawing.Size(722, 448);
            this.Controls.Add(this.webMail);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.btnDel);
            this.Controls.Add(this.lstMail);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.Label1);
            this.Controls.Add(this.pgBar);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnStart);
            this.Controls.Add(this.chkLeaveCopy);
            this.Controls.Add(this.lstProtocol);
            this.Controls.Add(this.lstAuthType);
            this.Controls.Add(this.chkSSL);
            this.Controls.Add(this.textPassword);
            this.Controls.Add(this.textUser);
            this.Controls.Add(this.textServer);
            this.Menu = this.mainMenu1;
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TextBox textServer;
        private System.Windows.Forms.TextBox textUser;
        private System.Windows.Forms.TextBox textPassword;
        private System.Windows.Forms.CheckBox chkSSL;
        private System.Windows.Forms.ComboBox lstAuthType;
        private System.Windows.Forms.ComboBox lstProtocol;
        private System.Windows.Forms.CheckBox chkLeaveCopy;
        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.ProgressBar pgBar;
        private System.Windows.Forms.Label Label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.ListView lstMail;
        private System.Windows.Forms.Button btnDel;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ColumnHeader colFrom;
        private System.Windows.Forms.ColumnHeader colSubject;
        private System.Windows.Forms.ColumnHeader colDate;
        private System.Windows.Forms.WebBrowser webMail;
    }
}

