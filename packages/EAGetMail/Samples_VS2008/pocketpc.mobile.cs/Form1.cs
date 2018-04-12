using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.IO;
using EAGetMail;


namespace pocketpc.mobile.cs
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private bool m_bcancel = false;
        private string m_uidlfile = "uidl.txt";
        private string m_curpath = "";
        private ArrayList m_arUidl = new ArrayList();

        #region EAGetMail Event Handler
        public void OnConnected(object sender, ref bool cancel)
        {
            lblStatus.Text = "Connected ...";
            cancel = m_bcancel;
            Application.DoEvents();
        }

        public void OnQuit(object sender, ref bool cancel)
        {
            lblStatus.Text = "Quit ...";
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

        public void OnIdle(object sender, ref bool cancel)
        {
            cancel = m_bcancel;
            Application.DoEvents();
        }

        public void OnAuthorized(object sender, ref bool cancel)
        {
            lblStatus.Text = "Authorized ...";
            cancel = m_bcancel;
            Application.DoEvents();
        }

        public void OnSecuring(object sender, ref bool cancel)
        {
            lblStatus.Text = "Securing ...";
            cancel = m_bcancel;
            Application.DoEvents();
        }
        #endregion

        #region UIDL Functions
        // uidl is the identifier of every email on POP3/IMAP4 server, to avoid retrieve
        // the same email from server more than once, we record the email uidl retrieved every time
        // if you delete the email from server every time and not to leave a copy of email on
        // the server, then please remove all the function about uidl.

        private bool _FindUIDL(MailInfo[] infos, string uidl)
        {
            int count = infos.Length;
            for (int i = 0; i < count; i++)
            {
                if (String.Compare(infos[i].UIDL, uidl, false) == 0)
                    return true;
            }
            return false;
        }

        //remove the local uidl which is not existed on the server.
        private void _SyncUIDL(MailServer oServer, MailInfo[] infos)
        {
            string s = String.Format("{0}#{1} ", oServer.Server, oServer.User);

            bool bcontinue = false;
            int n = 0;
            do
            {
                bcontinue = false;
                int count = m_arUidl.Count;
                for (int i = n; i < count; i++)
                {
                    string x = m_arUidl[i] as string;
                    if (String.Compare(s, 0, x, 0, s.Length, true) == 0)
                    {
                        int pos = x.LastIndexOf(' ');
                        if (pos != -1)
                        {
                            string uidl = x.Substring(pos + 1);
                            if (!_FindUIDL(infos, uidl))
                            {
                                //this uidl doesn't exist on server, 
                                //so we should remove it from local uidl list to save the storage.
                                bcontinue = true;
                                n = i;
                                m_arUidl.RemoveAt(i);
                                break;
                            }
                        }
                    }
                }
            } while (bcontinue);

        }

        private bool _FindExistedUIDL(MailServer oServer, string uidl)
        {
            string s = String.Format("{0}#{1} {2}", oServer.Server.ToLower(), oServer.User.ToLower(), uidl);
            int count = m_arUidl.Count;
            for (int i = 0; i < count; i++)
            {
                string x = m_arUidl[i] as string;
                if (String.Compare(s, x, false) == 0)
                    return true;
            }
            return false;
        }

        private void _AddUIDL(MailServer oServer, string uidl)
        {
            string s = String.Format("{0}#{1} {2}", oServer.Server.ToLower(), oServer.User.ToLower(), uidl);
            m_arUidl.Add(s);
        }

        private void _UpdateUIDL()
        {
            StringBuilder s = new StringBuilder();
            int count = m_arUidl.Count;
            for (int i = 0; i < count; i++)
            {
                s.Append(m_arUidl[i] as string);
                s.Append("\r\n");
            }

            string file = String.Format("{0}\\{1}", m_curpath, m_uidlfile);

            FileStream fs = null;
            try
            {
                fs = new FileStream(file, FileMode.Create, FileAccess.Write, FileShare.None);
                byte[] data = System.Text.Encoding.Default.GetBytes(s.ToString());
                fs.Write(data, 0, data.Length);
                fs.Close();
            }
            catch (Exception ep)
            {
                if (fs != null)
                    fs.Close();

                throw ep;
            }

        }

        private void _LoadUIDL()
        {
            m_arUidl.Clear();
            string file = String.Format("{0}\\{1}", m_curpath, m_uidlfile);
            StreamReader read = null;
            try
            {
                read = File.OpenText(file);
                while (true)
                {
                    string line = read.ReadLine().Trim("\r\n \t".ToCharArray());
                    m_arUidl.Add(line);
                }
            }
            catch (Exception ep)
            { }

            if (read != null)
                read.Close();
        }
        #endregion

        #region Parse and Display Mails
        private void LoadMails()
        {
            lstMail.Items.Clear();
            string mailFolder = String.Format("{0}\\inbox", m_curpath);
            if (!Directory.Exists(mailFolder))
                Directory.CreateDirectory(mailFolder);

            string[] files = Directory.GetFiles(mailFolder, "*.eml");
            int count = files.Length;
            for (int i = 0; i < count; i++)
            {
                string fullname = files[i];
                //For evaluation usage, please use "TryIt" as the license code, otherwise the 
                //"invalid license code" exception will be thrown. However, the object will expire in 1-2 months, then
                //"trial version expired" exception will be thrown.
                Mail oMail = new Mail("TryIt");

                // Load( file, true ) only load the email header to Mail object to save the CPU and memory
                // the Mail object will load the whole email file later automatically if bodytext or attachment is required..
                oMail.Load(fullname, true);

                ListViewItem item = new ListViewItem(oMail.From.ToString());
                item.SubItems.Add(oMail.Subject);
                item.SubItems.Add(oMail.ReceivedDate.ToString("yyyy-MM-dd HH:mm:ss"));
                item.Tag = fullname;
                lstMail.Items.Add(item);

                int pos = fullname.LastIndexOf(".");
                string mainName = fullname.Substring(0, pos);
                string htmlName = mainName + ".htm";
                if (!File.Exists(htmlName))
                {
                    // this email is unread, we set the font style to bold.
                    //item.Font = new System.Drawing.Font(item.Font, FontStyle.Bold);
                }

                oMail.Clear();
            }
        }

        private string _FormatHtmlTag(string src)
        {
            src = src.Replace(">", "&gt;");
            src = src.Replace("<", "&lt;");
            return src;
        }

        //we generate a html + attachment folder for every email, once the html is create,
        // next time we don't need to parse the email again.
        private void _GenerateHtmlForEmail(string htmlName, string emlFile, string tempFolder)
        {
            //For evaluation usage, please use "TryIt" as the license code, otherwise the 
            //"invalid license code" exception will be thrown. However, the object will expire in 1-2 months, then
            //"trial version expired" exception will be thrown.
            Mail oMail = new Mail("TryIt");
            oMail.Load(emlFile, false);

            
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
                    attname = attname.Replace("\t", "");
                    att.SaveAs(attname, true);
                    hdr.Append(String.Format("<a href=\"{0}\" target=\"_blank\">{1}</a> ", attname, att.Name));
                    if (att.ContentID.Length > 0)
                    {	//show embedded image.
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
                hdr.Insert(0, "<meta HTTP-EQUIV=\"Content-Type\" Content=\"text-html; charset=utf-8\">");
            }

            html = hdr.ToString() + "<hr>" + html;
            FileStream fs = new FileStream(htmlName, FileMode.Create, FileAccess.Write, FileShare.None);
            byte[] data = System.Text.UTF8Encoding.UTF8.GetBytes(html);
            fs.Write(data, 0, data.Length);
            fs.Close();
            oMail.Clear();
        }

        private void ShowMail(string fileName)
        {
            try
            {
                int pos = fileName.LastIndexOf(".");
                string mainName = fileName.Substring(0, pos);
                string htmlName = mainName + ".htm";

                string tempFolder = mainName;
                if (!File.Exists(htmlName))
                {	//we haven't generate the html for this email, generate it now.
                    _GenerateHtmlForEmail(htmlName, fileName, tempFolder);
                }

                webMail.Navigate(new System.Uri(htmlName));
            }
            catch (Exception ep)
            {
                MessageBox.Show(ep.Message);
            }
        }
        #endregion

        private void Form1_Load(object sender, EventArgs e)
        {

            System.Reflection.Assembly asb = System.Reflection.Assembly.GetExecutingAssembly();
            System.Reflection.Module[] ms = asb.GetModules();

            string file = "pocketpc.mobile.cs.exe";
            int count = ms.Length;
            for (int i = 0; i < count; i++)
            {
                if (String.Compare(file, ms[i].Name, true) == 0)
                {
                    file = ms[i].FullyQualifiedName;
                    break;
                }
            }

            string path = file;
            int pos = path.LastIndexOf('\\');
            if (pos != -1)
                path = path.Substring(0, pos);

            m_curpath = path;

            webMail.Navigate(new System.Uri("about:blank"));
            lstProtocol.Items.Add("POP3");
            lstProtocol.Items.Add("IMAP4");
            lstProtocol.Items.Add("Exchange Web Service - 2007/2010");
            lstProtocol.Items.Add("Exchange WebDAV - Exchange 2000/2003");

            lstProtocol.SelectedIndex = 0;

            lstAuthType.Items.Add("USER/LOGIN");
            lstAuthType.Items.Add("APOP");
            lstAuthType.Items.Add("NTLM");
            lstAuthType.SelectedIndex = 0;

            LoadMails();

        }

        private void btnStart_Click(object sender, EventArgs e)
        {
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
            btnCancel.Enabled = true;

            ServerAuthType authType = ServerAuthType.AuthLogin;
            if (lstAuthType.SelectedIndex == 1)
                authType = ServerAuthType.AuthCRAM5;
            else if (lstAuthType.SelectedIndex == 2)
                authType = ServerAuthType.AuthNTLM;

            ServerProtocol protocol = (ServerProtocol)lstProtocol.SelectedIndex;

            MailServer oServer = new MailServer(server, user, password,
                chkSSL.Checked, authType, protocol);

            //For evaluation usage, please use "TryIt" as the license code, otherwise the 
            //"invalid license code" exception will be thrown. However, the object will expire in 1-2 months, then
            //"trial version expired" exception will be thrown.
            MailClient oClient = new MailClient("TryIt");

            //Catching the following events is not necessary, 
            //just make the application more user friendly.
            //If you use the object in asp.net/windows service or non-gui application, 
            //You need not to catch the following events.
            //To learn more detail, please refer to the code in EAGetMail EventHandler region
            oClient.OnAuthorized += new MailClient.OnAuthorizedEventHandler(OnAuthorized);
            oClient.OnConnected += new MailClient.OnConnectedEventHandler(OnConnected);
            oClient.OnIdle += new MailClient.OnIdleEventHandler(OnIdle);
            oClient.OnSecuring += new MailClient.OnSecuringEventHandler(OnSecuring);
            oClient.OnReceivingDataStream += new MailClient.OnReceivingDataStreamEventHandler(OnReceivingDataStream);

            bool bLeaveCopy = chkLeaveCopy.Checked;
            try
            {
                // uidl is the identifier of every email on POP3/IMAP4 server, to avoid retrieve
                // the same email from server more than once, we record the email uidl retrieved every time
                // if you delete the email from server every time and not to leave a copy of email on
                // the server, then please remove all the function about uidl.
                _LoadUIDL();

                string mailFolder = String.Format("{0}\\inbox", m_curpath);
                if (!Directory.Exists(mailFolder))
                    Directory.CreateDirectory(mailFolder);

                m_bcancel = false;
                oClient.Connect(oServer);
                MailInfo[] infos = oClient.GetMailInfos();
                lblStatus.Text = String.Format("Total {0} email(s)", infos.Length);

                _SyncUIDL(oServer, infos);
                int count = infos.Length;

                for (int i = 0; i < count; i++)
                {
                    MailInfo info = infos[i];
                    if (_FindExistedUIDL(oServer, info.UIDL))
                    {
                        //this email has existed on local disk.
                        continue;
                    }

                    lblStatus.Text = String.Format("Retrieving {0}/{1}...", info.Index, count);

                    Mail oMail = oClient.GetMail(info);
                    System.DateTime d = System.DateTime.Now;
                    System.Globalization.CultureInfo cur = new System.Globalization.CultureInfo("en-US");
                    string sdate = d.ToString("yyyyMMddHHmmss", cur);
                    string fileName = String.Format("{0}\\{1}{2}{3}.eml", mailFolder, sdate, d.Millisecond.ToString("d3"), i);
                    oMail.SaveAs(fileName, true);

                    ListViewItem item = new ListViewItem(oMail.From.ToString());
                    item.SubItems.Add(oMail.Subject);
                    item.SubItems.Add(oMail.ReceivedDate.ToString("yyyy-MM-dd HH:mm:ss"));
                   // item.Font = new System.Drawing.Font(item.Font, FontStyle.Bold);
                    item.Tag = fileName;
                    lstMail.Items.Insert(0, item);
                    oMail.Clear();

                   // lblTotal.Text = String.Format("Total {0} email(s)", lstMail.Items.Count);

                    if (bLeaveCopy)
                    {
                        //add the email uidl to uidl file to avoid we retrieve it next time. 
                        _AddUIDL(oServer, info.UIDL);
                    }
                }

                if (!bLeaveCopy)
                {
                    lblStatus.Text = "Deleting ...";
                    for (int i = 0; i < count; i++)
                        oClient.Delete(infos[i]);
                }
                // Delete method just mark the email as deleted, 
                // Quit method pure the emails from server exactly.
                oClient.Quit();

            }
            catch (Exception ep)
            {
                MessageBox.Show(ep.Message);
            }

            //update the uidl list to a text file and then we can load it next time.
            _UpdateUIDL();

            lblStatus.Text = "Completed";
            pgBar.Maximum = 100;
            pgBar.Minimum = 0;
            pgBar.Value = 0;
            btnStart.Enabled = true;
            btnCancel.Enabled = false;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            m_bcancel = true;
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            ListView.SelectedIndexCollection items = lstMail.SelectedIndices;
            if (items.Count == 0)
                return;

            if (MessageBox.Show("Do you want to delete all selected emails",
                            "",
                            MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.No)
                return;

            ArrayList ar = new ArrayList();
            for( int i = 0; i < items.Count; i++ )
            {
                int index = items[i];
                ListViewItem oItem = lstMail.Items[index];
                ar.Add(oItem);
                try
                {
                    string fileName = oItem.Tag as string;
                    File.Delete(fileName);
                    int pos = fileName.LastIndexOf(".");
                    string tempFolder = fileName.Substring(0, pos);
                    string htmlName = tempFolder + ".htm";
                    if (File.Exists(htmlName))
                        File.Delete(htmlName);

                    if (Directory.Exists(tempFolder))
                    {
                        Directory.Delete(tempFolder, true);
                    }
                }
                catch (Exception ep)
                {
                    MessageBox.Show(ep.Message);
                    break;
                }
            }

            for (int i = 0; i < ar.Count; i++)
            {
                ListViewItem oItem = ar[i] as ListViewItem;
                lstMail.Items.Remove(oItem);
            }

            //lblTotal.Text = String.Format("Total {0} email(s)", lstMail.Items.Count);


            webMail.Navigate( new System.Uri("about:blank"));
        }

        private void lstMail_SelectedIndexChanged(object sender, EventArgs e)
        {
            ListView.SelectedIndexCollection items = lstMail.SelectedIndices;
            if (items.Count == 0)
                return;

            ListViewItem oItem = lstMail.Items[items[0]] as ListViewItem;
            ShowMail(oItem.Tag as string);
        }

    }
}