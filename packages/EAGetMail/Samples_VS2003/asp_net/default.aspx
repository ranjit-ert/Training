<%@ Page language="c#" AutoEventWireup="true" validateRequest="false" CodePage="65001"%>
<%@ Import Namespace="System.IO"%>
<%@ Import Namespace="EAGetMail"%>

<script language="c#" runat="server">
private string m_uidlfile = "uidl.txt";
private string m_curpath = "";
		
private ArrayList m_arUidl = new ArrayList();
		
private void Page_Load(object sender, System.EventArgs e)
{
	// Put user code to initialize the page here
	ListMail();
}
				
private void btnCheck_Click(object sender, System.EventArgs e)
{
	if( textServer.Text.Length == 0 ||
		textUser.Text.Length == 0 ||
		textPassword.Text.Length == 0 )
	{
		lblDesc.Text = "Please input server, user and password!" ;
		return;
	}

	string tempFolder = String.Format( "{0}\\inbox",  Server.MapPath(""));
	m_curpath  = tempFolder;
	RetrieveEmail();
	ListMail();
}

private void ListMail()
{
	string tempFolder = String.Format( "{0}\\inbox",  Server.MapPath(""));
	m_curpath  = tempFolder;
	string mailFolder = String.Format( "{0}", m_curpath );
	if( !System.IO.Directory.Exists( mailFolder ))
		Directory.CreateDirectory( mailFolder );

	tbMail.Rows.Clear();

	TableRow row = new TableRow();
	TableCell cel = new TableCell();
	cel.Text = "<b>From</b>";
	row.Cells.Add( cel );
	
	cel = new TableCell();
	cel.Text = "<b>Subject</b>";
	row.Cells.Add( cel );

	cel = new TableCell();
	cel.Text = "<b>Date</b>";
	row.Cells.Add( cel );

	cel = new TableCell();
	cel.Text = "<b>&nbsp;</b>";
	row.Cells.Add( cel );

	tbMail.Rows.Add(row);

	string [] files = Directory.GetFiles( mailFolder, "*.eml" );
	int count = files.Length;
	for( int i = 0; i < count; i++ )
	{
		string fullname = files[i];
		//For evaluation usage, please use "TryIt" as the license code, otherwise the 
		//"invalid license code" exception will be thrown. However, the object will expire in 1-2 months, then
		//"trial version expired" exception will be thrown.
		Mail oMail = new Mail( "TryIt" );
		oMail.Load( fullname, true );
		row = new TableRow();
		cel = new TableCell();
		string s = Server.HtmlEncode(oMail.From.ToString());
		if( s.Length == 0 )
			s = "[no sender]";

		s = String.Format( "<a href=\"display.aspx?name={0}\" onclick=\"javascript:fnDisplay( this.href );return false;\">{1}</a>",
			Server.UrlEncode(fullname), s );

		cel.Text = s;
		row.Cells.Add( cel );
	
		cel = new TableCell();
		cel.Text = Server.HtmlEncode(oMail.Subject);
		row.Cells.Add( cel );

		cel = new TableCell();
		cel.Text = oMail.ReceivedDate.ToString( "yyyy-MM-dd HH:mm:ss");
		row.Cells.Add( cel );

		s = String.Format( "<a href=\"delete.aspx?name={0}\">Delete</a>",
			Server.UrlEncode(fullname) );

		cel = new TableCell();
		cel.Text = s;
		row.Cells.Add( cel );

		tbMail.Rows.Add(row);
		// Load( file, true ) only load the email header to Mail object to save the CPU and memory
		// the Mail object will load the whole email file later automatically if bodytext or attachment is required..
		oMail.Load( fullname, true );

		oMail.Clear();
	}
}
private void RetrieveEmail()
{
    int prtl = Int32.Parse(lstProtocol.SelectedValue);

	MailServer oServer = new MailServer( textServer.Text,
		textUser.Text, textPassword.Text, chkSSL.Checked,
        ServerAuthType.AuthLogin, (ServerProtocol)prtl);

    if( String.Compare( lstProtocol.SelectedValue, "1" ) == 0 )
    {
        oServer.AuthType = ServerAuthType.AuthCRAM5;
    }
    else if( String.Compare( lstProtocol.SelectedValue, "1" ) == 0 )
    {
        oServer.AuthType = ServerAuthType.AuthNTLM;
    }
    
	MailClient oClient = new MailClient("TryIt");
	bool bLeaveCopy = chkLeaveCopy.Checked;

    // UIDL is the identifier of every email on POP3/IMAP4/Exchange server, to avoid retrieve
    // the same email from server more than once, we record the email UIDL retrieved every time
    // if you delete the email from server every time and not to leave a copy of email on
    // the server, then please remove all the function about uidl.
    // UIDLManager wraps the function to write/read uidl record from a text file.
    UIDLManager oUIDLManager = new UIDLManager();
     
	try
	{
        // load existed uidl records to UIDLManager
        string uidlfile = String.Format("{0}\\{1}", m_curpath, m_uidlfile);
        oUIDLManager.Load(uidlfile);

		string mailFolder = String.Format( "{0}", m_curpath );
		if( !Directory.Exists( mailFolder ))
			Directory.CreateDirectory( mailFolder );

		oClient.Connect( oServer );
		MailInfo[] infos = oClient.GetMailInfos();
		Response.Write( String.Format( "Total {0} email(s)<br>" , infos.Length ));

        // remove the local uidl which is not existed on the server.
        oUIDLManager.SyncUIDL(oServer, infos);
        oUIDLManager.Update();
        
		int count = infos.Length;

		for( int i = 0; i < count; i++ )
		{
			MailInfo info = infos[i];
            if (oUIDLManager.FindUIDL(oServer, info.UIDL) != null)
            {
                //this email has been downloaded before.
                continue;
            }

			//lblStatus.Text = String.Format( "Retrieving {0}/{1}...", info.Index, count );
			
			Mail oMail = oClient.GetMail( info );
			System.DateTime d = System.DateTime.Now;
			System.Globalization.CultureInfo cur = new System.Globalization.CultureInfo("en-US");			
			string sdate = d.ToString("yyyyMMddHHmmss", cur);
			string fileName = String.Format( "{0}\\{1}{2}{3}.eml", mailFolder, sdate, d.Millisecond.ToString("d3"), i );
			oMail.SaveAs( fileName, true );

			
			if( bLeaveCopy )
			{
                //add the email uidl to uidl file to avoid we retrieve it next time. 
                oUIDLManager.AddUIDL(oServer, info.UIDL, fileName);
			}
		}

		if( !bLeaveCopy )
		{
			//lblStatus.Text = "Deleting ...";
            for (int i = 0; i < count; i++)
            {
                oClient.Delete(infos[i]);
                // Remove UIDL from local uidl file.
                oUIDLManager.RemoveUIDL(oServer, infos[i].UIDL);           
            }
		}
		// Delete method just mark the email as deleted, 
		// Quit method pure the emails from server exactly.
		oClient.Quit();
		Response.Write("Completed!<br>");
	
	}
	catch( Exception ep )
	{
		Response.Write( ep.Message + "<br>");
	}

    // Update the uidl list to local uidl file and then we can load it next time.
    oUIDLManager.Update();

}
</script>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<html>
	<head>
		<title>EAGetMail ASP_NET SAMPLE</title>
		<meta http-equiv="Content-Type" content="text-html; charset=utf-8" />
		<meta name="CODE_LANGUAGE" Content="C#" />
		<meta name="vs_defaultClientScript" content="JavaScript"/>
		<script type="text/jscript">
		function fnDisplay( v )
		{
			window.open( v, 
				"", 
				"height=640,width=800,menubar=no,resizable=yes,scrollbars=yes,status=yes", 
				false );
		}
		</script>
	</head>
	<body ms_positioning="GridLayout">
		<form id="Form1" method="post" runat="server">
		<p>
	<asp:Label id="lblDesc" runat="server" Width="560px" ForeColor="Red"></asp:Label>
</p>
			<table>
				<tr>
				    <td><asp:Label id="Label1"  runat="server" Width="73px" Height="24px">Server:</asp:Label></td>
				    <td><asp:TextBox id="textServer" runat="server" Width="248px"></asp:TextBox></td>
				</tr>
				<tr>
				    <td><asp:Label id="Label2" runat="server" Width="48px">User:</asp:Label></td>
				    <td><asp:TextBox id="textUser" runat="server" Width="248px"></asp:TextBox></td>
				</tr>
				<tr>
				    <td><asp:Label id="Label3" runat="server" Height="24px" Width="56px">Password:</asp:Label></td>
				    <td><asp:TextBox id="textPassword" runat="server" Width="248px"></asp:TextBox></td>
				</tr>
				<tr>
				    <td></td>
				    <td><asp:CheckBox id="chkSSL" runat="server" Height="24px" Width="328px" Text="My server requires SSL connection"></asp:CheckBox></td>
				</tr>
				<tr>
				    <td>Auth Type</td>
				    <td>
				<asp:DropDownList id="lstAuthType" runat="server" Height="24px" Width="248px">
					<asp:ListItem Value="0" Selected="True">User Login</asp:ListItem>
					<asp:ListItem Value="1">APOP/CRAM-MD5</asp:ListItem>
					<asp:ListItem Value="2">NTLM</asp:ListItem>
				</asp:DropDownList>					    
				    </td>
				</tr>
				<tr>
				    <td>Protocol</td>
				    <td><asp:DropDownList id="lstProtocol" runat="server" Height="24px" Width="248px">
					<asp:ListItem Value="0" Selected="True">POP3 Server</asp:ListItem>
					<asp:ListItem Value="1">IMAP4 Server</asp:ListItem>
					<asp:ListItem Value="2">Exchange Web Service - 2007/2010</asp:ListItem>
					<asp:ListItem Value="3">Exchange WebDAV - 2000/2003</asp:ListItem>
	
				</asp:DropDownList>
				<p>
				 Notice: By default, Exchange Web Service requires SSL connection.
       For other protocol, please set SSL connection based on your server setting manually
       </p>
				</td>
				</tr>
				<tr>
				    <td></td>
				    <td><asp:CheckBox id="chkLeaveCopy" runat="server" Height="16px" Width="336px" Text="Leave a copy of message on Server"></asp:CheckBox>

		        </td>
				</tr>
				<tr>
				    <td></td>
				    <td><asp:Button id="btnCheck"  runat="server" Height="32px" Width="336px" Text="Retrieve Email" OnClick="btnCheck_Click"></asp:Button>
                </td>
				</tr>
				</table>
				<asp:Table id="tbMail"  runat="server" Height="56px" Width="488px"></asp:Table>
		</form>
	</body>
</html>