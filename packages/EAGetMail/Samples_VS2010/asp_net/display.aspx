<%@ Page language="c#" AutoEventWireup="true" validateRequest="false" CodePage="65001"%>
<%@ Import Namespace="EAGetMail"%>

<script language="c#" runat="server">
private void Page_Load(object sender, System.EventArgs e)
{
	displayMail( Request.QueryString["name"]);
	// Put user code to initialize the page here
}

private void displayMail( string emlFile )
{
	//For evaluation usage, please use "TryIt" as the license code, otherwise the 
	//"invalid license code" exception will be thrown. However, the object will expire in 1-2 months, then
	//"trial version expired" exception will be thrown.
	Mail oMail = new Mail("TryIt");
	oMail.Load( emlFile, false );

	string html = oMail.HtmlBody;
	System.Text.StringBuilder hdr = new System.Text.StringBuilder();

	hdr.Append( "<font face=\"Courier New,Arial\" size=2>" );
	hdr.Append(  "<b>From:</b> " + _FormatHtmlTag(oMail.From.ToString()) + "<br>" );
	MailAddress [] addrs = oMail.To;
	int count = addrs.Length;
	if( count > 0 )
	{
		hdr.Append( "<b>To:</b> ");
		for( int i = 0; i < count; i++ )
		{
			hdr.Append(  _FormatHtmlTag(addrs[i].ToString()));
			if( i < count - 1 )
			{
				hdr.Append( ";" );
			}
		}
		hdr.Append( "<br>" );
	}

	addrs = oMail.Cc;

	count = addrs.Length;
	if( count > 0 )
	{
		hdr.Append( "<b>Cc:</b> ");
		for( int i = 0; i < count; i++ )
		{
			hdr.Append(  _FormatHtmlTag(addrs[i].ToString()));
			if( i < count - 1 )
			{
				hdr.Append( ";" );
			}
		}
		hdr.Append( "<br>" );
	}

	hdr.Append( String.Format( "<b>Subject:</b>{0}<br>\r\n",  _FormatHtmlTag(oMail.Subject)));

	Attachment [] atts = oMail.Attachments;
	count = atts.Length;
	if( count > 0 )
	{
		hdr.Append( "<b>Attachments:</b>" );
		for( int i = 0; i < count; i++ )
		{
			Attachment att = atts[i];
		
			string attname = String.Format( "attachment.aspx?name={0}&seq={1}", Server.UrlEncode(emlFile), i  );
			
			hdr.Append( String.Format("<a href=\"{0}\" target=\"_blank\">{1}</a> ", attname, att.Name ));
			if( att.ContentID.Length > 0 )
			{	//show embedded image.
				html = html.Replace( "cid:" + att.ContentID, attname );
			}
			else if( String.Compare( att.ContentType, 0, "image/", 0, "image/".Length, true ) == 0 )
			{
				//show attached image.
				html = html + String.Format( "<hr><img src=\"{0}\">", attname );
			}
			else if( String.Compare( att.ContentType, 0, "text/plain", 0, "text/plain".Length, true ) == 0 )
			{
				html = html + "<pre>";
				html = html + System.Text.Encoding.Default.GetString( att.Content ) + "<pre>";
			}
		}
	}

	System.Text.RegularExpressions.Regex reg = new System.Text.RegularExpressions.Regex( "(<meta[^>]*charset[ \t]*=[ \t\"]*)([^<> \r\n\"]*)", 
		System.Text.RegularExpressions.RegexOptions.Multiline | System.Text.RegularExpressions.RegexOptions.IgnoreCase );
	html = reg.Replace( html, "$1utf-8" );
	if( !reg.IsMatch( html ))
	{
		hdr.Insert( 0, "<meta HTTP-EQUIV=\"Content-Type\" Content=\"text/html; charset=utf-8\">" );
	}

	html = hdr.ToString() + "<hr>" + html;
	oMail.Clear();
	oMail = null;
	Response.Write( html );
}

private string _FormatHtmlTag( string src )
{
	src = src.Replace( ">", "&gt;" );
	src = src.Replace( "<", "&lt;" );
	return src;
}
</script>