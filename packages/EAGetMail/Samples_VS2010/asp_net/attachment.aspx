<%@ Page language="c#" AutoEventWireup="true" validateRequest="false" CodePage="65001"%>
<%@ Import Namespace="EAGetMail"%>

<script language="c#" runat="server">
private void Page_Load(object sender, System.EventArgs e)
{

	string emlFile =  Request.QueryString["name"];
	int index = Int32.Parse( Request.QueryString["seq"] );
	downloadAttachment( emlFile, index );
}

private void downloadAttachment( string emlFile, int index )
{
	Mail oMail = new Mail("TryIt");
	oMail.Load( emlFile, false );

	Attachment[] atts = oMail.Attachments;
	Attachment att = atts[index];
	Response.AppendHeader( "Content-Disposition", String.Format( "attachment;filename={0}", att.Name ));
	Response.ContentType = att.ContentType;
	Response.BinaryWrite( att.Content );

}
</script>