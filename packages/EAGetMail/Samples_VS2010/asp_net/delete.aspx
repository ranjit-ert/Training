<%@ Page language="c#" AutoEventWireup="true" validateRequest="false" CodePage="65001"%>
<%@ Import Namespace="System.IO"%>
<%@ Import Namespace="EAGetMail"%>
<script language="c#" runat="server">
private void Page_Load(object sender, System.EventArgs e)
{
	// Put user code to initialize the page here
	string filename = Request.QueryString["name"];
	string tempFolder = String.Format( "{0}\\inbox",  Server.MapPath(""));
	if( String.Compare( filename, 0, tempFolder, 0, tempFolder.Length, true ) == 0 )
	{
	//only delete the file under temp folder
	    System.IO.File.Delete( filename );
	}
	Response.Redirect( "default.aspx" );
}
</script>		
