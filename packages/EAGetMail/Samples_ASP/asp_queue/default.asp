<%@  language="VBScript" codepage="65001" %>
<%
'  ===============================================================================
' |    THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF      |
' |    ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO    |
' |    THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A         |
' |    PARTICULAR PURPOSE.                                                    |
' |    Copyright (c)2010 ADMINSYSTEM SOFTWARE LIMITED                         |
' |
' |    File: Default.asp
' |    Project: This project demonstrates how to use EAGetMailObj to retrieve email by EAGetMail Service
' |    Author: Ivan Lui ( ivan@emailarchitect.net )
'  ===============================================================================


Response.Charset = "utf-8"

Dim oTools
Set oTools = Server.CreateObject("EAGetMailObj.Tools")

Dim mailFolder
mailFolder = Server.MapPath(Request.ServerVariables("PATH_INFO"))
pos = InStrRev( mailFolder, "\" )
If pos > 0 Then
	mailFolder = Mid( mailFolder, 1, pos-1 )
End If

mailFolder = mailFolder & "\inbox" 

If Not oTools.ExistFile( mailFolder ) Then
    oTools.CreateFolder mailFolder
End If


If Request.QueryString("act") = "get" Then
    GetMail
End If

Sub GetMail()
       Dim oClient, oServer
       
       Set oClient = Server.CreateObject("EAGetMailObj.MailClient")
        'For evaluation usage, please use "TryIt" as the license code, otherwise the
        '"invalid license code" exception will be thrown. However, the object will expire in 1-2 months, then
        '"trial version expired" exception will be thrown.
        oClient.LicenseCode = "TryIt"
        
        Set oServer = Server.CreateObject("EAGetMailObj.MailServer")
        
        'oClient.LogFileName = mailFolder & "\pop3.txt" 'generate a log file
        
        
        oServer.Server = Request.Form("Server")
        oServer.User = Request.Form("User")
        oServer.Password = Request.Form("Password")
        If Request.Form("chkSSL") <> "" Then
            oServer.SSLConnection = True
        End If
        oServer.AuthType = Request.Form("lstAuthType")
        oServer.Protocol = Request.Form("lstProtocol")
        
        Const MailServerPop3 = 0
        Const MailServerImap4 = 1
        Const MailServerEWS = 2
        Const MailServerDAV = 3

        If oServer.Protocol = MailServerImap4 Then
            If oServer.SSLConnection Then
                oServer.Port = 993 'SSL IMAP4
            Else
                oServer.Port = 143 'IMAP4 normal
            End If
        Else
            If oServer.SSLConnection Then
                oServer.Port = 995 'SSL POP3
            Else
                oServer.Port = 110 'POP3 normal
            End If
        End If
        
        Dim leaveCopy
        LeaveCopy = True
        If Request.Form("chkLeaveCopy") = "" Then
            LeaveCopy = False
        End If
       
        On Error Resume Next
        'send request to EAGetMail Service, then EAGetMail Service retrieves email
        'in background and 'this method returns immediately
        oClient.GetMailsByQueue oServer, mailFolder, LeaveCopy
        
        If Err.Number <> 0 Then
            Response.Write  "<font color=""red"">" & Err.Description & " please make sure you installed EAGetMail Service!</font>"
            On Error Goto 0
            Exit Sub
        End If
        
        Response.Write "<font color=""blue"">Task has been submitted to EAGetMail Service successfully, EAGetMail will retrieve emails in background, please click <b>Refresh local emails</b> later to check retrieved email.</font>"
End Sub


%>
<html>
<head>
    <meta http-equiv="content-type" content="text/html; charset=utf-8" />
    <title>EAGetMailObj ASP + EAGetMail Service Sample</title>

    <script type="text/jscript">
		function fnDisplay( v )
		{
			window.open( v, 
				"", 
				"height=640,width=800,menubar=no,resizable=yes,scrollbars=yes,status=yes", 
				false );
		}
   //Verifying input
   function fnSubmit()
   {
		if( document.thisForm.server.value == "" || 
		    document.thisForm.user.value  == "" ||
			document.thisForm.password.value  == ""
 		)
 		{
 			alert("Please input server address, username and password!")
 			return
 		}
 		document.thisForm.submit()
 		return
   }		
    </script>

</head>
<body>
    <form action="default.asp?act=get" name="thisForm" method="post">
    <p>
        To run this sample, you should download <a href="http://www.emailarchitect.net/webapp/download/eagetmailservice.exe">
            EAGetMail Service</a> and install it on the server.
    </p>
    <table>
        <tr>
            <td width="100">
                Server
            </td>
            <td>
                <input type="text" name="server" size="25" />
            </td>
        </tr>
        <tr>
            <td>
                User
            </td>
            <td>
                <input type="text" name="user" size="25" />
            </td>
        </tr>
        <tr>
            <td>
                Password
            </td>
            <td>
                <input type="password" name="password" size="26" />
            </td>
        </tr>
        <tr>
            <td>
                &nbsp;
            </td>
            <td>
                <input type="checkbox" name="chkSSL" value="1" />
                SSL Connection (*By default, Exchange Web Service requires SSL connection)
            </td>
        </tr>
        <tr>
            <td>
                Auth Type
            </td>
            <td>
                <select name="lstAuthType">
                    <option value="0">USER/LOGIN</option>
                    <option value="1">APOP(CRAM-MD5)</option>
                    <option value="2">NTLM</option>
                </select>
            </td>
        </tr>
        <tr>
            <td>
                Protocol
            </td>
            <td>
                <select name="lstProtocol">
                    <option value="0">POP3</option>
                    <option value="1">IMAP4</option>
                    <option value="2">Exchange Web Service - 2007/2010</option>
                    <option value="3">Exchange WebDAV- 2000/2003</option>
                    
                </select>
            </td>
        </tr>
        <tr>
            <td>
                &nbsp;
            </td>
            <td>
                <input type="checkbox" name="chkLeaveCopy" value="1" />Leave a copy of messsage
                on Server.
            </td>
        </tr>
        <tr>
            <td>
            </td>
            <td>
                <input type="button" name="btn1" value="Retrieve by EAGetMail Service in background" onclick="fnSubmit()" />
            </td>
        </tr>
        <tr>
            <td>
            </td>
            <td>
                            <a href="default.asp">Refresh local emails.</a>

            </td>
        </tr>
    </table>
    </form>
    <hr />
    <table width="100%">
        <tr>
            <td width="300">
                <b>From</b>
            </td>
            <td width="400">
                <b>Subject</b>
            </td>
            <td width="200">
                <b>Date</b>
            </td>
            <td>
                <b>&nbsp;</b>
            </td>
        </tr>
        <%
Function FormatHtmlTag( src ) 
    src = Replace(src, ">", "&gt;")
    src = Replace(src, "<", "&lt;")
    FormatHtmlTag = src
End Function

Dim arFile
arFile = oTools.GetFiles( mailFolder & "\*.eml" )

Set oMail = Server.CreateObject("EAGetMailObj.Mail")
oMail.LicenseCode = "TryIt"

Sub DisplayMail( fileName )
    Dim s
    oMail.LoadFile arFile(i), True

    s = "<tr>"
    s = s & "<td><a href=""display.asp?name=" & Server.URLEncode(arFile(i)) & """ onclick=""javascript:fnDisplay( this.href );return false;"">" 
    s = s & FormatHtmlTag( oMail.From.Name & "<" & oMail.From.Address & ">" ) & "</a></td>"
    s = s & "<td>" & FormatHtmlTag( oMail.Subject ) & "&nbsp;</td>"
    s = s & "<td>" & oMail.ReceivedDate & "</td>"
    s = s & "<td><a href=""delete.asp?name=" & Server.URLEncode(arFile(i)) & """>Delete</a></td>"
    
    Response.Write( s )
End Sub


For i = LBound(arFile) To UBound(arFile)
    DisplayMail arFile(i)
Next

        %>
    </table>
    <p>
        Technical Support<br />
        <br />
        Email: <a href="mailto:support@emailarchitect.net">support@emailarchitect.net</a>
    </p>
    <p>
        <a href="http://www.emailarchitect.net" target="_blank">2010 Copyright &copy; AdminSystem
            Software Limited. All rights reserved.</a>
    </p>
</body>
</html>
