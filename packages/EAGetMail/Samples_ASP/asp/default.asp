<%@  language="VBScript" codepage="65001" %>
<%
'  ===============================================================================
' |    THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF      |
' |    ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO    |
' |    THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A         |
' |    PARTICULAR PURPOSE.                                                    |
' |    Copyright (c)2010 ADMINSYSTEM.NET                         |
' |
' |    File: Default.asp
' |    Project: This project demonstrates how to use EAGetMailObj to retrieve email from POP3 server and save it as an email file.
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

Dim m_uidlfile
m_uidlfile = mailFolder & "\uidl.txt"

Dim m_uidls
m_uidls = ""

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
       
        On Error Resume Next
        oClient.Connect oServer
        
        If Err.Number <> 0 Then
            Response.Write Err.Description
            On Error Goto 0
            Exit Sub
        End If
        ' uidl is the identifier of every email on POP3/IMAP4 server, to avoid retrieve
        ' the same email from server more than once, we record the email uidl retrieved every time
        ' if you delete the email from server every time and not to leave a copy of email on
        ' the server, then please remove all the function about uidl.
        LoadUIDL
        
        Dim arInfo
        arInfo = oClient.GetMailInfos()
        
         If Err.Number <> 0 Then
            Response.Write Err.Description
            On Error Goto 0
            Exit Sub
        End If
         
        On Error Goto 0
               
        SyncUIDL oServer, arInfo
        
        Dim i, Count
        Count = UBound(arInfo)
        For i = LBound(arInfo) To Count
            Dim oInfo
            Set oInfo = arInfo(i)
            
            'if this email has not retrieved before, then get it
            If Not FindExistedUIDL(oServer, oInfo.UIDL) Then
                Dim oMail
                Set oMail = oClient.GetMail(oInfo)

                Dim fileName
                'generate a random file name by current local datetime,
                'you can use your method to generate the filename if you do not like it
                fileName = mailFolder & "\" & oTools.GenFileName(i) & ".eml"
               
                oMail.SaveAs fileName, True
                
                
                If Request.Form("chkLeaveCopy") <> "" Then
                    'add the email uidl to uidl file to avoid we retrieve it next time.
                    AddUIDL oServer, oInfo.UIDL
                End If
            End If
        Next
        
        If Request.Form("chkLeaveCopy") = "" Then
            For i = LBound(arInfo) To Count
                oClient.Delete arInfo(i)
            Next
        End If
        
        ' Delete method just mark the email as deleted,
        ' Quit method pure the emails from server exactly.
        oClient.Quit
        Set oClient = Nothing
        Response.Write( "Total " & (Count + 1) & " email(s) on server" )
        'update the uidl list to a text file and then we can load it next time.
        UpdateUIDL
End Sub

'==========================================================================================
' uidl is the identifier of every email on POP3/IMAP4 server, to avoid retrieve
' the same email from server more than once, we record the email uidl retrieved every time
' if you delete the email from server every time and not to leave a copy of email on
' the server, then please remove all the function about uidl.
'==========================================================================================
Sub LoadUIDL()
    m_uidls = ""
    m_uidls = oTools.ReadTextFile(m_uidlfile, 0)
End Sub

Sub UpdateUIDL()
    oTools.WriteTextFile m_uidlfile, m_uidls, 0
End Sub

Sub SyncUIDL(ByRef oServer, ByRef infos)
    Dim arLocal
    Dim newuidls 
    Dim uidlpref 
    uidlpref = LCase(oServer.Server) & "#" & LCase(oServer.User) & " "
    
    arLocal = Split(m_uidls, vbCrLf)
    Dim i, Count
    For i = LBound(arLocal) To UBound(arLocal)
        Dim t 
        Dim bRemove
        bRemove = False
        t = arLocal(i)
        If Len(t) > Len(uidlpref) Then
            Dim curPref
            curPref = Left(t, Len(uidlpref))
            If StrComp(curPref, uidlpref, 0) = 0 Then
                Dim localuidl
                localuidl = Mid(t, Len(uidlpref) + 1)
               
                Dim bExistOnServer
                bExistOnServer = False
                
                Dim x
                For x = LBound(infos) To UBound(infos)
                    If StrComp(infos(x).UIDL, localuidl, vbBinaryCompare) = 0 Then
                        bExistOnServer = True
                    Exit For
                    End If
                Next
                If Not bExistOnServer Then
                    bRemove = True
                End If
            End If
        End If
        
        If Trim(t) = "" Then
            bRemove = True
        End If
        
        If Not bRemove Then
            newuidls = newuidls & t & vbCrLf
        End If
    Next
    
    m_uidls = newuidls
End Sub

Function FindExistedUIDL(ByRef oServer, ByRef UIDL)
        Dim s
        s = LCase(oServer.Server) & "#" & LCase(oServer.User) & " " & UIDL & vbCrLf
        If InStr(1, m_uidls, s, 0 ) > 0 Then
            FindExistedUIDL = True
        Else
            FindExistedUIDL = False
        End If
End Function

Sub AddUIDL(ByRef oServer, ByRef UIDL )
        Dim s
        s = LCase(oServer.Server) & "#" & LCase(oServer.User) & " " & UIDL & vbCrLf
        m_uidls = m_uidls & s
End Sub
'=========================================================================================
' UIDL function end
'=========================================================================================
%>
<html>
<head>
    <meta http-equiv="content-type" content="text/html; charset=utf-8" />
    <title>EAGetMailObj ASP Sample</title>

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
                <input type="button" name="btn1" value="Retrieve" onclick="fnSubmit()" />
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

Dim s
For i = LBound(arFile) To UBound(arFile)
    oMail.LoadFile arFile(i), True

    s = "<tr>"
    s = s & "<td><a href=""display.asp?name=" & Server.URLEncode(arFile(i)) & """ onclick=""javascript:fnDisplay( this.href );return false;"">" 
    s = s & FormatHtmlTag( oMail.From.Name & "<" & oMail.From.Address & ">" ) & "</a></td>"
    s = s & "<td>" & FormatHtmlTag( oMail.Subject ) & "&nbsp;</td>"
    s = s & "<td>" & oMail.ReceivedDate & "</td>"
    s = s & "<td><a href=""delete.asp?name=" & Server.URLEncode(arFile(i)) & """>Delete</a></td>"
    
    Response.Write( s )
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
