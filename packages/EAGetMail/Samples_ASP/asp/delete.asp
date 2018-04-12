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

Dim filename
filename = Request.QueryString("name")
Dim mailFolder
mailFolder = Server.MapPath(Request.ServerVariables("PATH_INFO"))
pos = InStrRev( mailFolder, "\" )
If pos > 0 Then
	mailFolder = Mid( mailFolder, 1, pos-1 )
End If

mailFolder = mailFolder & "\inbox" 

If InStr( 1, filename, mailFolder, 1 ) = 1 Then
    'only delete the specified file in the inbox 
    Dim oTools
    Set oTools = Server.CreateObject("EAGetMailObj.Tools")
    oTools.RemoveFile filename
End If

Response.Redirect "default.asp"
%>