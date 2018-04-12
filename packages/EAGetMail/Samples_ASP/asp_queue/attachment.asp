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
Dim oMail
Set oMail = Server.CreateObject("EAGetMailObj.Mail")

   'For evaluation usage, please use "TryIt" as the license code, otherwise the
    '"invalid license code" exception will be thrown. However, the object will expire in 1-2 months, then
    '"trial version expired" exception will be thrown.
    oMail.LicenseCode = "TryIt"
        
    oMail.LoadFile filename, False

Dim arAtt
arAtt = oMail.Attachments

Dim att
Set att = arAtt(Request.QueryString("index"))
Response.ContentType = att.ContentType
Response.AddHeader "Content-Disposition", "attachment;filename=" & att.Name 
Response.BinaryWrite att.Content
%>