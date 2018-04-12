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
Function FormatHtmlTag( src ) 
    src = Replace(src, ">", "&gt;")
    src = Replace(src, "<", "&lt;")
    FormatHtmlTag = src
End Function

Dim filename
filename = Request.QueryString("name")
Dim oMail
Set oMail = Server.CreateObject("EAGetMailObj.Mail")

   'For evaluation usage, please use "TryIt" as the license code, otherwise the
    '"invalid license code" exception will be thrown. However, the object will expire in 1-2 months, then
    '"trial version expired" exception will be thrown.
    oMail.LicenseCode = "TryIt"
        
    oMail.LoadFile filename, False

    
    Dim html
    html = oMail.HtmlBody
    Dim hdr
    hdr = hdr & "<font face=""Courier New,Arial"" size=2>"
    hdr = hdr & "<b>From:</b> " + FormatHtmlTag(oMail.From.Name & "<" & oMail.From.Address & ">") + "<br>"
    
    
    Dim addrs
    addrs = oMail.To
    Dim i, Count
    
    i = LBound(addrs)
    Count = UBound(addrs)
    If (Count >= i) Then
        hdr = hdr & "<b>To:</b> "
        For i = LBound(addrs) To Count
            hdr = hdr & FormatHtmlTag(addrs(i).Name & "<" & addrs(i).Address & ">")
            If (i < Count) Then
                hdr = hdr & ";"
            End If
        Next
        hdr = hdr & "<br>"
    End If
    
    addrs = oMail.Cc
    i = LBound(addrs)
    Count = UBound(addrs)
    If (Count >= i) Then
        hdr = hdr & "<b>Cc:</b> "
        For i = LBound(addrs) To Count
            hdr = hdr & FormatHtmlTag(addrs(i).Name & "<" & addrs(i).Address & ">")
            If (i < Count) Then
                hdr = hdr & ";"
            End If
        Next
        hdr = hdr & "<br>"
    End If
        
    hdr = hdr & "<b>Subject:</b>" & FormatHtmlTag(oMail.Subject) & "<br>" & vbCrLf
    Dim atts
    atts = oMail.Attachments
    i = LBound(atts)
    Count = UBound(atts)
    If (Count >= i) Then
       
        hdr = hdr & "<b>Attachments:</b>"
        For i = LBound(atts) To Count
            Dim att
            Set att = atts(i)

            Dim attname
            attname = "attachment.asp?name=" & Server.URLEncode(filename) & "&index=" & i
            
            hdr = hdr & "<a href=""" & attname & """ target=""_blank"">" & att.Name & "</a> "
            If Len(att.ContentID) > 0 Then
                'show embedded image.
                html = Replace(html, "cid:" + att.ContentID, attname)
            ElseIf InStr(1, att.ContentType, "image/", vbTextCompare) = 1 Then
                'show attached image.
                html = html & "<hr><img src=""" & attname & """>"
            End If

        Next
    End If
    
    hdr = "<meta HTTP-EQUIV=""Content-Type"" Content=""text/html; charset=utf-8"">" & hdr
    html = hdr & "<hr>" & html
   
    oMail.Clear
    Response.Write( html )
%>