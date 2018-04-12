'  ===============================================================================
' |    THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF      |
' |    ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO    |
' |    THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A         |
' |    PARTICULAR PURPOSE.                                                    |
' |    Copyright (c)2010  AdminSystem Software Limited                         |
' |
' |    Project: It demonstrates how to use EAGetMailObj to parse email summary in VBScript
' |
' |    Author: Ivan Lui ( ivan@emailarchitect.net )
'  ===============================================================================

Dim args, info
Set args = WScript.Arguments
If args.Count < 1 Then
	info =  "Usage: ParseEmail.vbs [email file path]" & Chr(13) & Chr(10)
	info = info & "   eg: ParseEmail.vbs ""d:\1.eml"""
	WScript.Echo info
	WScript.Quit
End If

Dim oMail
Set oMail = CreateObject("EAGetMailObj.Mail")
'for evaluation usage, please use "TryIt", for licensed user, please use your license code here
oMail.LicenseCode = "TryIt"

oMail.LoadFile args(0), False
If oMail.IsEncrypted  Then
	Set oMail = oMail.Decrypt( Nothing )
End If

Dim oCert
If oMail.IsSigned Then
	Set oCert = oMail.VerifySignature()
End If


Dim headerString
headerString = "From: " & oMail.From.Name & "<" & oMail.From.Address & ">" & Chr(13) & Chr(10)
headerString = headerString & "Reply-To: " & oMail.ReplyTo.Name & "<" & oMail.ReplyTo.Address & ">" & Chr(13) & Chr(10)


headerString = headerString & "To:"
ar = oMail.To
For i = LBound(ar) To UBound(ar)
	headerString = headerString &  ar(i).Name & "<" & ar(i).Address & ">;" & Chr(13) & Chr(10)
Next

headerString = headerString & "Cc:"
ar = oMail.Cc
For i = LBound(ar) To UBound(ar)
	headerString = headerString & ar(i).Name & "<" & ar(i).Address & ">;" & Chr(13) & Chr(10)
Next

headerString = headerString & "Subject: " & oMail.Subject & Chr(13) & Chr(10)
headerString = headerString & "Date: " & oMail.ReceivedDate & Chr(13) & Chr(10)

headerString = headerString & "Attachments: "
ar = oMail.Attachments
For i = LBound(ar) To UBound(ar)
	Dim oAtt
	Set oAtt = ar(i)
	headerString = headerString & oAtt.ContentType & ":" & oAtt.Name & ";" & Chr(13) & Chr(10)
Next


WScript.Echo headerString
WScript.Echo ""
WScript.Echo oMail.TextBody

' display the htmlbody
'WScript.Echo oMail.HtmlBody

oMail.Clear