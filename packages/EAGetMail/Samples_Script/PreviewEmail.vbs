'  ===============================================================================
' |    THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF      |
' |    ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO    |
' |    THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A         |
' |    PARTICULAR PURPOSE.                                                    |
' |    Copyright (c)2010  AdminSystem Software Limited                         |
' |
' |    Project: It demonstrates how to use EAGetMailObj to retrieve email summary in VBScript
' |
' |    Author: Ivan Lui ( ivan@emailarchitect.net )
'  ===============================================================================

Dim args, info
Set args = WScript.Arguments
If args.Count < 3 Then
	info =  "Usage: PreviewEmail.vbs [pop3server] [user] [password]" & Chr(13) & Chr(10)
	info = info & "   eg: PreviewEmail.vbs mail.adminsystem.net test@adminsystem.net test"
	WScript.Echo info
	WScript.Quit
End If

Const MailServerConnectType_ConnectSSLAuto = 0
Const MailServerConnectType_ConnectSSL = 1
Const MailServerConnectType_ConnectTLS = 2

Const ProxyType_Socks4 = 0
Const ProxyType_Socks5 = 1
Const ProxyType_Http = 2

Const MailServerAuthLogin = 0
Const MailServerAuthCRAM5 = 1
Const MailServerAuthNTLM = 2

Const MailServerPop3 = 0
Const MailServerImap4 = 1
Const MailServerEWS = 2 ' Exchange Web Service
Const MailServerDAV = 3 ' Exchange WebDAV


Dim oServer, oClient
Set oServer = CreateObject("EAGetMailObj.MailServer")
Set oClient = CreateObject("EAGetMailObj.MailClient")

'for evaluation usage, please use "TryIt", for licensed user, please use your license code here
oClient.LicenseCode = "TryIt"

oServer.Server = args(0)
oServer.User = args(1)
oServer.Password = args(2)

'if your pop3 requires SSL connection, please add
'oServer.SSLConnection = True
'oServer.Port = 995

'if your server is IMAP4 server, please add the following code
'oServer.Port = 143
'oServer.Protocol = MailServerImap4 

'if your server is IMAP4 server, and it requires SSL connection, please add the following code
'oServer.Protocol = MailServerImap4
'oServer.SSLConnection = True 
'oServer.Port = 993


WScript.Echo "Connecting " & args(0) & " ..."

oClient.Connect oServer

Dim arInfo
arInfo = oClient.GetMailInfos()

Dim i, Count
Count = UBound(arInfo)

WScript.Echo "Total " & Count - LBound(arInfo) & " email(s)" & Chr(13) & Chr(10)

For i = LBound(arInfo) To Count
	Dim oMail
	Set oMail = CreateObject("EAGetMailObj.Mail")
	'for evaluation usage, please use "TryIt", for licensed user, please use your license code here
	oMail.LicenseCode = "TryIt"
	
	'get email headers only, then you dont have to download entire email
	'but if you want to get email body/attachments, you should use 
	'Set oMail = oClient.GetMail(arInfo(i))
	
	WScript.Echo "get " & i + 1 & " email..."
	oMail.Load( oClient.GetMailHeader( arInfo(i)))
	
	Dim oHeaders
	Set oHeaders = oMail.Headers

	Dim headerString
	Dim x, y
	y = oHeaders.Count
	headerString = ""
	For x = 0 To y - 1
		Dim oHeader
		Set oHeader = oHeaders.Item(x)
		headerString = headerString & oHeader.HeaderKey & ":" & oHeader.HeaderValue & Chr(13) & Chr(10)
	Next
	WScript.Echo headerString
Next



oClient.Quit
