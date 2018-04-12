// ===============================================================================
// |    THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF      |
// |    ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO    |
// |    THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A         |
// |    PARTICULAR PURPOSE.                                                    |
// |    Copyright (c)2010  AdminSystem Software Limited                       |
// |
// |    Project: It demonstrates how to use EAGetMailObj to retrieve email summary in JScript
// |
// |    Author: Ivan Lui ( ivan@emailarchitect.net )
//  ===============================================================================

var info = "";
var args = WScript.Arguments;
if( args.Length < 3 )
{
	info =  "Usage: PreviewEmail.js [pop3server] [user] [password]\r\n";
	info = info + "   eg: PreviewEmail.js mail.adminsystem.net test@adminsystem.net test";
	WScript.Echo( info );
	WScript.Quit();
}

var MailServerConnectType_ConnectSSLAuto = 0;
var MailServerConnectType_ConnectSSL = 1;
var MailServerConnectType_ConnectTLS = 2;

var ProxyType_Socks4 = 0;
var ProxyType_Socks5 = 1;
var ProxyType_Http = 2;

var MailServerAuthLogin = 0;
var MailServerAuthCRAM5 = 1;
var MailServerAuthNTLM = 2;

var MailServerPop3 = 0;
var MailServerImap4 = 1;
var MailServerEWS = 2; // Exchange Web Service
var MailServerDAV = 3; // Exchange WebDAV


var oServer = new ActiveXObject("EAGetMailObj.MailServer");
var oClient = new ActiveXObject("EAGetMailObj.MailClient");

//for evaluation usage, please use "TryIt", for licensed user, please use your license code here
oClient.LicenseCode = "TryIt";

//oServer.Server = args(0);
//oServer.User = args(1);
//oServer.Password = args(2);

//if your pop3 requires SSL connection, please add
oServer.SSLConnection = true;
oServer.Port = 995;

//if your server is IMAP4 server, please add the following code
//oServer.Port = 143;
//oServer.Protocol = MailServerImap4;

//if your server is IMAP4 server, and it requires SSL connection, please add the following code
//oServer.Protocol = MailServerImap4;
//oServer.SSLConnection = true;
//oServer.Port = 993;


WScript.Echo( "Connecting " + args(0) + " ..." );

oClient.Connect( oServer )

var arInfo = new VBArray(oClient.GetMailInfos()).toArray();

var i, Count;
Count = arInfo.length;

WScript.Echo( "Total "+ Count + " email(s)\r\n" );

for( i = 0; i < Count; i++ )
{
	var oMail = new ActiveXObject("EAGetMailObj.Mail");
	//for evaluation usage, please use "TryIt", for licensed user, please use your license code here
	oMail.LicenseCode = "TryIt";
	
	//get email headers only, then you dont have to download entire email
	//but if you want to get email body/attachments, you should use 
	//oMail = oClient.GetMail(arInfo(i));
	
	WScript.Echo("get " + ( i + 1 )+ " email..." );
	oMail.Load( oClient.GetMailHeader( arInfo[i]));
	
	var oHeaders = oMail.Headers;

	var headerString = ""
	var x = 0, y = 0;
	y = oHeaders.Count;
	for( x = 0; x < y; x++ )
	{
		var oHeader = oHeaders.Item(x);
		headerString += oHeader.HeaderKey + ":" + oHeader.HeaderValue + "\r\n";
	}
	WScript.Echo( headerString )
}

oClient.Quit();