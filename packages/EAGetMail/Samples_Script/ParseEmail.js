// ===============================================================================
// |    THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF      |
// |    ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO    |
// |    THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A         |
// |    PARTICULAR PURPOSE.                                                    |
// |    Copyright (c)2010  AdminSystem Software Limited                       |
// |
// |    Project: It demonstrates how to use EAGetMailObj to parse email summary in JScript
// |
// |    Author: Ivan Lui ( ivan@emailarchitect.net )
//  ===============================================================================
var info = "";
var args = WScript.Arguments;
if( args.Length < 1 )
{
	info =  "Usage: ParseEmail.js [email file path]\r\n";
	info = info + "   eg: ParseEmail.js \"d:\\1.eml\"";
	WScript.Echo( info );
	WScript.Quit();
}

var oMail = new ActiveXObject("EAGetMailObj.Mail");
//for evaluation usage, please use "TryIt", for licensed user, please use your license code here
oMail.LicenseCode = "TryIt";

oMail.LoadFile( args(0), false );
if( oMail.IsEncrypted )
{
	oMail = oMail.Decrypt( null );
}


if( oMail.IsSigned )
{	
	var oCert = oMail.VerifySignature();
}


var headerString = "";
headerString = "From: " + oMail.From.Name + "<" + oMail.From.Address + ">\r\n";
headerString += "Reply-To: " + oMail.ReplyTo.Name + "<" + oMail.ReplyTo.Address + ">\r\n"; 


headerString += "To:";
var ar = new VBArray(oMail.To).toArray();
var i = 0;
for( i = 0; i < ar.length; i++ )
{
	headerString += ar[i].Name + "<" + ar[i].Address + ">;\r\n";
}

headerString += "Cc:";
ar = new VBArray(oMail.Cc).toArray();
for( i = 0; i < ar.length; i++ )
{
	headerString += ar[i].Name + "<" + ar[i].Address + ">;\r\n";
}

headerString +=  "Subject: " + oMail.Subject + "\r\n";
headerString +=  "Date: " + oMail.ReceivedDate + "\r\n";

headerString += "Attachments: ";
ar = new VBArray(oMail.Attachments).toArray();
for( i = 0; i < ar.length; i++ )
{
	var oAtt = ar[i];
	headerString += oAtt.ContentType + ":" + oAtt.Name + ";\r\n";
}


WScript.Echo( headerString );
WScript.Echo( "" );
WScript.Echo( oMail.TextBody );

// display the htmlbody
// WScript.Echo( oMail.HtmlBody )

oMail.Clear();