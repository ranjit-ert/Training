using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.IO;
using NUnit.Framework;
using TestTeamAutomationFrameNew.Functions;
using System.Data.SqlClient;
using System.Data;
using System.Net.Mail;
using System.Net;
using System.Web;
using Microsoft.Office.Interop.Outlook;


namespace TestTeamAutomationFrameNew.SendEmail
{
    public class SendEmail
    {

        public static void sendEmail(string url) //outlook
        {
          
            // Create the Outlook application.
            Microsoft.Office.Interop.Outlook.Application oApp;
                        oApp = new Microsoft.Office.Interop.Outlook.Application();

            // Create a new mail item.
            Microsoft.Office.Interop.Outlook.MailItem oMsg;
                        oMsg = (Microsoft.Office.Interop.Outlook.MailItem)oApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);

                        Microsoft.Office.Interop.Outlook.Recipients oRecips;
                         Microsoft.Office.Interop.Outlook.Recipient oRecip;
                         string sourcePath = AppDomain.CurrentDomain.BaseDirectory.ToString();
                        string imageCid = "image001.jpg@123"; //can be any string
            // Set HTMLBody. 
            //add the body of the email
           // oMsg.HTMLBody = url;


            //Add an attachment.
            //String attach = "Attachment to add to the Mail";
            //int x = (int)oMsg.Body.Length + 1;
            //int y = (int)Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue;

            ///Attach the File here
            // Microsoft.Office.Interop.Outlook.Attachment oAttach = oMsg.Attachments.Add(FileLocation, y, x, attach); //no longer needed as we are using dropbox


            string error = "Failed";
            //string inc = "Inconclusive";
            if (url.Equals(error))// || (url.Equals(inc))))
            {
               
                    Microsoft.Office.Interop.Outlook.Attachment attachments = oMsg.Attachments.Add(sourcePath + "monkey.tif",
                    Microsoft.Office.Interop.Outlook.OlAttachmentType.olEmbeddeditem, null, "Some image display name");

                    attachments.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imageCid);
                    oMsg.HTMLBody = String.Format(
                      "<body><img src=\"cid:{0}\"><br> <b>Oh</b> <b>Oh</b> seems like something has gone wrong, don't worry our trained <strike>monkeys</strike> technicians are busy <strike>googling</strike> resolving the situation </br></body>"
                      , imageCid);
                    oMsg.Subject = "Smoke Test Automation Error";

                    oRecips = (Microsoft.Office.Interop.Outlook.Recipients)oMsg.Recipients;


                    oRecip = (Microsoft.Office.Interop.Outlook.Recipient)//oRecips.Add("PM@excointouch.com");
                    //oRecips.Add("testteam@excointouch.com"); oRecips.Add("supportstaff@excointouch"); oRecips.Add("dev@excointouch.com");
                    oRecips.Add("testteam@excointouch.com");
                    oRecip.Resolve();

                    //send the mail using send method
                    oMsg.Send();
                
            }
            else
            {
                //embed item
             
                    Microsoft.Office.Interop.Outlook.Attachment attachment = oMsg.Attachments.Add(sourcePath + "fonz.tif",
                    Microsoft.Office.Interop.Outlook.OlAttachmentType.olEmbeddeditem, null, "Some image display name");

                    attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imageCid);
                    oMsg.HTMLBody = String.Format(
                      "<body><img src=\"cid:{0}\"><br> The fonz says the Engage Automation Smoke Test has <b><font color=green>passed</font></b>, please click on the link below </br><br>" + url + "</br></body>"
                     , imageCid
                     );



                    ///Add the Subject of mail item
                    oMsg.Subject = "Automated PDF Report of Engage Smoke Test Automation";

                    //Add the mail id to which you want to send mail.

                    oRecips = (Microsoft.Office.Interop.Outlook.Recipients)oMsg.Recipients;
                    oRecip = (Microsoft.Office.Interop.Outlook.Recipient) //oRecips.Add("PM@excointouch.com");
                        //oRecips.Add("testteam@excointouch.com"); oRecips.Add("supportstaff@excointouch"); oRecips.Add("dev@excointouch.com");
                    oRecips.Add("testteam@excointouch.com");
                    oRecip.Resolve();

                    //send the mail using send method
                    oMsg.Send();
                }
              
           
           
        }
        //[Test]
        public static void SendeMail(string url) //Net Mail
        {
            var embeddedImage = "C:\\Users\\Wonderson.Chideya\\Downloads\\fonz.jpg";
            //Holds message information.
            System.Net.Mail.MailMessage mailMessage = new MailMessage();
            //Add basic information.
            mailMessage.From = new System.Net.Mail.MailAddress("testteam@excointouch.com");
            mailMessage.To.Add("testteam@excointouch.com");
            mailMessage.Subject = "mail title";
            //Create two views, one text, one HTML.
            System.Net.Mail.AlternateView htmlView = System.Net.Mail.AlternateView.CreateAlternateViewFromString("your content" + "<image src=cid:HDIImage>", null, "text/html");
            //Add image to HTML version

            //System.Net.Mail.LinkedResource imageResource = new System.Net.Mail.LinkedResource(fileImage.PostedFile.FileName);
            // System.Net.Mail.LinkedResource imageResource = new System.Net.Mail.LinkedResource(@"C:\Users\v-fuzh\Desktop\apple.jpg");

            System.Net.Mail.LinkedResource imageResource = new System.Net.Mail.LinkedResource(embeddedImage);
            imageResource.ContentId = "HDIImage";
            htmlView.LinkedResources.Add(imageResource);
            //Add two views to message.
            mailMessage.AlternateViews.Add(htmlView);
            //Send message
            System.Net.Mail.SmtpClient smtpClient = new System.Net.Mail.SmtpClient();
            smtpClient.Send(mailMessage);
        }
    }
}

   
       

