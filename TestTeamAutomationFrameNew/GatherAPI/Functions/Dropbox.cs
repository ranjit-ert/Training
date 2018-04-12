using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DropNet;
using System.IO;
using TestTeamAutomationFrameNew.SendEmail;

namespace TestTeamAutomationFrameNew.Functions
{
    public class Dropbox
    {
        
        public static void uploadFile(string fileName)
        {
          
            
            /// You'll need to go on this following site "https://www.dropbox.com/developers/apps/create" and select Dropbox API app
            var _appKey = "i1lao88ewl6jjqf"; //You'll get these once you've created you app in dropbox
            var _appSecret = "s3538tkq57rfzrz"; //You'll get these once you've created you app in dropbox
            var _accessToken = "f3fhn9gjsxh1zugl";
            var _accesTokenSecret = "lw0hfw3zvk13qr1";
            System.Net.IWebProxy proxy = null; //If there's a proxy put it here
            var _client = new DropNetClient(_appKey, _appSecret, _accessToken, _accesTokenSecret, proxy);

            //_client.GetToken();
            //_client.UserLogin = token;
            //var url = _client.BuildAuthorizeUrl(); //need to login to get the _accessToken and _accessTokenSecret
            //var log = _client.GetAccessToken();
            //var f = log.Token;
            //var y = log.Secret;

            _client.UseSandbox = true; //allow access to root folder
            var content = File.ReadAllBytes(@"" + fileName + ""); //read the bytes for the file that you wish to upload
            var uploaded = _client.UploadFile("/", fileName, content); //upload

            var sharefile = _client.GetShare("/" + fileName + ""); //Share the file

            SendEmail.SendEmail.sendEmail(sharefile.Url);
        }
    }
}


