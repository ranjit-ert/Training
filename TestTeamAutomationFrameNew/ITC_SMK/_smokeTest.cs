using System;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.PhantomJS;
using OpenQA.Selenium.Remote;
using TestTeamAutomationFrameNew.Functions;
using System.IO;
using System.Threading;
using System.Linq;
using System.Diagnostics;
using SelectPdf;




namespace TestTeamAutomationFrameNew.ITC_SMK
{
    /// <summary>
    /// This is the class where the Lowlevel tests are referenced
    /// </summary>
    //[TestFixture(typeof(ChromeDriver))] //Pass in ChromeDriver
    //[TestFixture(typeof(FirefoxDriver))] //Pass in FirefoxDriver
    //[TestFixture(typeof(InternetExplorerDriver))] //Pass in FirefoxDriver
    //public class _smokeTest<TWebDriver> where TWebDriver : IWebDriver, new() //Run multiple browsers 

    public class HeadlessSmokeTest//Run multiple browsers 
    {
        readonly _startUp _startUp;

        public HeadlessSmokeTest()
        {
            _startUp = new _startUp();

        }

        [SetUp]

        public void SetUp()
        {
            var reporttime = DateTime.Now.ToString("dd-MM-yyyy_HH.mm");
            var ResultsDir = @"Results" + reporttime;
            var testFileName = "Results" + reporttime + ".html";

            if (!System.IO.File.Exists(ResultsDir))
            {
                //System.IO.File.Create(ResultsDir);
                _startUp.resultsFileName = testFileName;

            }
            else
            {
                Console.WriteLine("File already exists");//_startUp.resultsFileName = System.IO.Path.Combine(ResultsDir, testFileName);
            }
            //_startUp.driver = new TWebDriver();
            // _startUp.driver.Manage().Window.Maximize();
            //var caps = DesiredCapabilities.PhantomJS();
            //caps.SetCapability("phantomjs.page.settings.userAgent", "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:18.0) Gecko/20100101 PhantomJS/1.9.7 Firefox/18.0");

            _startUp.driver = new FirefoxDriver();
            //_startUp.driver = new PhantomJSDriver();
            _startUp.driver.Manage().Window.Maximize();
            _startUp.Setup();

        }


        //[Test]
        public void SmokeTest()
        {

            Reporting.ReportHeader("Engage Smoke Test");
            // Login
            Reporting.TestName("Login");
            _itcActions.ITC_Login(_startUp.driver);

            // Enrol Subject
            Reporting.TestName("Enrol");
            int subject1 = _itcActions.ITC_EnrolSubject(_startUp.driver);

            //check last event date
            Reporting.TestName("Check last event date");
            var LastEvent = DateTime.Today.AddDays(4).ToString("dd MMMM yyyy");
            _itcActions.ITC_CheckLastEventDate(subject1, LastEvent, _startUp.driver);

            // View Subject
            Reporting.TestName("View Subject");
            _itcActions.ITC_ViewSubject(subject1, _startUp.driver);

            // Check the Help Text
            //_itcActions.ITC_CheckHelpText(driver); //currently there is no help text displayed on the site so need to run test

            // Edit Subject
            Reporting.TestName("Edit Subject");
            _itcActions.ITC_EditSubject(subject1, _startUp.driver);

            //Add a test to view scheduled messages

            // Withdraw
            Reporting.TestName("Withdraw");
            _itcActions.ITC_Withdraw(subject1, _startUp.driver);

            Reporting.ReportFooter();

            //Delete all JPEGS in folder
            string sourcePath = AppDomain.CurrentDomain.BaseDirectory.ToString();
            string htmlLink = "file:///" + sourcePath + _startUp.resultsFileName; //locate the file
            string fileName = "_TestReport" + _startUp.resultsFileName +".pdf";
            HtmlToPdf conver = new HtmlToPdf(); //Initialisethe object
            PdfDocument doc = conver.ConvertUrl(htmlLink); // convert the link, you can convert raw HTML if you wish
            doc.Save(fileName); //save the file
            doc.Close(); //close

            Dropbox.uploadFile(fileName); //upolad to dropbox

            
            _startUp.Teardown(_startUp.driver);
        }


        //[Test]
        public void WhiteBeltTest()
        {

            _startUp.driver.Navigate().GoToUrl("https://test.reactintouch.com");
            GenericFunctions.ClickOnButton("ctl00_middlecontent_Login1_LoginButton", _startUp.driver);


        }


        //[Test]
        public void WhiteBeltTraining()
        {


            _startUp.driver.Navigate().GoToUrl("https://test.itc.excointouch.com/");
            GenericFunctions.CheckTextIsOnPage("Enter your username and password:", _startUp.driver);
            GenericFunctions.CheckTextIsNotOnPage("Reports", _startUp.driver);
            GenericFunctions.CheckTextIsOnPage("Username", _startUp.driver);
            GenericFunctions.Type("abrown", "UserName", _startUp.driver);
            GenericFunctions.Type("Password", "Password", _startUp.driver);
            GenericFunctions.ClickOnButton("logonBtn", _startUp.driver);
            GenericFunctions.SelectFromList("DEMO1", "studyId", _startUp.driver);
            GenericFunctions.CheckTextIsOnPage("Subject Enrolment", _startUp.driver);
            GenericFunctions.ClickOnLink("Reporting", _startUp.driver);


          
    
        }










        //[Test]

        //public void Dothistest()
        //{
        //    ///Not needed
        //    //string strCmd = @"C:\phantomjs-2.0.0-windows\bin";
        //    //string args = "/c phantomjs github.js"; //need to put /c to tell command prompt your request
        //    //Process p = new Process();
        //    //p.StartInfo.WorkingDirectory = strCmd;
        //    //p.StartInfo.Arguments = args;
        //    //p.StartInfo.FileName = "cmd.exe";
        //    //p.Start();

        //    string sourcePath = AppDomain.CurrentDomain.BaseDirectory.ToString();
        //    string htmlLink = "file:///" + sourcePath +  "Results05-03-2015_14.21.html";
        //    string fileName = "_TestReport" + _startUp.resultsFileName +".pdf";
        //    HtmlToPdf conver = new HtmlToPdf();
        //    try
        //    {
        //        PdfDocument doc = conver.ConvertUrl(htmlLink);
        //        doc.Save(fileName);
        //        doc.Close();
        //    }
        //    catch (Exception e)
        //    {
        //    }


        //   //Dropbox.uploadFile(fileName);

        //    _startUp.Teardown(_startUp.driver);


        //      }



        //}


    }
    }

