using System.Collections.Generic;
using NUnit.Framework;
using OpenQA.Selenium;
using TestTeamAutomationFrameNew.Functions;
using System.Linq;
using System;
using System.IO;

namespace TestTeamAutomationFrameNew.ITC_SMK
{
   
    public class _startUp
    {
        public IWebDriver driver;

        public static string testSummaryFileName;
        public static  string resultsFileName;
        public static int stepsFailed = 0;
        public static int TestsPassed = 0;
        public static string todaysDate;
        public List<string> subjectids = new List<string>();
        public string BROWSER = "";
        public static int Passes = 0;
        public static int Fails = 0;
        public static int Warnings = 0;
        public static int TestPasses = 0;
        public static int TestFails = 0;
        public static int TestWarnings = 0;
        public static int TestID = 1;
        public static string Contents;
        public static string RunEnv;

        [SetUp]
        public void Setup()
        {
          //getSubjectId();
            todaysDate = System.DateTime.Now.ToString("yyyyMMddhhmm");
            todaysDate = todaysDate.Substring(0, todaysDate.Length - 4);
            todaysDate = todaysDate + "0001";
            //BROWSER = "FF";
        }

        [TearDown]
        public void Teardown(IWebDriver driver)
        {
            //if (TestContext.CurrentContext.Result.Status == TestStatus.Inconclusive)
            //{
            //   GenericFunctions.GetScreenShot(driver, "ITC_SMK_Error");
            //   SendEmail.SendEmail.sendEmail(TestContext.CurrentContext.Result.Status.ToString());
            //}
            string sourcePath = AppDomain.CurrentDomain.BaseDirectory.ToString();
            if (System.IO.Directory.Exists(sourcePath))
            {
                string supportedExtensions = "*.jpg,*.html, *.pdf";
                foreach (string imageFile in Directory.GetFiles(sourcePath, "*.*", SearchOption.AllDirectories).Where(s => supportedExtensions.Contains(Path.GetExtension(s).ToLower())))
                {
                    if (imageFile.EndsWith(supportedExtensions))
                    {

                    }

                    File.Delete(imageFile);
                }

                
            }
        
           driver.Quit();
        }

        

        public void Open(string webpage)
        {
            driver.Navigate().GoToUrl(webpage);
        }

    }
}

