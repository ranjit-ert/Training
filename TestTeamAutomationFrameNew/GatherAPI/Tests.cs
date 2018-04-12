using System;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium;
using NUnit.Framework;
using TestTeamAutomationFrameNew.Functions;
using TestTeamAutomationFrameNew.ITC_SMK;


namespace TestTeamAutomationFrameNew.GatherAPI
{
    class Tests
    {

        #region Setup

        public static string RunEnv = GenericFunctions.goAndGet("RUNENVIRONMENT");
        public static IWebDriver driver;


        [SetUp]
        public void SetUp()
        {
            ITC_SMK._startUp _startUp;
            _startUp = new ITC_SMK._startUp();
            var reporttime = DateTime.Now.ToString("dd-MM-yyyy_HH.mm");
            var ResultsDir = @"Results" + reporttime;
            var testFileName = "Results" + reporttime + ".html";


            // If the results directory doesn't exist, create it
            if (!System.IO.File.Exists(ResultsDir))
            {

                ITC_SMK._startUp.resultsFileName = testFileName;

            }
            else
            {
                Console.WriteLine("File already exists");
            }

        }

        #endregion


        //[Test]
        public void SmokeTest()
        {

            string Auth = null;
            string ApiKey = "9F5F459C-DAE4-4F98-AC7D-D79A38BD2F3A";
            string Username = "abrown";
            string Password = "Ir0nman>";
            string StudyName = "Temptemp";
            string SiteName = "";
            string SubjectID = "";
            // 8e45fa5e-ca1a-4e14-a264-829a48619db2


            // Setup the Reporting
            Reporting.ReportHeader("Gather API - SmokeTest");

            // Login
            Reporting.TestName("Send a valid Login Request");
            Auth = Functions.Login(Username, Password);


            // ActivateStudyConfiguration
            Reporting.TestName("Send a valid ActivateStudyConfiguration Request");
            Functions.ActivateStudyConfiguration(StudyName, ApiKey);


            // RetrieveStudyConfiguration
            Reporting.TestName("Send a valid RetrieveStudyConfiguration Request");
            Functions.GetStudyConfiguration(StudyName, ApiKey);


            // ListPatientsBySite
            Reporting.TestName("Send a valid ListPatientsBySite Request");
            Functions.ListPatientsBySite(StudyName, SiteName, ApiKey);


            // ActivateSubject
            Reporting.TestName("Send a valid ActivateSubject Request");
            Functions.ActivateSubject(StudyName, SiteName, SubjectID, ApiKey);


            // Teardown the Report
            Reporting.ReportFooter();

        }

    }

}
