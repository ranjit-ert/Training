using System;
using System.Data.SqlClient;
using System.Globalization;
using System.Threading;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using TestTeamAutomationFrameNew.Functions;
using System.Text.RegularExpressions;
using DropNet;
using System.IO;
using OpenQA.Selenium.Firefox;

namespace TestTeamAutomationFrameNew.ITC_SMK
{
    /// <summary>
    /// This is a class that contains all of the ITC tests
    /// </summary>
    public class _itcActions
    {
        
        public static void ITC_Login(IWebDriver driver)
        {

            string username = GenericFunctions.goAndGet("USERNAME");
            string password = GenericFunctions.goAndGet("PASSWORD");
            string studyName = GenericFunctions.goAndGet("STUDYID");

            // Login
            driver.Navigate().GoToUrl(GenericFunctions.goAndGet("UIUXURL"));
            GenericFunctions.CheckWindowTitle("Exco Intouch Online Services", driver);
            GenericFunctions.CheckTextIsOnPage("Online Services", driver);
            GenericFunctions.CheckTextIsOnPage("Enter your username and password:", driver);
            GenericFunctions.CheckTextIsOnPage("Username", driver);
            GenericFunctions.CheckTextIsOnPage("Password", driver);
            GenericFunctions.CheckTextIsOnPage("Knowledge Centre", driver);
            GenericFunctions.Type(username, "UserName", driver);
            GenericFunctions.Type(password, "Password", driver);
            Functions.Reporting.ReportScreenshot("Log", driver);
            GenericFunctions.ClickOnButton("logonBtn", driver);
            
   
            CheckPageText(driver);
            
        }

        public static int ITC_EnrolSubject(IWebDriver driver)
        {

            var studyName = GenericFunctions.goAndGet("STUDYID");
            var siteName = GenericFunctions.goAndGet("SITEID");
            var subjMob = GenericFunctions.goAndGet("SUBJECTPHONE");
            
           var subjectId = (_dbActions.CountSubjects(driver) == 0 ? Convert.ToInt32(GenericFunctions.goAndGet("SITEPREFIX") + "0001") : _dbActions.GetLatestSubjectId(studyName, siteName, driver));
            GenericFunctions.ClickOnText("enrolmentLink", driver);
            GenericFunctions.Wait(5);
            GenericFunctions.SelectDropDownByText(siteName, "siteId", driver);
            GenericFunctions.Wait(5);
            //GenericFunctions.Type("English", "patientLanguageDropDown", driver);
            GenericFunctions.Type("" + subjectId, "userIdentifier", driver);
            GenericFunctions.Type(subjectId + "-email@email.com", "SubjectEMail", driver);
            GenericFunctions.Type(subjMob, "SubjectMSISDN", driver);
            //GenericFunctions.Type("7700910000", "SubjectMSISDN", driver);
            GenericFunctions.Wait(5);
            
            var myDate = DateTime.Today.AddHours(23);

            datePicker(driver);

            //ItcTimePicker(2,myDate,driver);

            GenericFunctions.Type("In what city or town does your nearest sibling live?", "SecurityQuestionDropDownList", driver);
            GenericFunctions.Type("Nottingham","Patient_SecurityQuestionAnswer", driver);

            Functions.Reporting.ReportScreenshot("Reg",driver); ;  
            //GenericFunctions.Wait(5);
            GenericFunctions.waitforElement("submitBtn",driver);
            GenericFunctions.ClickOnButton("submitBtn", driver);

            GenericFunctions.Wait(2);

            GenericFunctions.CheckTextIsOnPage(subjectId + " has been successfully enrolled in the study.", driver);
            Functions.Reporting.ReportScreenshot("Enrol", driver);
            return subjectId;

        }

        public static void ITC_Withdraw(int SubjectID, IWebDriver driver)
        {

            string sEnv = GenericFunctions.goAndGet("ENVIRONMENT");
            int IntSubjectID;

            //This is no longer needed. WC
            IntSubjectID = _dbActions.GetRowIdForSubject(SubjectID, driver);
            
            GenericFunctions.ClickOnLink("Subject Administration", driver);
            
            GenericFunctions.Wait(10);
            //((IJavaScriptExecutor)driver).ExecuteScript("window.alert = function(){}");
            ((IJavaScriptExecutor)driver).ExecuteScript("window.confirm = function(){return true;}");
            GenericFunctions.ClickElement(driver.FindElement(By.XPath("//td[2]/a[@id='edit" + SubjectID + "']")),driver);
            GenericFunctions.Wait(5);
            Functions.Reporting.ReportScreenshot("Withdraw", driver);

            //IAlert alert = driver.SwitchTo().Alert();
            //alert.Accept(); //for two buttons, choose the affirmative one

            GenericFunctions.Wait(2);
            GenericFunctions.CheckTextIsOnPage("Subject " + SubjectID+ " has been successfully withdrawn from the study.", driver);
           
         }

        public static void ITC_ViewSubject(int SubjectID, IWebDriver driver)
        {
            string mySubId = SubjectID.ToString();
            string StudyName = GenericFunctions.goAndGet("STUDYID");
            string SiteName = GenericFunctions.goAndGet("SITEID");
            GenericFunctions.ClickOnLink("My Subjects", driver);
            GenericFunctions.Wait(12);
            Functions.Reporting.ReportScreenshot("View_Subjects", driver);

            //GenericFunctions.Type(mySubId, "subjectId", driver);
            //GenericFunctions.ClickOnButton("searchBtn", driver);
            //GenericFunctions.waitforElement(mySubId, driver);

            GenericFunctions.ClickElement(driver.FindElement(By.XPath("//*[@id=\"edit" + _dbActions.GetRowIdForSubject(SubjectID, driver) + "\"]")), driver);

            GenericFunctions.Wait(2);
            GenericFunctions.CheckTextIsOnPage("Subject Details", driver);
            GenericFunctions.CheckTextIsOnPage("Site", driver);
            GenericFunctions.CheckTextIsOnPage("Subject Identifier", driver);
            GenericFunctions.CheckTextIsOnPage("Language", driver);
            GenericFunctions.CheckTextIsOnPage("Email", driver);
            GenericFunctions.CheckTextIsOnPage("Event Schedule", driver);

            Functions.Reporting.ReportScreenshot("View_Subjectsched", driver);

        }

        public static void ITC_ViewSubSched(int SubjectID, IWebDriver driver)
        {

            ITC_ViewSubject(SubjectID, driver);
            GenericFunctions.ClickOnText("Event Schedule", driver);
            GenericFunctions.CheckTextIsOnPage("Subject Schedule", driver);
            GenericFunctions.CheckTextIsOnPage("Click schedule icons to view message details.", driver);
            ITC_CheckScheduleKey(driver);
            GenericFunctions.CheckTextIsOnPage("Baseline Schedule", driver);
            GenericFunctions.CheckTextIsOnPage("Non-Baseline Schedule", driver);
            GenericFunctions.ClickElement(driver.FindElement(By.XPath("//span[@class='ui-close']")),driver);
            Functions.Reporting.ReportScreenshot("View_Subject_Sched", driver); 
        }

        public static void ITC_CheckScheduleKey(IWebDriver driver)
        {
            GenericFunctions.CheckTextIsOnPage("+Show key", driver);
            GenericFunctions.ClickOnText("+Show key", driver);
            GenericFunctions.CheckTextIsOnPage("-Hide key", driver);
            GenericFunctions.CheckTextIsOnPage("Set In the Past", driver);
            GenericFunctions.CheckTextIsOnPage("Scheduled", driver);
            GenericFunctions.CheckTextIsOnPage("Send Within 24 hours", driver);
            GenericFunctions.CheckTextIsOnPage("Sent", driver);
            GenericFunctions.ClickOnText("-Hide key", driver);
            Functions.Reporting.ReportScreenshot("Schedule_key", driver);

        }

        public static void ITC_EditSubject(int SubjectID, IWebDriver driver)
        {

            string StudyName = GenericFunctions.goAndGet("STUDYID");
            string SiteName = GenericFunctions.goAndGet("SITEID");
            GenericFunctions.ClickOnButton("editBtn", driver);
            GenericFunctions.Wait(5);
            Functions.Reporting.ReportScreenshot("Edit_", driver);
            GenericFunctions.Clear("SubjectEMail", driver);
            GenericFunctions.Type(SubjectID + "@email-modified.com", "SubjectEMail", driver);

             //ClickElement(driver.FindElement(By.XPath("//*[@id=\"patientEditSaveButton\"]/span")));
            GenericFunctions.ClickElement(driver.FindElement(By.XPath("//*[@id=\"patientEditSaveButton\"]")),driver);
            GenericFunctions.Wait(20);

            var InputtedSecretQuestion = "In what city or town does your nearest sibling live?";
            var getSecretQuestion = driver.FindElement(By.Id("Patient_SecurityQuestion")).GetAttribute("value");
            //var getSecretQuestion = driver.FindElement(By.Id("SecurityQuestionDropDownList")).GetAttribute("value");

            var InputtedSecretQuestionAnswer = "Nottingham";
            var getSecretQuestionAnswer = driver.FindElement(By.Id("Patient_SecurityQuestionAnswer")).GetAttribute("value");

            if (getSecretQuestion.Equals(InputtedSecretQuestion) && (getSecretQuestionAnswer.Equals(InputtedSecretQuestionAnswer)))
            {
                GenericFunctions.ReportTheTestPassed("The secret question (" + getSecretQuestion + ") and secret answer (" + getSecretQuestionAnswer + ")  are correct ");
                _startUp.TestPasses++;
                _startUp.Passes++;
            }
            else
            {
                GenericFunctions.ReportTheTestFailed("The secret question is incorrect. The expected question is: " + InputtedSecretQuestion + "The actual question is: " + getSecretQuestion);
                GenericFunctions.ReportTheTestFailed("The secret question answer is incorrect. The expected question answer is: " + InputtedSecretQuestionAnswer + "The actual question is: " + getSecretQuestionAnswer);
                _startUp.TestFails++;
                _startUp.Fails++;

            }
            
            GenericFunctions.Wait(2);
            GenericFunctions.ClickOnLink("Subject Administration", driver);
            GenericFunctions.Wait(5);
            Functions.Reporting.ReportScreenshot("Edit", driver);
        }

        public static void ITC_Invalid(string IDORPASS, IWebDriver driver)
        {

            driver.Navigate().GoToUrl(GenericFunctions.goAndGet("UIUXURL"));
            GenericFunctions.CheckWindowTitle("Exco Intouch Online Services", driver);
            GenericFunctions.CheckTextIsOnPage("Online Services", driver);
            GenericFunctions.CheckTextIsOnPage("Enter your username and password:", driver);
            GenericFunctions.CheckTextIsOnPage("Username", driver);
            GenericFunctions.CheckTextIsOnPage("Password", driver);
            GenericFunctions.Type((IDORPASS == "Username" ? "INVALIDUSERNAME" : GenericFunctions.goAndGet("USERNAME")), "UserName", driver);
            GenericFunctions.Type((IDORPASS == "Password" ? "INVALIDPASSWORD" : GenericFunctions.goAndGet("PASSWORD")), "Password", driver);
            GenericFunctions.ClickOnButton("logonBtn", driver);
            GenericFunctions.CheckTextIsOnPage("Login was unsuccessful. Please correct the errors and try again.", driver);
            GenericFunctions.CheckTextIsOnPage("Your login attempt was not successful. Please try again", driver);

        }

        public static void datePicker(IWebDriver driver)
        {
            //created by someone who's long gone
            var myDate = DateTime.Today.ToString("dd/MMM");
            string myDate_ = DateTime.Today.ToString("dd");
            string mydate = "1";
            var removeZerofromDate = myDate_.TrimStart('0');
            Regex myMonths= new Regex("28/Feb|29/Feb|30/Sep|30/Apr|30/Jun|30/Nov|31/Jan|31/Mar|31/May|31/Jul|31/Aug|31/Oct|31/Dec");
           
            if (myMonths.IsMatch(myDate))
            {
                driver.FindElement(By.CssSelector("span.t-icon.t-icon-calendar")).Click();
                Thread.Sleep(2000);
                driver.FindElement(By.CssSelector("span.t-icon.t-arrow-next")).Click(); //Go to the next month
                driver.FindElement(By.LinkText(mydate)).Click();//select the 1st day 

            }
            else    
            {
               
                driver.FindElement(By.CssSelector("span.t-icon.t-icon-calendar")).Click();
                Thread.Sleep(2000);
                driver.FindElement(By.LinkText(removeZerofromDate)).Click();
              }
          }
       
        public static void ItcTimePicker(int iIndex, DateTime dDateTime, IWebDriver driver)
        {
          ////Work out the required time by rounding the time up to the nearest 5 minute increment
          //  var endTime = Convert.ToDateTime(Convert.ToDateTime(GenericFunctions.RoundUp(DateTime.Parse(dDateTime.ToShortTimeString()), TimeSpan.FromMinutes(5))).ToShortTimeString());
          // //Work out the dropdown index number for the required time by calculating the difference in minutes from 00:00:00 then dividing it by 5 and adding 1
          //  //it's divided by 5 as thats the increments of this dropdown and the time was rounded to the nearest 5 minutes
          //  //you add 1 as the index starts at 1 and 00:00:00 calculates to a 0 index number
          //  var startTime = Convert.ToDateTime("12:00 AM");
          //  var span = endTime.Subtract(startTime);
          //  int myIndex = (Convert.ToInt32(span.TotalMinutes.ToString(CultureInfo.InvariantCulture)) / 5) + 1;
          ////click on the time picker
          //  var vWait = new WebDriverWait(driver, TimeSpan.FromSeconds(20));
          //  vWait.Until(ExpectedConditions.ElementExists(By.XPath("//*[@class='editor-field short'][" + iIndex.ToString(CultureInfo.InvariantCulture) + "]/div/div/span/span[@class='t-icon t-icon-clock']")));
          //GenericFunctions.Wait(0.5);
          //  driver.FindElement(By.XPath("//*[@class='editor-field short'][" + iIndex.ToString(CultureInfo.InvariantCulture) + "]/div/div/span/span[@class='t-icon t-icon-clock']")).Click();
          // //click on the time
          //  vWait = new WebDriverWait(driver, TimeSpan.FromSeconds(20));
          //  vWait.Until(ExpectedConditions.ElementExists(By.XPath("//*[@class='t-animation-container' and contains(@style, 'block')]/div/ul/li[" + myIndex.ToString() + "]")));
          //  GenericFunctions.Wait(0.5);
          //  //Scroll so that all of the dropdown is in view.  this is to stop it crashing when dropdown goes off the bottom of the screen.
          //  System.Drawing.Point point = ((OpenQA.Selenium.Remote.RemoteWebElement)driver.FindElement(By.XPath("//*[@class='t-animation-container' and contains(@style, 'block')]/div/ul/li[" + myIndex.ToString() + "]"))).LocationOnScreenOnceScrolledIntoView;
          //  driver.FindElement(By.XPath("//*[@class='t-animation-container' and contains(@style, 'block')]/div/ul/li[" + myIndex.ToString(CultureInfo.InvariantCulture) + "]")).Click();
         }

      public static void CheckPageText(IWebDriver driver)
        {

            GenericFunctions.CheckWindowTitle("InTouch Clinical", driver);
            GenericFunctions.CheckTextIsOnPage("InTouch Clinical", driver);
            GenericFunctions.CheckTextIsOnPage("Welcome " + GenericFunctions.goAndGet("USERNAME"), driver);
            GenericFunctions.CheckTextIsOnPage("Study: " + GenericFunctions.goAndGet("STUDYID"), driver);
            GenericFunctions.CheckTextIsOnPage("Log Off", driver);
            GenericFunctions.CheckTextIsOnPage("Subject Administration", driver);
            GenericFunctions.CheckTextIsOnPage("Reporting", driver);
            GenericFunctions.CheckTextIsOnPage("Subjects", driver);

        }

        public static void ITC_CheckHelpText(IWebDriver driver)
        {
            var studyName = GenericFunctions.goAndGet("STUDYID");


            using (var conn = new SqlConnection(GenericFunctions.goAndGet("DBCONNECTION")))
            {

                conn.Open();

                string helpText;
                using (var command = new SqlCommand("select help_text from tbl_study where name = '" + studyName + "'"))
                {
                    command.CommandType = System.Data.CommandType.Text;
                    command.Connection = conn;

                    var obj = command.ExecuteScalar();
                    helpText = obj.ToString();
                    helpText = helpText.Replace("<b>", "");
                    helpText = helpText.Replace("</b>", "");
                    helpText = helpText.Replace("<u>", "");
                    helpText = helpText.Replace("</u>", "");
                    helpText = helpText.Replace("<i>", "");
                    helpText = helpText.Replace("</i>", "");
                    helpText = helpText.Replace("</b>", "");
                    //help_text = help_text.Replace("<br>", "\n");
                    helpText = helpText.Replace("<center>", "");
                    helpText = helpText.Replace("</center>", "");

                }

                conn.Close();
                GenericFunctions.ClickElement(driver.FindElement(By.XPath("//div/a/img")),driver);
                GenericFunctions.CheckTextIsOnPage(helpText, driver);
                GenericFunctions.ClickElement(driver.FindElement(By.XPath("//span[@class='ui-icon ui-icon-closethick']")),driver);

            }

        }

        public static void ITC_AlertEmail(IWebDriver driver)
        {
            GenericFunctions.ClickOnText("scheduleFeedbackMenu", driver);
            GenericFunctions.CheckTextIsOnPage("Alert Email Followup", driver);
            GenericFunctions.CheckTextIsOnPage("Select Site", driver);
            GenericFunctions.CheckTextIsOnPage("Select Subject", driver);
            GenericFunctions.CheckTextIsOnPage("Alert Email Type", driver);
            Functions.Reporting.ReportScreenshot("AlertEmail", driver);
        }

        public static void ITC_CheckLastEventDate(int sub, string expectedDate, IWebDriver driver)
        {
            ////Search for subject
            
            string subID = sub.ToString();
            GenericFunctions.Wait(2);
            GenericFunctions.ClickOnLink("My Subjects", driver);
            GenericFunctions.waitforElement("subjectId", driver);
            driver.FindElement(By.Id("subjectId")).Clear();
            driver.FindElement(By.Id("subjectId")).SendKeys(subID);
            GenericFunctions.Wait(2);
            GenericFunctions.ClickOnButton("searchBtn", driver);
            GenericFunctions.CheckTextIsOnPage(subID, driver);
            GenericFunctions.Wait(2);
            //Get last event date
            string EventDate = driver.FindElement(By.XPath("//*[@id='Grid']/div[2]/table/tbody/tr/td[6]")).Text;
            

            Functions.Reporting.ReportScreenshot("AlertEmail", driver);

            //Compare actual last event date to expected last event date
            if (expectedDate.Equals(EventDate))
            {
               GenericFunctions.ReportTheTestPassed("The last event date (" + EventDate + ") is correct");
                _startUp.TestPasses++;
                _startUp.Passes++;
            }
            else
            {
                GenericFunctions.ReportTheTestFailed("The last event date is incorrect. The expected date is: " + expectedDate + " The actual date is: " + EventDate);
                _startUp.TestFails++;
                _startUp.Fails++;
            }
        }

        public void ITC_LogOff(IWebDriver driver)
        {

            GenericFunctions.ClickElement(driver.FindElement(By.LinkText("Log Off")),driver);

        }
    }
}
