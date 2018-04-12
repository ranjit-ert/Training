using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.IE;
using NUnit.Framework;
using OpenQA.Selenium.Support.Events;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System.Threading;
//using Microsoft.VisualBasic.FileIO;
using System.IO;
using System.Data;
using System.Drawing;
using TestTeamAutomationFrameNew.Functions;


namespace TestTeamAutomationFrameNew
{
    public class ItcEnrol
    {

        //public static Boolean ITClogin(string sEnvironment, string sUsername, string sPassword, IWebDriver _driver)
        //{
        //    Boolean bResult = false;
        //    //Work out the URL for ITC and open it
        //    string sItcUrl = GenericFunctions.goAndGet("UIUXURL", sEnvironment);
        //    _driver.Navigate().GoToUrl(sItcUrl);
        //    GenericFunctions.TakePicture(_driver); 

        //    Thread.Sleep(1000);
        //    PageCheck.CaptureVisibleElements("pcItcLogin.csv", _driver);
        //    GenericFunctions.TakePicture(_driver); 
        //    //Enter the username and password
        //    if (GenericFunctions.EnterTextBox("ID", "UserName", sUsername, 20, _driver) == false)
        //    {   //Failed to enter username - abort
        //        Console.WriteLine("ITClogin Finished Function:  bResult={0}  [WARNING] Unable to enter username (sURL={1})", bResult, sItcUrl);
        //        return bResult;
        //    }
        //    GenericFunctions.TakePicture(_driver); 
        //    if (GenericFunctions.EnterTextBox("ID", "Password", sPassword, 20, _driver) == false)
        //    {   //Failed to enter username - abort
        //        Console.WriteLine("ITClogin Finished Function:  bResult={0}  [WARNING] Unable to enter password (sURL={1})", bResult, sItcUrl);
        //        return bResult;
        //    }
        //    //Click 'Log On'
        //    GenericFunctions.TakePicture(_driver); 
        //    if (GenericFunctions.ClickSomething("ID", "logonBtn", 20, _driver) == false)
        //    {   //Failed to enter username - abort
        //        Console.WriteLine("ITClogin Finished Function:  bResult={0}  [WARNING] Unable to click Login (sURL={1})", bResult, sItcUrl);
        //        return bResult;
        //    }
        //    //Check it logged in successfully
        //    GenericFunctions.TakePicture(_driver); 
        //    if (GenericFunctions.CheckElementExists("XPATH", "//*[@id='logindisplay']", 20, _driver) == false)
        //    {   //Logout link NOT found, which means it failed to login - abort
        //        Console.WriteLine("ITClogin Finished Function:  bResult={0}  [WARNING] Login failed (sURL={1})", bResult, sItcUrl);
        //        return bResult;
        //    }
        //    else
        //    {   //Logout link found, which means it's successfully logged in
        //        bResult = true;
        //        Console.WriteLine("ITClogin Finished Function:  bResult={0}, sURL={1}", bResult, sItcUrl);
        //        return bResult;
        //    }
        //}



        //public static bool ITClogout(IWebDriver _driver)
        //{
        //    //if (_genericFunctions.ClickSomething("XPATH", "//a[contains(@href, 'LogOff')]", 20, _driver) == false)
        //    //{
        //    //    Console.WriteLine("ITClogout = false [WARNING] Failed to log out of ITC");
        //    //    return false;
        //    //}
        //    //else
        //    //{
        //    //    Console.WriteLine("ITClogout = true");
        //    //    return true;
        //    //}

        //    if (GenericFunctions.ClickSomething("XPATH", "Log Off", 20, _driver) == false)
        //    {
        //        Console.WriteLine("ITClogout = false [WARNING] Failed to log out of ITC");
        //        return false;
        //    }
        //    else
        //    {
        //        Console.WriteLine("ITClogout = true");
        //        return true;
        //    }


        //}



        //public static bool ITCselectStudy(string sStudy, IWebDriver _driver)
        //{

        //    if (GenericFunctions.SelectDropDown("studyId", sStudy, 20, _driver) == true)
        //    {
        //        //Success - study selected in dropdown
        //        GenericFunctions.TakePicture(_driver); 
        //       return true;
        //    }
        //    else
        //    {
        //        //Failed to select study from dropdown
        //        //Check if study already selected (user only assigned to one study)
        //        if (GenericFunctions.CheckElementExists("XPATH", "//*[@id='studyForm']/em", 20, _driver) == true)
        //        {
        //            //Study lable exists - check the value
        //            string sSelected = _driver.FindElement(By.XPath("//*[@id='studyForm']/em")).Text.ToString();
        //          if (sSelected != sStudy)
        //            {
        //                //Failed to select study - return fail
        //                GenericFunctions.TakePicture(_driver); 
        //                return false;
        //            }
        //            else
        //            {
        //                //Study already selected - return success
        //                GenericFunctions.TakePicture(_driver); 
        //                return true;
        //            }
        //        }
        //        else
        //        {
        //            //Study lable not found - return fail
        //            GenericFunctions.TakePicture(_driver); 
        //            return false;
        //        }
        //    }
        //}


        //public static Boolean EnrolSubject(string sSite, string sLanguage, string sSubjectID, string sFileName, IWebDriver _driver)
        //{
        //    Console.WriteLine("EnrolSubject:  Start Function");
        //    Console.WriteLine("  sSite:  " + sSite);
        //    Console.WriteLine("  sLanguage:  " + sLanguage);
        //    Console.WriteLine("  sSubjectID:  " + sSubjectID);
        //    Console.WriteLine("  sFileName:  " + sFileName);
        //    Boolean bResult = false;

        //    WaitForLoadingToComplete(_driver);

        //    //Click 'Enrol Subject'
        //    if (GenericFunctions.ClickSomething("ID", "enrolmentLink", 30, _driver) == false)
        //    {
        //        //Failed to enter username - abort
        //        Console.WriteLine("  [FAIL]:  Unable to click Subject Enrollment link");
        //        Console.WriteLine("EnrolSubject:  Finished Function");
        //        return bResult;
        //    }
        //    Thread.Sleep(500);
        //    Console.WriteLine("  Clicked Enrol Subject");

        //    //Wait for enrollment form to open
        //    WaitForLoadingToComplete(_driver);
        //    if (GenericFunctions.CheckElementExists("ID", "enrol-patient", 120, _driver) == false)
        //    {
        //        //Failed to enter username - abort
        //        Console.WriteLine("  [FAIL]:  Enrollment form not opened");
        //        Console.WriteLine("EnrolSubject:  Finished Function");
        //        return bResult;
        //    }
        //    Console.WriteLine("   - Subject enrolment form opened");

        //    //Select a Site if site is required
        //    if (GenericFunctions.CheckElementExists("ID", "siteId", 1, _driver) == true)
        //    {
        //        //Site is required
        //        if (sSite == null || sSite == "")
        //        {
        //            //Site is required, but no site was supplied - abort subject enrollment
        //            bResult = false;
        //            Console.WriteLine("  [ERROR] Site is required, but no site was supplied - subject enrollment aborted");
        //            Console.WriteLine("EnrolSubject:  Finished Function");
        //            return bResult;
        //        }
        //        //Select the required site
        //        if (GenericFunctions.SelectDropDown("siteId", sSite, 20, _driver) == false)
        //        {
        //            //Failed to select study - abort
        //            Console.WriteLine("  [FAIL] Enrolment unsuccessful - Test aborted");
        //            Console.WriteLine("EnrolSubject:  Finished Function");
        //            return bResult;
        //        }
        //        //Site successfully selected
        //        Console.WriteLine("  Site '" + sSite + "' successfully selected.");
        //        Thread.Sleep(3000);
        //    }
        //    else
        //    {
        //        //Site NOT required\not visible
        //        if (sSite != null & sSite != "")
        //        {
        //            //Site is NOT required, but a site was supplied - raise a warning
        //            Console.WriteLine("  [WARNING] Site is NOT required, but a site was supplied.  Site requested = " + sSite);
        //            Console.WriteLine("    - Enrollment will continue, but the site the subject will be enrolled into is unknown.");
        //        }
        //    }

        //    //Select a Language if language is required
        //    ;
        //    //if (_genericFunctions.CheckElementExists("ID", "Patient_LanguageId", 20, _driver) == true)

        //    if (_driver.FindElement(By.XPath("//*[@id='patientLanguageDropDown']")).Displayed == true) //Patient_LanguageId - old language dropdown id

        //    {
        //        //Language is required
        //        if (sLanguage == null || sLanguage == "")
        //        {
        //            //Language is required, but no language was supplied - abort subject enrollment
        //            bResult = false;
        //            Console.WriteLine("  [ERROR] Language is required, but no language was supplied - Subject enrollment ABORTED");
        //            Console.WriteLine("EnrolSubject:  Finished Function");
        //            return bResult;
        //        }
        //        //Select the language

        //      if (GenericFunctions.SelectDropDown("patientLanguageDropDown", sLanguage, 20, _driver) == false) //Patient_LanguageId

        //        {
        //            //Failed to select study - abort
        //            Console.WriteLine("  [FAIL] Enrolment unsuccessful - Test aborted");
        //            Console.WriteLine("EnrolSubject:  Finished Function");
        //            return bResult;
        //        }
        //        Console.WriteLine("  Entered Language in language dropdown:  " + sLanguage);
        //        Thread.Sleep(500);
        //    }
        //    else
        //    {
        //        //Language NOT required
        //        if (sLanguage != null & sLanguage != "")
        //        {
        //            //Language is NOT required, but a language was supplied - raise a warning
        //            Console.WriteLine("  [WARNING] Language is NOT required, but a language was supplied.  Language requested = " + sLanguage);
        //            Console.WriteLine("    - Enrollment will continue, but the language the subject will be enrolled with is unknown.");
        //        }
        //    }

        //    //Enter a Subject ID
        //    if (GenericFunctions.EnterTextBox("ID", "userIdentifier", sSubjectID, 20, _driver) == false)
        //    {
        //        //Failed to enter username - abort
        //        Console.WriteLine("  [FAIL]:  Unable to enter subject ID");
        //        Console.WriteLine("EnrolSubject:  Finished Function");
        //        return bResult;
        //    }
        //    Console.WriteLine("  Entered Subject ID:  " + sSubjectID);

        //    //Load data from CSV and enter the required dat in the enrollment form
        //    string[,] saValues = GenericFunctions.LoadCsv(sFileName);
        //    int iRows = saValues.GetUpperBound(0);
        //    //Console.WriteLine("iRows:  " + iRows.ToString());
        //    int iCols = saValues.GetUpperBound(1);
        //    //Console.WriteLine("iCols:  " + iCols.ToString());
        //    for (int r = 1; r <= iRows; r++)
        //    {
        //        string sField = saValues[r, 0].ToString();
        //        Console.WriteLine("  sField:  " + sField);
        //        string sType = saValues[r, 1].ToString();
        //        Console.WriteLine("  sType:  " + sType);
        //        string sID = saValues[r, 2].ToString();
        //        Console.WriteLine("  sID:  " + sID);
        //        string sValue = saValues[r, 3].ToString();
        //        Console.WriteLine("  sValue:  " + sValue);

        //        DateTime dEvent;
        //        int iDateBoxIndex;
        //        switch (sType.ToUpper())
        //        {
        //            case "TEXTBOX":
        //                //Enter TextBox data
        //                if (GenericFunctions.EnterTextBox("ID", sID, sValue, 20, _driver) == false)
        //                {
        //                    //Failed to enter something - raise warning and carry on
        //                    Console.WriteLine("  [WARNING]:  Unable to enter data for field " + sField + "):  " + sValue);
        //                    Console.WriteLine("EnrolSubject:  Finished Function");
        //                    return bResult;
        //                }
        //                Console.WriteLine("  Entered TextBox Data (" + sField + "):  " + sValue);
        //                break;
        //            case "DATETIME":
        //                dEvent = GenericFunctions.GetDate(sValue);
        //                iDateBoxIndex = Convert.ToInt32(sID);
        //                Console.WriteLine("  dEvent.ToLongDateString() = " + dEvent.ToLongDateString());
        //                Console.WriteLine("  DateTime.Today.ToLongDateString() = " + DateTime.Today.ToLongDateString());
        //                if (dEvent.ToLongDateString() == DateTime.Today.ToLongDateString())
        //                {
        //                    //The date is today - only need to pick the time
        //                    ItcEnrolTimePicker(iDateBoxIndex, dEvent, _driver);
        //                }
        //                else
        //                {
        //                    //The date is NOT today - need to pick both date and time
        //                    ItcEnrolDatePicker(iDateBoxIndex, dEvent, _driver);
        //                    ItcEnrolTimePicker(iDateBoxIndex, dEvent, _driver);
        //                }
        //                Thread.Sleep(500);
        //                Console.WriteLine("  Entered DateTime(" + sField + "):  " + dEvent.ToString());
        //                break;
        //            case "DATE":
        //                dEvent = GenericFunctions.GetDate(sValue);
        //                iDateBoxIndex = Convert.ToInt32(sID);
        //                ItcEnrolDatePicker(iDateBoxIndex, dEvent, _driver);
        //                Thread.Sleep(500);
        //                Console.WriteLine("  Entered DateTime(" + sField + "):  " + dEvent.ToString());
        //                break;
        //            default:
        //                Console.WriteLine("  [ERROR] EnrolSubject - Unknown Filed Type(" + sField + "):  " + sType);
        //                break;
        //            case "DROPDOWN":
        //                //Enter DropDown data
        //                GenericFunctions.SelectDropDownByText(sValue, sID, _driver);
        //                break;
        //        }
        //    }

        //    //Click 'Enrol Subject'
        //    if (GenericFunctions.ClickSomething("ID", "submitBtn", 20, _driver) == false)
        //    {
        //        //Failed to enter username - abort
        //        Console.WriteLine("  [FAIL]:  Unable to click Submit button");
        //        Console.WriteLine("EnrolSubject:  Finished Function");
        //        return bResult;
        //    }
        //    //Console.WriteLine("  Clicked Enrol Subject button");
        //    Thread.Sleep(1000);
        //    //GenericFunctions.TakePicture(_driver);
        //    try
        //    {
        //        IAlert Alert = _driver.SwitchTo().Alert();
        //        Console.WriteLine("  [Alert]:  Enrollment alert = " + Alert.Text.ToString());
        //        Alert.Accept();
        //    }
        //    catch { }
        //    //Check to see if the enrolment was successful - if the subject ID field is still enabled then it's still in the enrollment form and therefore it's failed
        //    if (GenericFunctions.CheckElementExists("XPATH", "//span[contains(.,'has been successfully enrolled in the study')]", 20, _driver) == true)
        //    {
        //        //Enrollment successful - return success
        //        bResult = true;
        //        Console.WriteLine("  Success - Subject Enrolled");
        //        return bResult;
        //    }
        //    else
        //    {
        //        //Enrolment failed - check for errors
        //        Console.WriteLine("  [ERROR] Enrolment unsuccessful");
        //        bResult = false;
        //        int iErrors = _driver.FindElements(By.XPath("//span[@class='field-validation-error' and @data-valmsg-replace='true']")).Count();
        //        if (iErrors > 0)
        //        {
        //            //Errors found - report errors
        //            string sErrors = iErrors.ToString();
        //            Console.WriteLine("    " + sErrors + " error(s) found:");
        //            for (int i = 0; i <= (iErrors - 1); i++) // Go through all the error results and report each error
        //            {
        //                sErrors = "      " + _driver.FindElements(By.XPath("//span[@class='field-validation-error' and @data-valmsg-replace='true']")).ElementAt(i).GetAttribute("data-valmsg-for");
        //                sErrors = sErrors + ":  " + _driver.FindElements(By.XPath("//span[@class='field-validation-error' and @data-valmsg-replace='true']")).ElementAt(i).GetAttribute("innerHTML").ToString();
        //                Console.WriteLine(sErrors);
        //            }
        //            Console.WriteLine("  bResult:  " + bResult.ToString());
        //            Console.WriteLine("EnrolSubject:  Finished Function");
        //            return bResult;
        //        }
        //        else
        //        {
        //            //No errors found - report mystery!
        //            Console.WriteLine("  No errors found???  iErrors = " + iErrors.ToString());
        //            Console.WriteLine("  bResult:  " + bResult.ToString());
        //            Console.WriteLine("EnrolSubject:  Finished Function");
        //            return bResult;
        //        }
        //    }
        //    //Function completed
        //}



        //public static void ItcEnrolDatePicker(int iIndex, DateTime dDate, IWebDriver _driver)
        //{
        //    //Record parameters passed
        //    Console.WriteLine("  ItcEnrolDatePicker:  Start Function\r    iIndex: " + iIndex.ToString() + "\r    dDate: " + dDate.ToString());
        //    //Click the date picker icon to open the date picker
        //    var vWait = new WebDriverWait(_driver, TimeSpan.FromSeconds(20));
        //    vWait.Until(ExpectedConditions.ElementExists(By.XPath("//*[@class='editor-field short'][" + iIndex.ToString() + "]/div/div/span/span[@class='t-icon t-icon-calendar']")));
        //    Thread.Sleep(500);
        //    _driver.FindElement(By.XPath("//*[@class='editor-field short'][" + iIndex.ToString() + "]/div/div/span/span[@class='t-icon t-icon-calendar']")).Click();
        //    Console.WriteLine("    Clicked DatePicker Icon - DatePicker opened");
        //    //Check whether the year is different from this year
        //    if (dDate.ToString("yyyy") != DateTime.Today.ToString("yyyy"))
        //    {
        //        //The year is different, so change the year and select the relevant month
        //        //Click the month\year in the middle top of the calendar
        //        Console.WriteLine("    Year is different from current year - Change the year");
        //        vWait = new WebDriverWait(_driver, TimeSpan.FromSeconds(20));
        //        vWait.Until(ExpectedConditions.ElementExists(By.XPath("//*[@class='t-widget t-calendar t-popup t-datepicker-calendar' and contains(@style, 'block')]/div/a[2]")));
        //        Thread.Sleep(500);
        //        _driver.FindElement(By.XPath("//*[@class='t-widget t-calendar t-popup t-datepicker-calendar' and contains(@style, 'block')]/div/a[2]")).Click();
        //        Console.WriteLine("      Clicked top centre - display months");
        //        //Click it again to get to the years
        //        vWait = new WebDriverWait(_driver, TimeSpan.FromSeconds(20));
        //        vWait.Until(ExpectedConditions.ElementExists(By.XPath("//*[@class='t-widget t-calendar t-popup t-datepicker-calendar' and contains(@style, 'block')]/div/a[2]")));
        //        Thread.Sleep(500);
        //        _driver.FindElement(By.XPath("//*[@class='t-widget t-calendar t-popup t-datepicker-calendar' and contains(@style, 'block')]/div/a[2]")).Click();
        //        Console.WriteLine("      Clicked top centre - display years");
        //        //Click the relevant year
        //        vWait = new WebDriverWait(_driver, TimeSpan.FromSeconds(20));
        //        vWait.Until(ExpectedConditions.ElementExists(By.LinkText(dDate.ToString("yyyy"))));
        //        Thread.Sleep(500);
        //        _driver.FindElement(By.LinkText(dDate.ToString("yyyy"))).Click();
        //        Console.WriteLine("      Clicked Year:  " + dDate.ToString("yyyy"));
        //        //Click the relevant month
        //        vWait = new WebDriverWait(_driver, TimeSpan.FromSeconds(20));
        //        vWait.Until(ExpectedConditions.ElementExists(By.LinkText(dDate.ToString("MMM"))));
        //        Thread.Sleep(500);
        //        _driver.FindElement(By.LinkText(dDate.ToString("MMM"))).Click();
        //        Console.WriteLine("      Clicked Month:  " + dDate.ToString("MMM"));
        //    }
        //    else
        //    {
        //        //Check whether the month is different from this month
        //        if (dDate.ToString("MM") != DateTime.Today.ToString("MM"))
        //        {
        //            //The month is different, so change the month
        //            Console.WriteLine("    Month is different from current month - Change the month");
        //            //Click the month\year in the middle top of the calendar
        //            vWait = new WebDriverWait(_driver, TimeSpan.FromSeconds(20));
        //            vWait.Until(ExpectedConditions.ElementExists(By.XPath("//*[@class='t-widget t-calendar t-popup t-datepicker-calendar' and contains(@style, 'block')]/div/a[2]")));
        //            Thread.Sleep(500);
        //            _driver.FindElement(By.XPath("//*[@class='t-widget t-calendar t-popup t-datepicker-calendar' and contains(@style, 'block')]/div/a[2]")).Click();
        //            Console.WriteLine("      Clicked top centre - display months");
        //            //Click the relevant month
        //            vWait = new WebDriverWait(_driver, TimeSpan.FromSeconds(20));
        //            vWait.Until(ExpectedConditions.ElementExists(By.LinkText(dDate.ToString("MMM"))));
        //            Thread.Sleep(500);
        //            _driver.FindElement(By.LinkText(dDate.ToString("MMM"))).Click();
        //            Console.WriteLine("      Clicked Month:  " + dDate.ToString("MMM"));
        //        }
        //    }
        //    //Work out the day
        //    string sDay = dDate.ToString("dd");
        //    sDay = sDay.TrimStart('0');
        //    //Click on required day
        //    vWait = new WebDriverWait(_driver, TimeSpan.FromSeconds(20));
        //    vWait.Until(ExpectedConditions.ElementExists(By.LinkText(sDay)));
        //    Thread.Sleep(500);
        //    _driver.FindElement(By.LinkText(sDay)).Click();
        //    Console.WriteLine("    Clicked Day:  " + sDay);
        //    Console.WriteLine("  ItcEnrolDatePicker:  Finished Function");
        //}


        //public static void ItcEnrolTimePicker(int iIndex, DateTime dDateTime, IWebDriver _driver)
        //{
        //    Console.WriteLine("  ItcEnrolTimePicker:  Start Function\r    iIndex: " + iIndex.ToString() + "\r    dDate: " + dDateTime.ToString());
        //    //Work out the required time by rounding the time up to the nearest 5 minute increment
        //    DateTime endTime = Convert.ToDateTime(Convert.ToDateTime(GenericFunctions.RoundUp(DateTime.Parse(dDateTime.ToShortTimeString()), TimeSpan.FromMinutes(5))).ToShortTimeString());
        //    Console.WriteLine("    Rounded endTime:  " + endTime.ToString());
        //    //Work out the dropdown index number for the required time by calculating the difference in minutes from 00:00:00 then dividing it by 5 and adding 1
        //    //it's divided by 5 as thats the increments of this dropdown and the time was rounded to the nearest 5 minutes
        //    //you add 1 as the index starts at 1 and 00:00:00 calculates to a 0 index number
        //    DateTime startTime = Convert.ToDateTime("12:00 AM");
        //    TimeSpan span = endTime.Subtract(startTime);
        //    int myIndex = (Convert.ToInt32(span.TotalMinutes.ToString()) / 5) + 1;
        //    Console.WriteLine("    Dropdown Index:  " + myIndex.ToString());
        //    //click on the time picker
        //    var vWait = new WebDriverWait(_driver, TimeSpan.FromSeconds(20));
        //    vWait.Until(ExpectedConditions.ElementExists(By.XPath("//*[@class='editor-field short'][" + iIndex.ToString() + "]/div/div/span/span[@class='t-icon t-icon-clock']")));
        //    Thread.Sleep(500);
        //    _driver.FindElement(By.XPath("//*[@class='editor-field short'][" + iIndex.ToString() + "]/div/div/span/span[@class='t-icon t-icon-clock']")).Click();
        //    Console.WriteLine("    Clicked TimePicker Icon");
        //    //click on the time
        //    Thread.Sleep(500);
        //    vWait = new WebDriverWait(_driver, TimeSpan.FromSeconds(20));
        //    vWait.Until(ExpectedConditions.ElementExists(By.XPath("//*[@class='t-animation-container' and contains(@style, 'block')]/div/ul/li[" + myIndex.ToString() + "]")));
        //    Console.WriteLine("     - TimePicker dropdown opened");
        //    //Scroll so that all of the dropdown is in view.  this is to stop it crashing when dropdown goes off the bottom of the screen.
        //    System.Drawing.Point point = ((OpenQA.Selenium.Remote.RemoteWebElement)_driver.FindElement(By.XPath("//*[@class='t-animation-container' and contains(@style, 'block')]/div/ul/li[" + myIndex.ToString() + "]"))).LocationOnScreenOnceScrolledIntoView;
        //    _driver.FindElement(By.XPath("//*[@class='t-animation-container' and contains(@style, 'block')]/div/ul/li[" + myIndex.ToString() + "]")).Click();
        //    Console.WriteLine("    Clicked TimePicker dropdown Index: " + myIndex.ToString() + " Clicked");
        //    Console.WriteLine("  ItcEnrolTimePicker:  Finished Function");
        //}



        //public static string GetNextSubjectID(int iLength, char cPad)
        //{
        //    string sSQL = "select max(s.subject_id)+1 from React30.dbo.tbl_subject s";
        //    string sSubjectID = GenericFunctions.GetQueryResult(sSQL);
        //    int iLen = sSubjectID.Length;
        //    if (iLen < iLength)
        //    {
        //        sSubjectID = sSubjectID.PadLeft(iLength, cPad);
        //    }
        //    Console.WriteLine("GetNextSubjectID Function:  sSubjectID=" + sSubjectID + ", iLength=" + iLength.ToString() + ", cPad=" + cPad.ToString());
        //    return sSubjectID;
        //}



        //public static void WaitForLoadingToComplete(IWebDriver _driver)
        //{
        //    Console.WriteLine("WaitForLoadingToComplete Starting:  " + DateTime.Now.ToLongTimeString());
        //    bool bLoading = false;
        //    Thread.Sleep(1000);
        //    int i = 1;
        //    do
        //    {
        //        bLoading = _driver.FindElements(By.XPath("//*[@id='loading'and contains(@style, 'block')]")).Count() != 0;
        //        //Console.WriteLine("     count = " + _driver.FindElements(By.XPath("//*[@id='loading'and contains(@style, 'block')]")).Count().ToString());
        //        //Console.WriteLine("     i = " + i.ToString());
        //        i++;
        //    }
        //    while(bLoading == true);
        //    Console.WriteLine("WaitForLoadingToComplete Finshed:  " + DateTime.Now.ToLongTimeString());
        //}








    }
}
