using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.IE;
using NUnit.Framework;
using OpenQA.Selenium.Support.Events;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using TestTeamAutomationFrameNew.Functions;

namespace TestTeamAutomationFrameNew
{
    //public  class ReactCell
    //{

    //    public static bool GetSubjectsShortenedURL(string sStudy, string sSite, string sSubjectID, out string sShortURL)
    //    {
    //        //Find the subjects shortened URL from the database
    //        Console.WriteLine("GetSubjectsShortenedURL Starting Function:  sStudy=" + sStudy + ", sSite = " + sSite + ", sSubjectID=" + sSubjectID);
    //        string sSQL = "SELECT ul.subject_unique_lookup FROM React30.dbo.tbl_subject su left outer join React30.dbo.tbl_site si on su.site_id = si.site_id left outer join React30.dbo.tbl_study st on si.study_id = st.study_id left outer join React30.dbo.tbl_subject_unique_lookup ul on su.subject_id = ul.subject_id WHERE st.name = '" + sStudy + "' and si.name = '" + sSite + "' and su.user_subject_id = '" + sSubjectID + "'";
    //        sShortURL = _dbActions.GetQueryResult(sSQL);
    //        if (sShortURL == "")
    //        {
    //            //Failed to get shortened URL from the database
    //            Console.WriteLine("GetSubjectsShortenedURL Finished Function[FAIL]:  Failed to get shortened URL from the database");
    //            return false;
    //        }
    //        sShortURL = GenericFunctions.goAndGet("SHORTENEDURL") + sShortURL;
    //        Console.WriteLine("GetSubjectsShortenedURL Finished Function[SUCCESS]:  sShortURL=" + sShortURL);
    //        return true;
    //    }


    //    private static void PopulateFields(string sFileName, IWebDriver _driver)
    //    {

    //        //Load data from CSV and enter the required dat in the enrollment form
    //        string[,] saValues = GenericFunctions.LoadCsv(sFileName);
    //        int iRows = saValues.GetUpperBound(0);
    //        //Console.WriteLine("iRows:  " + iRows.ToString());
    //        int iCols = saValues.GetUpperBound(1);
    //        //Console.WriteLine("iCols:  " + iCols.ToString());
    //        for (int r = 1; r <= iRows; r++)
    //        {
    //            string sField = saValues[r, 0].ToString();
    //            Console.WriteLine("  sField:  " + sField);
    //            string sType = saValues[r, 1].ToString();
    //            Console.WriteLine("  sType:  " + sType);
    //            string sBy = saValues[r, 2].ToString();
    //            Console.WriteLine("  sBy:  " + sBy);
    //            string sID = saValues[r, 3].ToString();
    //            Console.WriteLine("  sID:  " + sID);
    //            string sValue = saValues[r, 4].ToString();
    //            Console.WriteLine("  sValue:  " + sValue);
    //            DateTime dEvent;
    //            int iDateBoxIndex;
    //            switch (sType.ToUpper())
    //            {
    //                case "TEXTBOX":
    //                    //Enter TextBox data
    //                    if (GenericFunctions.EnterTextBox("ID", sID, sValue, 20, _driver) == false)
    //                    {
    //                        //Failed to enter something - raise warning and carry on
    //                        Console.WriteLine("  [WARNING]:  Unable to enter data for field " + sField + "):  " + sValue);
    //                    }
    //                    Console.WriteLine("  Entered TextBox Data (" + sField + "):  " + sValue);
    //                    break;
    //                case "DATETIME":
    //                    dEvent = GenericFunctions.GetDate(sValue);
    //                    iDateBoxIndex = Convert.ToInt32(sID);
    //                    //Console.WriteLine("  dEvent.ToLongDateString() = " + dEvent.ToLongDateString());
    //                    //Console.WriteLine("  DateTime.Today.ToLongDateString() = " + DateTime.Today.ToLongDateString());
    //                    if (dEvent.ToLongDateString() == DateTime.Today.ToLongDateString())
    //                    {
    //                        //The date is today - only need to pick the time
    //                        ItcEnrol.ItcEnrolTimePicker(iDateBoxIndex, dEvent, _driver);
    //                    }
    //                    else
    //                    {
    //                        //The date is NOT today - need to pick both date and time
    //                        ItcEnrol.ItcEnrolDatePicker(iDateBoxIndex, dEvent, _driver);
    //                        ItcEnrol.ItcEnrolTimePicker(iDateBoxIndex, dEvent, _driver);
    //                    }
    //                    Thread.Sleep(500);
    //                    Console.WriteLine("  Entered DateTime(" + sField + "):  " + dEvent.ToString());
    //                    break;
    //                case "DATE":
    //                    dEvent = GenericFunctions.GetDate(sValue);
    //                    iDateBoxIndex = Convert.ToInt32(sID);
    //                    ItcEnrol.ItcEnrolDatePicker(iDateBoxIndex, dEvent, _driver);
    //                    Thread.Sleep(500);
    //                    Console.WriteLine("  Entered DateTime(" + sField + "):  " + dEvent.ToString());
    //                    break;
    //                default:
    //                    Console.WriteLine("  Unknown Filed Type(" + sField + "):  " + sType);
    //                    break;
    //            }
    //        }
    //    }



    //    public static bool CheckEditChecks(string sEditChecks, string sEditIcons, IWebDriver _driver)
    //    {
    //        Console.WriteLine("CheckEditChecks Starting Function:  sEditChecks=" + sEditChecks + ", sEditIcons=" + sEditIcons);
    //        bool bResult = false;
    //        List<string> slEditChecks = new List<string>();
    //        List<string> slEditIcons = new List<string>();
    //        string sFoundEditChecks;
    //        string sFounEditIcons;
    //        CaptureEditChecks(ref slEditChecks, ref slEditIcons, _driver);
    //        if (slEditChecks == null) { sFoundEditChecks = ""; } else { sFoundEditChecks = string.Join(";", slEditChecks); }
    //        if (sEditChecks != sFoundEditChecks) { Console.WriteLine("     Missmatch on Messages:  sFoundEditChecks=" + sFoundEditChecks); ; }  //If edit check MESSAGES are different from expected - report what was found
    //        if (slEditIcons == null) { sFounEditIcons = ""; } else { sFounEditIcons = string.Join(";", slEditIcons); }
    //        if (sEditIcons != sFounEditIcons) { Console.WriteLine("     Missmatch on ICONS:  sFounEditIcons=" + sFounEditIcons); ; }  //If edit check ICONS are different from expected - report what was found
    //        if (sEditChecks == sFoundEditChecks && sEditIcons == sFounEditIcons) { bResult = true; }  //If both edit checks are correct - returun true
    //        Console.WriteLine("CheckEditChecks Finished Function:  bResult=" + bResult.ToString());
    //        return bResult;
    //    }



    //    public static bool CaptureEditChecks(ref List<string> slEditChecks, ref List<string> slEditIcons, IWebDriver _driver)
    //    {
    //        bool bResult = false;
    //        Console.WriteLine("CaptureEditChecks Starting Function");
    //        //capture a list of edit check messages
    //        slEditChecks = PageCheck.ListElementCollectionsTextValuesToStringList("//*[@class='errorValidationSummary']/ul/li", _driver);
    //        //capture a list of edit check icons id names
    //        slEditIcons = PageCheck.ListElementCollectionsAttributeToStringList("//*[@class='errorValidationItem' and not(contains(@style, 'hidden'))]", "id", _driver);
    //        //check to see if any results were found and report true if there were
    //        if (slEditChecks != null || slEditIcons != null) { bResult = true; }
    //        Console.WriteLine("CaptureEditChecks Finished Function:  bResult" + bResult);
    //        return bResult;
    //    }



    //    public static bool CheckEditWarnings(string sEditWarnings, string sEditWarningIcons, IWebDriver _driver)
    //    {
    //        Console.WriteLine("CheckEditWarnings Starting Function:  sEditWarnings=" + sEditWarnings + ", sEditWarningIcons=" + sEditWarningIcons);
    //        bool bResult = false;
    //        List<string> slEditWarnings = new List<string>();
    //        List<string> slEditWarningIcons = new List<string>();
    //        string sFoundEditWarnings;
    //        string sFounEditWarningIcons;
    //        CaptureEditWarnings(ref slEditWarnings, ref slEditWarningIcons, _driver);
    //        if (slEditWarnings == null) { sFoundEditWarnings = ""; } else { sFoundEditWarnings = string.Join(";", slEditWarnings); }
    //        if (sEditWarnings != sFoundEditWarnings) { Console.WriteLine("     Missmatch on Messages:  sFoundEditWarnings=" + sFoundEditWarnings); }  //If edit check MESSAGES are different from expected - report what was found
    //        if (slEditWarningIcons == null) { sFounEditWarningIcons = ""; } else { sFounEditWarningIcons = string.Join(";", slEditWarningIcons); }
    //        if (sEditWarningIcons != sFounEditWarningIcons) { Console.WriteLine("     Missmatch on ICONS:  sFounEditWarningIcons=" + sFounEditWarningIcons); }  //If edit check ICONS are different from expected - report what was found
    //        if (sEditWarnings == sFoundEditWarnings && sEditWarningIcons == sFounEditWarningIcons) { bResult = true; }  //If both edit checks are correct - returun true
    //        Console.WriteLine("CheckEditWarnings Finished Function:  bResult=" + bResult.ToString());
    //        return bResult;
    //    }



    //    public static bool CaptureEditWarnings(ref List<string> slEditWarnings, ref List<string> slEditWarningIcons, IWebDriver _driver)
    //    {
    //        bool bResult = false;
    //        Console.WriteLine("CaptureEditWarnings Starting Function");
    //        //capture a list of edit check messages
    //        slEditWarnings = PageCheck.ListElementCollectionsTextValuesToStringList("//*[@class='warnValidationSummary']/ul/li", _driver);
    //        //capture a list of edit check icons id names
    //        slEditWarningIcons = PageCheck.ListElementCollectionsAttributeToStringList("//*[@class='warnValidationItem' and not(contains(@style, 'hidden'))]", "id", _driver);
    //        //check to see if any results were found and report true if there were
    //        if (slEditWarnings != null || slEditWarningIcons != null) { bResult = true; }
    //        Console.WriteLine("CaptureEditWarnings Finished Function:  bResult" + bResult);
    //        return bResult;
    //    }




    //    public static Boolean ReactRadioButton(string sID, int iOption, int iWait, IWebDriver _driver)
    //    {
    //        Console.WriteLine("ReactRadioButton Starting Function:  sID:  " + sID + ", iOption:  " + iOption.ToString() + ", iWait:  " + iWait.ToString());
    //        Boolean bResult = false;
    //        //Check to see if the required element exists
    //        if (GenericFunctions.CheckElementExists("ID", sID, iWait, _driver) == false)
    //        {
    //            //Unable to find dropdown - abort
    //            Console.WriteLine("  [FAIL]:  Element NOT found.");
    //            Console.WriteLine("ReactRadioButton Finished Function:  bResult=" + bResult.ToString());
    //            return bResult;
    //        }
    //        //Element exists - Check it has the required option
    //        string sXpath = "//*[@id='" + sID + "']/tbody/tr/td/span/label";
    //        int iOptions = _driver.FindElements(By.XPath(sXpath)).Count();
    //        Console.WriteLine("  iOptions=" + iOptions.ToString());
    //        if (iOptions == 0)
    //        {
    //            //No dropdown options found
    //            Console.WriteLine("  [FAIL]:  No radio options found for element:  sXpath=" + sXpath);
    //            Console.WriteLine("ReactRadioButton Finished Function:  bResult=" + bResult.ToString());
    //            return bResult;
    //        }
    //        if (iOptions < iOption + 1)
    //        {
    //            //No dropdown options found
    //            Console.WriteLine("  [FAIL]:  Requested radio option (iOption=" + iOption.ToString() + ") not found for element(iOptions=" + iOptions.ToString() + "):  sXpath=" + sXpath);
    //            Console.WriteLine("ReactRadioButton Finished Function:  bResult=" + bResult.ToString());
    //            return bResult;
    //        }
    //        //Option exists - select it
    //        _driver.FindElements(By.XPath(sXpath)).ElementAt(iOption).Click();
    //        bResult = true;
    //        Console.WriteLine("ReactRadioButton Finished Function:  bResult=" + bResult.ToString());
    //        return bResult;
    //    }

    //}
}
