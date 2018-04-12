using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using NUnit.Framework;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Xml;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Internal;
using OpenQA.Selenium.Support.UI;
using System.Text.RegularExpressions;
using TestTeamAutomationFrameNew.ITC_SMK;
//using TestTeamAutomationFrameNew.AndroidAutomation;
using TestTeamAutomationFrameNew.Functions;


namespace TestTeamAutomationFrameNew.Functions
{
    public static class GenericFunctions
    {
        #region React

        public static void LogInToReact(string username, string pswrd, IWebDriver driver)
        {
            //driver.Navigate().GoToUrl("http://test.reactintouch.com/login.aspx?th=DEFAULT");
            driver.Navigate().GoToUrl(goAndGet("ReactLoginURL"));
            //selenium.Open(goAndGet("ReactLoginURL"));

            //CheckWindowTitle("InTouch-Login");-- Remeber to uncomment
            driver.FindElement(By.Id("ctl00_middlecontent_Login1_UserName")).SendKeys(username);
            driver.FindElement(By.Id("ctl00_middlecontent_Login1_Password")).SendKeys(pswrd);
            //driver.FindElement(By.Id("ctl00_middlecontent_Login1_Password")).SendKeys(Keys.Enter);
            ClickOnButton("ctl00_middlecontent_Login1_LoginButton", driver);
            CheckTextIsOnPage("Welcome:", driver);
            ReportTheTestPassed("Logged in as user \"" + username + "\"");
        }

        public static void LogOutOfReact(IWebDriver driver)
        {
            //ClickOnText("LOGOUT", driver);
        }

        public static void UploadSitesAndMessages(string sites, string messages, IWebDriver driver)
        {
            driver.FindElement(By.Id("ctl00_middlecontent_UploadSites1_FileUploadSite")).SendKeys(sites);
            driver.FindElement(By.Id("ctl00_middlecontent_UploadSites1_Btn_UploadSite")).Click();
            driver.FindElement(By.Id("ctl00_middlecontent_UploadSites1_FileUploadSiteMsg")).SendKeys(messages);
            driver.FindElement(By.Id("ctl00_middlecontent_UploadSites1_Btn_UploadSiteMessage")).Click();
        }

        public static void CheckSuperUserLinks(IWebDriver driver)
        {
            CheckTextIsOnPage("STUDIES", driver);
            CheckLinkIsOnPage("SITES", driver);
            CheckLinkIsOnPage("USERS", driver);
            CheckLinkIsOnPage("SUBJECT", driver);
            CheckLinkIsOnPage("LOGOUT", driver);
            MouseOver("STUDIES", driver);
            //Actions action = new Actions(driver);
            //action.MoveToElement(driver.FindElement(By.LinkText("STUDIES"))).Perform();
            CheckLinkIsOnPage("VIEW STUDY", driver);
            CheckLinkIsOnPage("ADD STUDY", driver);
            MouseOver("SITES", driver);
            CheckLinkIsOnPage("VIEW SITE", driver);
            CheckLinkIsOnPage("ADD/EDIT SITE", driver);
            MouseOver("USERS", driver);
            CheckLinkIsOnPage("VIEW USER", driver);
            CheckLinkIsOnPage("ADD USER", driver);
            MouseOver("SUBJECT", driver);
            CheckLinkIsOnPage("ENROLL SUBJECT", driver);
            CheckLinkIsOnPage("WITHDRAW SUBJECT", driver);
            CheckLinkIsOnPage("EDIT SUBJECT", driver);
            CheckLinkIsOnPage("VIEW SUBJECT", driver);
            CheckLinkIsOnPage("HOME", driver);
            MouseOver("LOGOUT", driver);
            //ReportTheTestPassed("All links in React are enabled for the user.");
        }

        //public static void GoToEditUserPage(IWebDriver driver)
        //{
        //    MouseOver("USERS", driver);
        //    ClickOnText("VIEW USER", driver);
        //    ClickOnText("S", driver);
        //    IWebElement elm =
        //        driver.FindElement(
        //            By.XPath(
        //                "//table[@id='ctl00_middlecontent_GrdVw_ViewUsers']/tbody/tr[td[position()=1 and contains(.,'" +
        //                goAndGet("NEWUSER") + "')]]/td[5]/input[@type='submit']"));
        //    ClickElement(elm, driver);
        //}

        public static void RemoveAllUserRoles(IWebDriver driver)
        {
            var selectitem = new SelectElement(driver.FindElement(By.Id("ctl00_middlecontent_Lst_Roles")));
            selectitem.DeselectAll();
        }

        public static void SetStudyAndSiteForUser(string study, IWebDriver driver)
        {
            IWebElement cb =
                driver.FindElement(
                    By.XPath("//table[@id='ctl00_middlecontent_Grd_Study']/tbody/tr[td[position()=1 and contains(.,'" +
                             goAndGet(study) + "')]]/td[3]/input[@type='checkbox']"));
            ClickElement(cb, driver);
            //selenium.Check("//table[@id='ctl00_middlecontent_Grd_Study']/tbody/tr[td[position()=1 and contains(.,'" + goAndGet(study) + "')]]/td[3]/input[@type='checkbox']");
            ClickOnButton("ctl00_middlecontent_Btn_StudySave", driver);
            IWebElement cb1 =
                driver.FindElement(
                    By.XPath("//table[@id='ctl00_middlecontent_Grd_Site']/tbody/tr[td[position()=1 and contains(.,'" +
                             goAndGet(study) + "')]]/td[4]/input[@type='checkbox']"));
            ClickElement(cb1, driver);
            //selenium.Check("//table[@id='ctl00_middlecontent_Grd_Site']/tbody/tr[td[position()=1 and contains(.,'" + goAndGet(study) + "')]]/td[4]/input[@type='checkbox']");
            ClickOnButton("ctl00_middlecontent_Btn_SiteSave", driver);
        }

        #endregion


        #region WebFormChecks

        public static void CheckCheckboxChecked(string Selected, string sId, IWebDriver driver)
        {
            string Checked = driver.FindElement(By.Id(sId)).GetAttribute("checked");
            if (Selected == "Yes")
            {
                if (Checked == "true")
                {
                    ReportTheTestPassed("The correct checkbox is checked");
                }
                else
                {
                    ReportTheTestFailed("The checkbox is NOT checked");
                    _startUp.stepsFailed++;
                }
            }
            else if (Selected == "No")
            {
                if (Checked == "false")
                {
                    ReportTheTestFailed("The checkbox is checked");
                    _startUp.stepsFailed++;
                }
                else
                {
                    ReportTheTestPassed("The correct checkbox is NOT checked");
                }
            }

        }

        public static void CheckWindowTitle(string titletext, IWebDriver driver)
        {
            if (titletext != driver.Title)
            {
                ReportTheTestFailed("The window title is incorrect: \"" + titletext + "\"");
                _startUp.stepsFailed++;

            }
            else
            {
                ReportTheTestPassed("The window title was correct \"" + titletext + "\"");
            }
        }




        public static void CheckTextIsOnPageNew(string textmsg, IWebDriver driver)
        {
            IWebElement bodyTag = driver.FindElement(By.TagName("body"));
            if (bodyTag.Text.Contains(textmsg))
            {
                
                Console.WriteLine("Report Pass: Text '"+ textmsg + "'" + " is on page");
                
            }
            else
            {
                // Wait for 1 seconds and try and find the text again, otherwise fail it.
                Wait(1);
                if (bodyTag.Text.Contains(textmsg))
                {
                    Console.WriteLine("Report Pass: Text '" + textmsg + "'" + " is on page");
                }
                else
                    Console.WriteLine("Report Fail: Text '" + textmsg + "'" + " is NOT on page");
            }
        }

        public static void CheckTextIsOnPage(string textmsg, IWebDriver driver)
        {
            IWebElement bodyTag = driver.FindElement(By.TagName("body"));
            if (bodyTag.Text.Contains(textmsg))
            {
                ReportTheTestPassed("The text is present: \"" + textmsg + "\"");
                _startUp.TestPasses++;
                _startUp.Passes++;
            }
            else
            {
                // Wait for 1 seconds and try and find the text again, otherwise fail it.
                Wait(1);
                if (bodyTag.Text.Contains(textmsg))
                {
                    Warning("The text is present: \"" + textmsg + "\"");
                    _startUp.Warnings++;
                    _startUp.TestPasses++;
                    _startUp.Passes++;
                }
                else
                    ReportTheTestFailed("The text is not present on the page: \"" + textmsg + "\"");
                _startUp.TestFails++;
                _startUp.Fails++;
            }
        }

        // public static void CheckTextBoxValue(string textmsg, string id, IWebDriver driver)
        //{
        //    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
        //    var element = wait.Until(ExpectedConditions.ElementIsVisible(By.Id(id)));
        //    IWebElement TextboxValue = driver.FindElement(By.Id(id));            
        //    if (TextboxValue.GetAttribute("value").Contains(textmsg))
        //    {
        //        ReportTheTestPassed("The text is present: \"" + textmsg + "\"");
        //        _startUp.TestPasses++;
        //        _startUp.Passes++;
        //    }
        //    else
        //    {
        //        // Wait for 0.1 seconds and try and find the text again, otherwise fail it.
        //        Thread.Sleep(100);
        //        if (TextboxValue.Text.Contains(textmsg))
        //        {
        //            ReportTheTestPassed("The text is present: \"" + textmsg + "\"");
        //            _startUp.Warnings++;
        //            _startUp.TestPasses++;
        //            _startUp.Passes++;
        //        }
        //        else
        //            ReportTheTestFailed("The text is not present on the page: \"" + textmsg + "\"");
        //        _startUp.TestFails++;
        //        _startUp.Fails++;
        //    }
        //}

        //public static void CheckDropDownValue(string textmsg, string id, IWebDriver driver)
        //{
        //    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
        //    var element = wait.Until(ExpectedConditions.ElementIsVisible(By.Id(id)));
        //    IWebElement ComboBox = driver.FindElement(By.Id(id));
        //    SelectElement selectedValue = new SelectElement(ComboBox);
        //    //string ComboValue = selectedValue.SelectedOption.Text;
        //    if (selectedValue.SelectedOption.Text.Contains(textmsg))
        //    {
        //        ReportTheTestPassed("The text is present: \"" + textmsg + "\"");
        //        _startUp.TestPasses++;
        //        _startUp.Passes++;
        //    }
        //    else
        //    {
        //        // Wait for 0.1 seconds and try and find the text again, otherwise fail it.
        //        Thread.Sleep(100);
        //        if (selectedValue.SelectedOption.Text.Contains(textmsg))
        //        {
        //            ReportTheTestPassed("The text is present: \"" + textmsg + "\"");
        //            _startUp.Warnings++;
        //            _startUp.TestPasses++;
        //            _startUp.Passes++;
        //        }
        //        else
        //            ReportTheTestFailed("The text is not present on the page: \"" + textmsg + "\"");
        //        _startUp.TestFails++;
        //        _startUp.Fails++;
        //    }
        //}

        //public static void CheckAppTextOnScreen(string Text, AppiumDriver Appdriver)
        //{


        //    if (Standroid.IsAndroidElementPresent(Text, Appdriver) == true)
        //    {
        //        ReportTheTestPassed("The text is present: \"" + Text + "\"");
        //        _startUp.TestPasses++;
        //        _startUp.Passes++;
        //    }
        //    else
        //    {
        //        if (GenericFunctions.IsElementPresent(Text, Appdriver))
        //        {
        //            ReportTheTestPassed("The text is present: \"" + Text + "\"");
        //            _startUp.Warnings++;
        //            _startUp.TestPasses++;
        //            _startUp.Passes++;
        //        }
        //        else
        //        {
        //            GenericFunctions.ReportTheTestFailed("The text is not present in the app: \"" + Text + "\"");
        //            _startUp.TestFails++;
        //            _startUp.Fails++;
        //        }

        //    }


        //}
        public static string ReplaceAllnonAlphabString(string value)
        {
            var regex = new Regex(@"(\s*)\w*:"); // strip all of the non alphabetic characters 
            return value = regex.Replace(value, string.Empty);
        }

        public static string ReplaceAllTextString(string value)
        {
            var regex = new Regex("[^0-9 :]"); // strip all of the alphabetic characters except :'
            return value = regex.Replace(value, string.Empty);
        }



        public static void CheckIfAnythingFailed(string TestCase, int stepsFailed)
        {
            if (stepsFailed > 0)
            {
                WriteTestNameToReport(TestCase +
                                      "test case <span style=\"color: red;\"><strong>FAILED</strong></span> - " +
                                      stepsFailed + " step(s) FAILED.");
            }
            else
            {
                ReportTheTestPassed(TestCase + "test case <span style=\"color: green;\"><strong>PASSED</strong></span>");
                _startUp.TestsPassed++;
            }
        }

        private static void CheckAlert(string alertmessage, IWebDriver driver)
        {
            IAlert alert = driver.SwitchTo().Alert();
            string alertText = alert.Text;
            if (alertText == alertmessage)
            {
                ReportTheTestPassed("The alert is present: \"" + alertmessage + "\"");
            }
            else
            {
                ReportTheTestFailed("The alert is not present on the page: \"" + alertmessage + "\"");
                _startUp.stepsFailed++;
            }
            // Accept the alert
            alert.Accept();

        }

        public static void CheckLinkIsOnPage(string textmsg, IWebDriver driver)
        {
            IWebElement bodyTag = driver.FindElement(By.TagName("body"));
            if (bodyTag.Text.Contains(textmsg))
            {
                ReportTheTestPassed("The link is present: \"" + textmsg + "\"");
            }
            else
            {
                ReportTheTestFailed("The link is not present on the page: \"" + textmsg + "\"");
                _startUp.stepsFailed++;
            }
        }

        public static void CheckTextIsNotOnPage(string textmsg, IWebDriver driver)
        {
            IWebElement bodyTag = driver.FindElement(By.TagName("body"));
            if (bodyTag.Text.Contains(textmsg))
            {
                ReportTheTestFailed("The text is present on the page but shouln't be: \"" + textmsg + "\"");
                _startUp.stepsFailed++;
            }
            else
            {
                ReportTheTestPassed("The text is not present which is correct: \"" + textmsg + "\"");
            }
        }

        public static void CheckTextIsNotOnPageNew(string textmsg, IWebDriver driver)
        {
            IWebElement bodyTag = driver.FindElement(By.TagName("body"));
            if (bodyTag.Text.Contains(textmsg))
            {

                Console.WriteLine("Report Fail: Text '" + textmsg + "'" + " should NOT be on the page");

            }
            else
            {
                // Wait for 1 seconds and try and find the text again, otherwise pass it.
                Wait(1);
                if (bodyTag.Text.Contains(textmsg))
                {
                    Console.WriteLine("Report Fail: Text '" + textmsg + "'" + " should NOT be on the page");
                }
                else
                    Console.WriteLine("Report Pass: Text '" + textmsg + "'" + " is NOT on page");
            }
        }



        public static void CheckItemIsOnPage(string textmsg, string item, IWebDriver driver)
        {
            IWebElement elm = driver.FindElement(By.Id(item));
            if (textmsg == elm.Text)
            {
                ReportTheTestPassed("The item or element is present: \"" + textmsg + "\"");
            }
            else
            {
                ReportTheTestFailed("The item or element is not present on the page: \"" + textmsg + "\"");
                _startUp.stepsFailed++;
            }
        }

        //#merge# added CheckElementExists
        public static Boolean CheckElementExists(string sBy, string sLocator, int iTries, IWebDriver _driver)
        {
            Boolean bFound = false;
            int i = 0;
            //Check to see if the element exists however many times it says in iTries
            for (i = 0; i <= iTries; i++)
            {
                switch (sBy.ToUpper())
                {
                    case "ID":
                        bFound = _driver.FindElements(By.Id(sLocator)).Count() != 0;
                        Console.WriteLine("     _driver.FindElements(By.Id(" + sLocator + ")).Count() = " +
                                          _driver.FindElements(By.Id(sLocator)).Count().ToString());
                        break;
                    case "XPATH":
                        bFound = _driver.FindElements(By.XPath(sLocator)).Count() != 0;
                        Console.WriteLine("     _driver.FindElements(By.XPath(" + sLocator + ")).Count() = " +
                                          _driver.FindElements(By.XPath(sLocator)).Count().ToString());
                        break;
                    default:
                        //Unknown BY type - abort
                        Console.WriteLine(
                            "CheckElementExists Function:  bFound={0} [WARNING]:  Unknown BY type (i={1}, sBy={2}, sLocator={3}, iTries={4})",
                            bFound.ToString(), i.ToString(), sBy, sLocator, iTries.ToString());
                        return bFound;
                }
                //Drop out if the element is found
                if (bFound == true)
                {
                    break;
                }
            }
            //Report result
            Console.WriteLine("CheckElementExists Function:  bFound={0}, i={1}, sBy={2}, sLocator={3}, iTries={4}",
                bFound.ToString(), i.ToString(), sBy, sLocator, iTries.ToString());
            return bFound;
        }

        //#merge# added CheckElementEnabled
        public static bool CheckElementEnabled(string sBy, string sLocator, IWebDriver _driver)
        {
            bool bResult = false;
            switch (sBy.ToUpper())
            {
                case "ID":
                    if (_driver.FindElement(By.Id(sLocator)).Enabled == true)
                    {
                        bResult = true;
                    }
                    break;
                case "XPATH":
                    if (_driver.FindElement(By.XPath(sLocator)).Enabled == true)
                    {
                        bResult = true;
                    }
                    break;
                default:
                    //Unknown BY type - abort
                    Console.WriteLine(
                        "CheckElementEnabled Function:  bResult={0}:  [WARNING] Unknown BY type (sBy={1}, sLocator={2})",
                        bResult, sBy, sLocator);
                    return bResult;
            }
            Console.WriteLine("CheckElementEnabled Function:  bResult={0}, sBy={1}, sLocator={2}", bResult, sBy,
                sLocator);
            return bResult;
        }

        #endregion


        #region WebFormCommands

        public static void Type(string value, string locator, IWebDriver driver)
        {
            driver.FindElement(By.Id(locator)).SendKeys(value);
        }

        public static void Clear(string locator, IWebDriver driver)
        {
            driver.FindElement(By.Id(locator)).Clear();
        }

        public static void ClickOnButton(string button, IWebDriver driver)
        {
            IWebElement elm = driver.FindElement(By.Id(button));
            elm.Click();
            //ClickElement(elm, driver);
            //GenericFunctions.Wait()

            //if (selenium.IsElementPresent(button))
            //{
            //    selenium.Click(button);
            //    selenium.WaitForPageToLoad("30000");
            //}
            //else
            //{
            //    ReportTheTestFailed("The button is not present on the page: \"" + button + "\"");
            //    stepsFailed++;
            //}
        }

        public static bool IsElementIdPresent(string element, IWebDriver driver)
        {
            try
            {
                driver.FindElement(By.Id(element));
                //Thread.Sleep(4000);
                return true; // Success! 
            }
            catch (NoSuchElementException)
            {
                Console.WriteLine("Error:" + element + " not displayed");
                return false; // Do not throw exception 
            }
        }


        public static bool IsElementPresent(string element, IWebDriver driver)
        {
            try
            {
                driver.FindElement(By.Name(element));
                //Thread.Sleep(4000);
                return true; // Success! 
            }
            catch (NoSuchElementException)
            {
                Console.WriteLine("Error:" + element + " not displayed");
                return false; // Do not throw exception 
            }
        }

        public static bool IsElementXpathPresent(string xpath, IWebDriver driver)
        {
            try
            {
                driver.FindElement(By.XPath(xpath));
                //Thread.Sleep(4000);
                return true; // Success! 
            }
            catch (NoSuchElementException)
            {
                Console.WriteLine("Error:" + xpath + " not displayed");
                return false; // Do not throw exception 
            }
        }

        public static void ClickOnText(string nametext, IWebDriver driver)
        {
            driver.FindElement(By.Id(nametext)).Click();
            //driver.FindElement(By.LinkText(nametext)).Click();
        }

        public static void waitforElement(string nametext, IWebDriver driver)
        {
            int counter = 0;


            while (!driver.FindElement(By.LinkText(nametext)).Displayed)
            {

                if (counter == 100)
                {
                    counter++;
                    break;
                }

            }
            
        }

        public static void ClickOnLink(string nametext, IWebDriver driver)
        {
            //var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(40));
            //wait.Until(ExpectedConditions.ElementExists(By.Id(nametext)));
            driver.FindElement(By.XPath("//a[.= '" + nametext + "']")).Click();
           // driver.Manage().Timeouts().ImplicitlyWait(new System.TimeSpan(40));
            //driver.FindElement(By.LinkText(nametext)).Click();
        }
        public static void ClickonBtn(string nametext, IWebDriver driver)
        {
            driver.FindElement(By.XPath("//input[@value= '" + nametext + "']")).Click();
            //driver.Manage().Timeouts().ImplicitlyWait(new System.TimeSpan(40));
        }

        public static void SelectFromList(string listof, string locator, IWebDriver driver)
        {
            SelectElement selectitem = new SelectElement(driver.FindElement(By.Id(locator)));
            selectitem.SelectByText(listof);
        }

        public static void UnSelectFromList(string listof, string locator, IWebDriver driver)
        {
            SelectElement selectitem = new SelectElement(driver.FindElement(By.Id(locator)));
            selectitem.DeselectByText(listof);
        }

        public static void ListBoxSelect(string listof, string locator, IWebDriver driver)
        {
            //Populate array with lsit of values
            string list = listof;
            string[] values = list.Split(',');
            var val = "";
            //Create a list of all the options in the listbox
            IWebElement elem = driver.FindElement(By.Id(locator));
            IList<IWebElement> options = elem.FindElements(By.TagName("option"));
            //Select the listbox
            SelectElement seldata = new SelectElement(driver.FindElement(By.Id(locator)));
            //loop through the list the user has entered
            foreach (string item in values)
            {
                string opt = item.Trim();
                //loop through the listbox options
                foreach (IWebElement option in options)
                {
                    //if the listbox option text matches the text entered by the user populate the val variable
                    if (option.Text == opt)
                    {
                        val = option.GetAttribute("value");
                    }
                }                
                seldata.SelectByValue(val);//Select the option using the val variable
            }            
        }

        public static void SelectBySubText(this SelectElement me, string subText)
        {
            foreach (var option in me.Options.Where(option => option.Text.Contains(subText)))
            {
                option.Click();
                return;
            }
            me.SelectByText(subText);
        }

        public static void MouseOver(string textname, IWebDriver driver)
        {
            Actions action = new Actions(driver);
            action.MoveToElement(driver.FindElement(By.LinkText(textname))).Perform();
        }

        public static void TickCheckBox(string checkbox, IWebDriver driver)
        {
            IWebElement cb = driver.FindElement(By.Id(checkbox));
            if (cb.Selected)
            {
            }
            else
                ClickElement(cb, driver);
        }

        

        public static void UnTickCheckBox(string checkbox, IWebDriver driver)
        {
            IWebElement cb = driver.FindElement(By.Id(checkbox));
            if (cb.Selected)
            {
                ClickElement(cb, driver);
            }
        }

        public static void ClickElement(IWebElement elm, IWebDriver driver)
        {

            string myBrowser = (String)((IJavaScriptExecutor)driver).ExecuteScript("return navigator.appName;");

            //string BROWSER = _genericFunctions.goAndGet("BROWSER");

            if (myBrowser == "Microsoft Internet Explorer")
                elm.SendKeys("\n");
            else if (myBrowser == "Netscape")
                // Do not change this, as the browser name is returned Netcape even though it is FF
                elm.Click();
            else
                elm.Click();
            //((ITakesScreenshot)driver).GetScreenshot().SaveAsFile("withdraw20.png", ImageFormat.Png);
        }

        public static void SelectDropDownByText(string text, string elementID, IWebDriver driver)
        {
            IWebElement ddl = driver.FindElement(By.Id(elementID));
            SelectElement se = new SelectElement(ddl);
            se.SelectByText(text);
            //ddl.SendKeys(text); // Type what you want   
            //ddl.SendKeys(Keys.Tab); //Tab out of drop down
        }

        public static void SelectSpanDropDownByText(string text, string DropDownXpath, string ListDropDown, IWebDriver driver)
        {
            IWebElement ddl = driver.FindElement(By.XPath(DropDownXpath)); //Finds the span which holds the UL
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("arguments[0].click();", ddl); //Opens the dropdown
            Thread.Sleep(2000);
            driver.FindElement(By.XPath(ListDropDown + "[text()='" + text + "']")).Click();            
        }

        public static void SetAttribute(this IWebElement element, string attributeName, string value)
        {
            var wrappedElement = element as IWrapsDriver;
            if (wrappedElement == null)
                throw new ArgumentException("element", "Element must wrap a web driver");
            IWebDriver driver = wrappedElement.WrappedDriver;
            var js = driver as IJavaScriptExecutor;
            if (js == null)
                throw new ArgumentException("element",
                    "Element must wrap a web driver that supports javascript execution");
            js.ExecuteScript("arguments[0].setAttribute(arguments[1], arguments[2])", element, attributeName, value);
        }
        /*
        public static void Click(string elementId, IWebDriver driver)
        {
            Click(By.LinkText(elementId), driver);
        }
        

        public static void Click(By by, IWebDriver driver)
        {
            IWebElement element = driver.FindElement(by);
            element.Highlight();
            element.Click();
        }
        
        public static void Highlight(this IWebElement element)
        {
            const int wait = 150;
            string orig = element.GetAttribute("style");
            //SetAttribute(element, "style", "color: yellow; border: 10px solid yellow; background-color: black;");
            //Thread.Sleep(wait);
            SetAttribute(element, "style", "color: black; border: 10px solid black; background-color: yellow;");
            Thread.Sleep(wait);
            //SetAttribute(element, "style", "color: yellow; border: 10px solid yellow; background-color: black;");
            //Thread.Sleep(wait);
            //SetAttribute(element, "style", "color: black; border: 10px solid black; background-color: yellow;");
            //Thread.Sleep(wait);
            //SetAttribute(element, "style", orig);
        }
        */

        public static void Wait(double wait)
        {
            Thread.Sleep((int)(wait * 1000));
        }

        /*
        public static void waitforElementbyXPath(int time, string element, IWebDriver driver)
        {
            var wait = new WebDriverWait(driver, new TimeSpan(time));
            wait.Until(ExpectedConditions.ElementExists(By.XPath(element)));
        }
        */
        //public static void waitforElementbyId(int time, string elementid, IWebDriver driver)
        //{
        //    var wait = new WebDriverWait(driver, new TimeSpan(0, 0, time));
        //    wait.Until(ExpectedConditions.ElementExists(By.Id(elementid)));
        //    return;
        //}

        //#merge# added ClickSomething
        public static Boolean ClickSomething(string sBy, string sLocator, int iWait, IWebDriver _driver)
        {
            Console.WriteLine("ClickSomething Starting Function:  sBy: " + sBy + ", sLocator=" + sLocator + ", iWait=" +
                              iWait);
            Boolean bResult = false;
            //Check to see if the required element exists
            if (CheckElementExists(sBy, sLocator, iWait, _driver) == false)
            {
                //Unable to find dropdown - abort
                Console.WriteLine("ClickSomething Finished Function:  bResult=" + bResult +
                                  "  [WARNING]:  Element NOT found.");
                return bResult;
            }
            //Element exists - Click it
            switch (sBy.ToUpper())
            {
                case "ID":
                    //Console.WriteLine("     I'm going to click it now (by ID):  sLocator=" + sLocator);
                    _driver.FindElement(By.Id(sLocator)).Click();
                    break;
                case "XPATH":
                    //Console.WriteLine("     I'm going to click it now (by XPATH):  sLocator=" + sLocator);
                    _driver.FindElement(By.XPath(sLocator)).Click();
                    break;
                default:
                    //Unknown BY type - abort
                    Console.WriteLine("ClickSomething Finished Function:  bResult=" + bResult +
                                      "  [WARNING]:  Unknown BY type (" + sBy + ")");
                    bResult = false;
                    return bResult;
            }
            bResult = true;
            Console.WriteLine("ClickSomething Finished Function:  bResult=" + bResult);
            return bResult;
        }

        //#merge# added EnterTextBox
        public static Boolean EnterTextBox(string sBy, string sLocator, string sValue, int iWait, IWebDriver _driver)
        {
            Console.WriteLine("EnterTextBox Starting Function:  sBy={0}, sLocator={1}, sValue={2}, iWait={3}", sBy,
                sLocator, sValue, iWait.ToString());
            Boolean bResult = false;
            //Check to see if the required element exists
            if (CheckElementExists(sBy, sLocator, iWait, _driver) == false)
            {
                //Unable to find dropdown - abort
                Console.WriteLine("EnterTextBox Finished Function:  bResult=" + bResult +
                                  "  [WARNING]:  Textbox not found.");
                return bResult;
            }
            //Check to see if the required element is enabled
            if (CheckElementEnabled(sBy, sLocator, _driver) == false)
            {
                //Unable to find dropdown - abort
                Console.WriteLine("EnterTextBox Finished Function:  bResult=" + bResult +
                                  "  [WARNING]:  Textbox not enabled.");
                return bResult;
            }
            //Text box element exists - Enter the data in it
            switch (sBy.ToUpper())
            {
                case "ID":
                    _driver.FindElement(By.Id(sLocator)).SendKeys(sValue);
                    break;
                case "XPATH":
                    _driver.FindElement(By.XPath(sLocator)).SendKeys(sValue);
                    break;
                default: //Unknown BY type - abort
                    Console.WriteLine("EnterTextBox Finished Function:  bResult=" + bResult +
                                      "  [WARNING]:  Unknown BY type.");
                    return bResult;
            }
            bResult = true;
            Console.WriteLine("EnterTextBox Finished Function:  bResult=" + bResult);
            return bResult;
        }

        //#merge# added SelectDropDown
        public static Boolean SelectDropDown(string sId, string sValue, int iWait, IWebDriver _driver)
        {
            Console.WriteLine("SelectDropDown Starting Function:  sId:  " + sId + ", sValue:  " + sValue);
            Boolean bResult = false;
            //Check to see if the required dropdown exists
            if (CheckElementExists("ID", sId, 20, _driver) == false)
            {
                //Unable to find dropdown - abort
                Console.WriteLine("SelectDropDown Finished Function:  bResult=" + bResult.ToString() +
                                  "[WARNING SelectDropDown]:  Dropdown NOT found.");
                return bResult;
            }
            //Check to see if the required dropdown is enabled
            if (CheckElementEnabled("ID", sId, _driver) == false)
            {
                //Unable to find dropdown - abort
                Console.WriteLine("SelectDropDown Finished Function:  bResult=" + bResult.ToString() +
                                  "  [WARNING]:  Dropdown NOT enabled.");
                return bResult;
            }
            //check the required value is in the dropdown list
            string sXpath = "//*[@id='" + sId + "']/option";
            int iOptions = _driver.FindElements(By.XPath(sXpath)).Count();
            if (iOptions == 0)
            {
                //No dropdown options found
                Console.WriteLine("SelectDropDown Finished Function:  bResult=" + bResult.ToString() +
                                  "[WARNING]:  No dropdown options found for element:  sXpath=" + sXpath);
                return bResult;
            }
            Console.WriteLine("  iOptions:  " + iOptions.ToString());
            string sOption;
            string sOptions = "";
            for (int i = 0; i <= (iOptions - 1); i++) // Go through all the options in the dropdown list
            {
                sOption = _driver.FindElements(By.XPath(sXpath)).ElementAt(i).Text;
                if (sOption == sValue)
                {
                    bResult = true; //Found the required dropdown list entry - click it
                    _driver.FindElements(By.XPath(sXpath)).ElementAt(i).Click();
                    Console.WriteLine("Finished Function:   bResult=" + bResult.ToString() + ", sOption=" + sOption +
                                      ", i=" + i.ToString() + ", sXpath=" + sXpath);
                    return bResult;
                }
                else
                {
                    if (i == 0)
                    {
                        sOptions = sOption;
                    }
                    else
                    {
                        sOptions = "," + sOption;
                    }
                }
            }
            //Option not found - report error
            Console.WriteLine("Finished Function:   bResult=" + bResult.ToString() +
                              "  [WARNING] Requested dropdown option not found in list:  sXpath=" + sXpath +
                              ", sOptions=" + sOptions);
            return bResult;
        }

        //#merge# added SelectDropDownByOptionNumber
        public static Boolean SelectDropDownByOptionNumber(string sId, int iOption, int iWait, IWebDriver _driver)
        {
            Console.WriteLine("SelectDropDownByOptionNumber Starting Function:  sId=" + sId + ", iOption=" +
                              iOption.ToString() + ", iWait=" + iWait.ToString());
            Boolean bResult = false;
            //Check to see if the required dropdown exists
            if (CheckElementExists("ID", sId, 20, _driver) == false)
            {
                //Unable to find dropdown - abort
                Console.WriteLine("SelectDropDownByOptionNumber Finished Function:  bResult=" + bResult.ToString() +
                                  "  [WARNING]:  Dropdown NOT found.");
                return bResult;
            }
            //Check to see if the required dropdown is enabled
            if (CheckElementEnabled("ID", sId, _driver) == false)
            {
                //Unable to find dropdown - abort
                Console.WriteLine("SelectDropDownByOptionNumber Finished Function:  bResult=" + bResult.ToString() +
                                  "  [WARNING]:  Dropdown NOT enabled.");
                return bResult;
            }
            //Dropdown exists - check the required option number is in the dropdown list
            string sXpath = "//*[@id='" + sId + "']/option";
            int iOptions = _driver.FindElements(By.XPath(sXpath)).Count();
            //Console.WriteLine("  iOptions:  " + iOptions.ToString());
            if (iOptions == 0)
            {
                //No dropdown options found
                Console.WriteLine("SelectDropDownByOptionNumber Finished Function:  bResult=" + bResult.ToString() +
                                  "  [WARNING]:  No dropdown options found for element:  sXpath=" + sXpath);
                return bResult;
            }
            if (iOption > iOptions)
            {
                //Invalid option ID
                Console.WriteLine(
                    "SelectDropDownByOptionNumber Finished Function:  bResult={0}  [WARNING]:  iOption({1}) greater than iOptions=({2}), sXpath={3}",
                    bResult.ToString(), iOption.ToString(), iOptions.ToString(), sXpath);
                return bResult;
            }
            bResult = true; //Pick the required dropdown option and return true
            _driver.FindElements(By.XPath(sXpath)).ElementAt(iOption).Click();
            Console.WriteLine("SelectDropDownByOptionNumber Finished Function:  bResult=" + bResult.ToString() +
                              ", OptionText=" +
                              _driver.FindElement(By.Id(sId)).FindElements(By.XPath("//option")).ElementAt(iOption).Text);
            return bResult;
        }

        #endregion


        #region Paramters

        public static string goAndGet(string xmlPath, string sEnvironment = "")
        {
            //Console.WriteLine("goAndGet.xmlpath = " + xmlPath);

            var outPutDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase); //Get the path of the solution
            var XMLdoc = Path.Combine(outPutDirectory, "parametersTestSand.xml"); //Set the path 
            var reader = new XmlTextReader(XMLdoc);
            var xmlDoc = new XmlDocument();
            xmlDoc.Load(reader);
            if (sEnvironment == "")
            {
                XmlNode environment = xmlDoc.SelectSingleNode("SAVED/ENVIRONMENT");
                XmlNode node = xmlDoc.SelectSingleNode("SAVED/" + environment.InnerText + "/" + xmlPath);
                Console.WriteLine("goAndGet.return = " + node.InnerText);
                return node.InnerText;
            }
            else
            {
                sEnvironment = sEnvironment.ToUpper();
                XmlNode node = xmlDoc.SelectSingleNode("SAVED/" + sEnvironment + "/" + xmlPath);
                Console.WriteLine("goAndGet.return = " + node.InnerText);
                return node.InnerText;
            }
        }

        public static int randomGen(int start, int end)
        {
            //int intdate = Convert.ToInt32(DateTime.Now);
            Random random = new Random();
            int randomNumber = random.Next(start, end);
            return randomNumber;
        }



        #endregion


        #region Dates

        public static DateTime RoundUp(DateTime dt, TimeSpan d)
        {
            return new DateTime(((dt.Ticks + d.Ticks - 1) / d.Ticks) * d.Ticks);
        }

        public static DateTime RoundDown(DateTime dt, TimeSpan d)
        {
            return new DateTime((dt.Ticks / d.Ticks) * d.Ticks, 0);
        }

        public static string GetSortDate()
        {
            string sDate = DateTime.Now.ToString("yyyyMMddHHmmss");
            return sDate;
        }

        public static DateTime GetDate(string sValue)
        {
            Console.WriteLine("GetDate Starting Function:  sValue=" + sValue);
            DateTime dDate;
            if (DateTime.TryParse(sValue, out dDate))
            {
                //It's a date - return the date
                Console.WriteLine("GetDate Function:  Converted '{0}' to {1} ({2}).", sValue, dDate, dDate.Kind);
                return dDate;
            }
            else
            {
                //It's not a date - calculate the date
                string[] slDate = sValue.Split('|');
                string sPorM = slDate[0].ToString();
                int iNumber = Convert.ToInt32(slDate[1].ToString());
                string sSize = slDate[2];
                if (sPorM == "-")
                {
                    switch (sSize.ToUpper())
                    {
                        case "MINUTES":
                            dDate = RoundUp(DateTime.Now.AddMinutes(-iNumber), TimeSpan.FromMinutes(5));
                            break;
                        case "HOURS":
                            dDate = RoundUp(DateTime.Now.AddHours(-iNumber), TimeSpan.FromMinutes(5));
                            break;
                        case "DAYS":
                            dDate = RoundUp(DateTime.Now.AddDays(-iNumber), TimeSpan.FromMinutes(5));
                            break;
                        case "MONTHS":
                            dDate = RoundUp(DateTime.Now.AddMonths(-iNumber), TimeSpan.FromMinutes(5));
                            break;
                        case "YEARS":
                            dDate = RoundUp(DateTime.Now.AddYears(-iNumber), TimeSpan.FromMinutes(5));
                            break;
                        default:
                            //fail - unable to work our size
                            Console.WriteLine(
                                "GetDate Function [WARNING] Unable to caluculate date:  sValue={0}, sPorM={1}, iNumber={2}, sSize={3}).",
                                sValue, sPorM, iNumber.ToString(), sSize);
                            break;
                    }
                }
                else
                {
                    switch (sSize.ToUpper())
                    {
                        case "MINUTES":
                            dDate = RoundUp(DateTime.Now.AddMinutes(iNumber), TimeSpan.FromMinutes(5));
                            break;
                        case "HOURS":
                            dDate = RoundUp(DateTime.Now.AddHours(iNumber), TimeSpan.FromMinutes(5));
                            break;
                        case "DAYS":
                            dDate = RoundUp(DateTime.Now.AddDays(iNumber), TimeSpan.FromMinutes(5));
                            break;
                        case "MONTHS":
                            dDate = RoundUp(DateTime.Now.AddMonths(iNumber), TimeSpan.FromMinutes(5));
                            break;
                        case "YEARS":
                            dDate = RoundUp(DateTime.Now.AddYears(iNumber), TimeSpan.FromMinutes(5));
                            break;
                        default:
                            //fail - unable to work our size
                            Console.WriteLine(
                                "GetDate Function [WARNING] Unable to caluculate date:  sValue={0}, sPorM={1}, iNumber={2}, sSize={3}).",
                                sValue, sPorM, iNumber.ToString(), sSize);
                            break;
                    }
                }
                //Success return caulculated date
                Console.WriteLine("GetDate Function:  dDate={0}, sValue={1}, iNumber={2}, sSize={3}",
                    dDate.ToLongTimeString(), sValue, iNumber.ToString(), sSize);
                return dDate;
            }
        }

        #endregion


        #region Csv

        public static string[,] LoadCsv(string sFileName)
        {
            Console.WriteLine("LoadCsv Function:  sFileName=" + sFileName);
            string whole_File = System.IO.File.ReadAllText(sFileName); // Get the File's text.
            whole_File = whole_File.Replace('\n', '\r'); // Split into lines.
            string[] lines = whole_File.Split(new char[] { '\r' }, StringSplitOptions.RemoveEmptyEntries);
            int num_rows = lines.Length; // See how many rows and columns there are.
            int num_cols = lines[0].Split(',').Length; // See how many rows and columns there are.
            string[,] values = new string[num_rows, num_cols]; // Allocate the data array.
            for (int r = 0; r < num_rows; r++) // Load the array.
            {
                string[] line_r = lines[r].Split(',');
                for (int c = 0; c < num_cols; c++)
                {
                    values[r, c] = line_r[c];
                    //Console.WriteLine("values[r{0}, c{1}]:  {2}", r, c, values[r, c].ToString());
                }
            }
            // Return the values.
            return values;
        }

        public static string ConvertToCSV(DataSet MyDataSet, string MyTableName, string sFileLocation)
        {
            Console.WriteLine("ConvertToCSV Starting Function:  MyDataSet={0},  MyTableName={1}, sFileLocation={2}",
                MyDataSet.DataSetName.ToString(), MyTableName, sFileLocation);
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            foreach (DataTable table in MyDataSet.Tables)
            {
                Console.WriteLine("     TableName=" + table.TableName.ToString());
            }

            DataTable MyDataTable = MyDataSet.Tables[MyTableName];
            string[] columnNames = MyDataTable.Columns.Cast<DataColumn>().Select(column => column.ColumnName).ToArray();
            sb.AppendLine(string.Join(",", columnNames));
            foreach (DataRow row in MyDataTable.Rows)
            {
                var fields = row.ItemArray.Select(field => field.ToString()).ToArray();
                sb.AppendLine(string.Join(",", fields));
            }
            //Console.WriteLine("     Result:  " + sb.ToString());
            try
            {


                File.WriteAllText(sFileLocation, sb.ToString());
                Console.WriteLine("ConvertToCSV Finished Function:  Result = Success");
            }
            catch
            {
                Console.WriteLine("ConvertToCSV Finished Function:  Result = Failed");
            }
            return sb.ToString();
        }

        #endregion


        #region Directories

        /*
        public static Boolean FolderCheckCreate(string sFolder)
        {
            Console.WriteLine("    sFolder=" + sFolder);
            bool bResult = false;
            string sMyFolder = sFolder;
            if (sMyFolder.Substring(sMyFolder.Length - 1, 1) == "\\")
            {
                sMyFolder = sMyFolder.Substring(0, sMyFolder.Length - 1);
            }
            ;
            Console.WriteLine("    sMyFolder=" + sMyFolder);
            if (Directory.Exists(sMyFolder) == false)
            {
                Console.WriteLine("     - doesn't exist");
                //Folder doesn't exist check and create each level of the directory required
                //Split the overall path into each folder level, so that each folder level can be checked and created
                List<string> slFolders = new List<string>();
                slFolders = sMyFolder.Split('\\').ToList();
                int iCount = slFolders.Count();
                Console.WriteLine("    iCount=" + iCount.ToString());
                string sCheckFolder = slFolders.ElementAt(0).ToString();
                for (int i = 1; i < iCount; i++) // Go through all the error results and report each error
                {
                    sCheckFolder = sCheckFolder.ToString() + "\\" + slFolders.ElementAt(i).ToString();
                    if (Directory.Exists(sCheckFolder) == false)
                    {
                        //Directory doesn't exist - create it
                        try
                        {
                            DirectoryInfo di = Directory.CreateDirectory(sCheckFolder);
                            Console.WriteLine("    Created=" + sCheckFolder);
                            bResult = true;
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("FolderCheckCreate:  FAILED to create directory: {0}", e.ToString());
                            Console.WriteLine("    " + sCheckFolder);
                            bResult = false;
                            break;
                        }
                    }
                    else
                    {
                        Console.WriteLine("    Exists=" + sCheckFolder);
                    }
                }
            }
            Console.WriteLine("FolderCheckCreate:  bResult=" + bResult.ToString() + ", sFolder=" + sFolder);
            return bResult;
        }
        */
        #endregion


        #region Database

        public static string GetQueryResult(string sSQL)
        {
            Console.WriteLine("GetQueryResult Starting Function:  sSQL=  " + sSQL);
            string sResult;
            using (SqlConnection conn = new SqlConnection(goAndGet("DBCONNECTION")))
            {
                conn.Open();
                using (SqlCommand command = new SqlCommand(sSQL))
                {
                    command.CommandType = System.Data.CommandType.Text;
                    command.Connection = conn;
                    Object obj = command.ExecuteScalar();
                    sResult = obj.ToString();
                }
                conn.Close();
                Console.WriteLine("GetQueryResult - Result:  " + sResult);
                return sResult;
            }
        }

        public static bool QueryToDataSet(string sSQL, ref DataSet ds)
        {
            Console.WriteLine("QueryToArray:  Starting Function");
            Console.WriteLine("     sSQL:  " + sSQL.ToString());
            bool bResult = false;
            SqlConnection SQLconn = new SqlConnection(GenericFunctions.goAndGet("DBCONNECTION"));
            SqlCommand SQLcommand = new SqlCommand(sSQL, SQLconn);
            SqlDataAdapter adapter = new SqlDataAdapter(SQLcommand);
            SQLconn.Close();
            DataSet SqlDataSet = new DataSet();
            Console.WriteLine("Table Count b4 = " + SqlDataSet.Tables.Count.ToString());
            try
            {
                adapter.Fill(SqlDataSet, "SqlDataTable");
            }
            catch
            {
                //Fill error - abort and return false
                Console.WriteLine("QueryToArray:  Finised Function [FAIL - Fill Error]");
                return bResult;
            }
            if (SqlDataSet.Tables.Count == 0)
            {
                //NO data found - abort and return false
                Console.WriteLine("QueryToArray:  Finised Function [FAIL - Table Count = 0]");
                return bResult;
            }
            else
            {
                //Data found - return the DataSet
                ds = SqlDataSet;
                bResult = true;
                Console.WriteLine("Table Count after = " + SqlDataSet.Tables.Count.ToString());
                Console.WriteLine("QueryToArray:  Finised Function [SUCCESS]");
                return bResult;
            }
        }

        #endregion


        #region ReactCell

        public static bool GetSubjectsShortenedURL(string sStudy, string sSite, string sSubjectID, out string sShortURL)
        {
            //Find the subjects shortened URL from the database
            Console.WriteLine("GetSubjectsShortenedURL Starting Function:  sStudy=" + sStudy + ", sSite = " + sSite +
                              ", sSubjectID=" + sSubjectID);
            string sSQL =
                "SELECT ul.subject_unique_lookup FROM React30.dbo.tbl_subject su left outer join React30.dbo.tbl_site si on su.site_id = si.site_id left outer join React30.dbo.tbl_study st on si.study_id = st.study_id left outer join React30.dbo.tbl_subject_unique_lookup ul on su.subject_id = ul.subject_id WHERE st.name = '" +
                sStudy + "' and si.name = '" + sSite + "' and su.user_subject_id = '" + sSubjectID + "'";
            sShortURL = GetQueryResult(sSQL);
            if (sShortURL == "")
            {
                //Failed to get shortened URL from the database
                Console.WriteLine(
                    "GetSubjectsShortenedURL Finished Function[FAIL]:  Failed to get shortened URL from the database");
                return false;
            }
            sShortURL = GenericFunctions.goAndGet("SHORTENEDURL") + sShortURL;
            Console.WriteLine("GetSubjectsShortenedURL Finished Function[SUCCESS]:  sShortURL=" + sShortURL);
            return true;
        }

        public static bool GetScheduledMessages(string sStudy, string sSubjectID, string sFileName)
        {
            string sExpectedMessages = "";
            string sReturnedMessages = "";

            //Create dataset
            var MyDataSet = new DataSet();

            //Load data from CSV to build string for in statement
            string[,] saValues = GenericFunctions.LoadCsv(sFileName);
            int iRows = saValues.GetUpperBound(0);

            for (int r = 1; r <= iRows; r++)
            {
                //Checks saValues isn't blank
                if (saValues[r, 1].ToString() != "")
                {
                    sExpectedMessages = sExpectedMessages + "'" + saValues[r, 0].ToString() + "',";
                }
                else
                {
                    //As soon as "" are returned set r to the same as iRows to jump out of the loop
                    r = iRows;
                }
            }
            //Strip last comma
            sExpectedMessages = sExpectedMessages.Remove(sExpectedMessages.Length - 1);

            //Find the subjects scheduled message from the database - Note this returns the scheduled messages in the order they will be sent so your File must list the messages in send order            
            string sSQL =
                "SELECT mt.name FROM React30.dbo.tbl_outbound_msg om INNER JOIN React30.dbo.tbl_msg_template mt ON om.msg_template_id = mt.msg_template_id INNER JOIN React30.dbo.tbl_actual_event ae ON ae.actual_event_id = om.actual_event_id INNER JOIN React30.dbo.tbl_subject sub ON sub.subject_id = ae.subject_id INNER JOIN React30.dbo.tbl_site si ON si.site_id = sub.site_id INNER JOIN React30.dbo.tbl_study st ON st.study_id = si.study_id WHERE st.name = '" +
                sStudy + "' AND sub.user_subject_id = '" + sSubjectID + "'and mt.name in (" + sExpectedMessages + ") ORDER BY send_time;";
            var sScheduledMessages = QueryToDataSet(sSQL, ref MyDataSet);
            if (sScheduledMessages == true)
            {
                //Extract messages from dataset and build string
                DataTable dt = MyDataSet.Tables["SqlDataTable"];
                for (int row = 0; row < dt.Rows.Count; ++row)
                {
                    sReturnedMessages = sReturnedMessages + "'" + dt.Rows[row][0].ToString() + "',";
                }
                //Strip last comma
                sReturnedMessages = sReturnedMessages.Remove(sReturnedMessages.Length - 1);
                //Compare expected against scheduled messages
                if (sExpectedMessages == sReturnedMessages)
                {
                    ReportTheTestPassed("The correct messages have been scheduled");
                    return true;
                }
                else
                {
                    ReportTheTestFailed("Scheduled messages are incorrect");
                    _startUp.stepsFailed++;
                    //Export the messages as csv to current users desktop - enables checking of incorrect scheduled messages
                    string userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
                    userName = userName.Remove(0, 5);
                    File.WriteAllText(@"C:\\Users\\" + userName + "\\Desktop\\" + sStudy + "_ScheduledMessages.csv", sReturnedMessages);
                }
            }
            return false;
        }

        public static bool GetCancelledMessages(string sStudy, string sSubjectID, string sFileName)
        {
            string sExpectedMessages = "";
            string sReturnedMessages = "";

            //Create dataset
            var MyDataSet = new DataSet();

            //Load data from CSV to build string for in statement
            string[,] saValues = GenericFunctions.LoadCsv(sFileName);
            int iRows = saValues.GetUpperBound(0);

            for (int r = 1; r <= iRows; r++)
            {
                //Checks saValues isn't blank
                if (saValues[r, 1].ToString() != "")
                {
                    sExpectedMessages = sExpectedMessages + "'" + saValues[r, 1].ToString() + "',";
                }
                else
                {
                    //As soon as "" are returned set r to the same as iRows to jump out of the loop
                    r = iRows;
                }
            }
            //Strip last comma
            sExpectedMessages = sExpectedMessages.Remove(sExpectedMessages.Length - 1);

            //Find the subjects cancelled message from the database - Note this returns the cancelled messages in the order they would have been sent so your File must list the messages in send order            
            string sSQL =
                "SELECT mt.name FROM React30.dbo.xlog_tbl_outbound_msg xom INNER JOIN(SELECT distinct outbound_msg_id,sub.subject_id FROM React30.dbo.xlog_tbl_outbound_msg om INNER JOIN React30.dbo.tbl_msg_template mt ON om.msg_template_id = mt.msg_template_id INNER JOIN React30.dbo.tbl_actual_event ae ON ae.actual_event_id = om.actual_event_id INNER JOIN React30.dbo.tbl_subject sub ON sub.subject_id = ae.subject_id	INNER JOIN React30.dbo.tbl_site si ON si.site_id = sub.site_id INNER JOIN React30.dbo.tbl_study st ON st.study_id = si.study_id	WHERE st.name = '" +
                sStudy + "' AND sub.user_subject_id = '" + sSubjectID + "'and mt.name in (" + sExpectedMessages + ")) I_S ON xom.outbound_msg_id = I_S.outbound_msg_id AND audit_action = 'DELETE' INNER JOIN React30.dbo.tbl_subject sub ON I_S.subject_id = sub.subject_id INNER JOIN React30.dbo.tbl_msg_template mt ON mt.msg_template_id= xom.msg_template_id ORDER BY send_time;";
            var sCancelledMessages = QueryToDataSet(sSQL, ref MyDataSet);
            if (sCancelledMessages == true)
            {
                //Extract messages from dataset and build string
                DataTable dt = MyDataSet.Tables["SqlDataTable"];
                for (int row = 0; row < dt.Rows.Count; ++row)
                {
                    sReturnedMessages = sReturnedMessages + "'" + dt.Rows[row][0].ToString() + "',";
                }
                //Strip last comma
                sReturnedMessages = sReturnedMessages.Remove(sReturnedMessages.Length - 1);
                //Compare expected against cancelled messages
                if (sExpectedMessages == sReturnedMessages)
                {
                    ReportTheTestPassed("The correct messages have been cancelled");
                    return true;
                }
                else
                {
                    ReportTheTestFailed("Cancelled messages are incorrect");
                    _startUp.stepsFailed++;
                    //Export the messages as csv to current users desktop - enables checking of incorrect cancelled messages
                    string userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
                    userName = userName.Remove(0, 5);
                    File.WriteAllText(@"C:\\Users\\" + userName + "\\Desktop\\" + sStudy + "_CancelledMessages.csv", sReturnedMessages);
                }
            }
            return false;
        }

        public static bool CheckMessagesSent(string sStudy, string sSubjectID)
        {
            //set starttime
            DateTime StartTime = Convert.ToDateTime(Convert.ToDateTime(GenericFunctions.RoundDown(DateTime.Parse(DateTime.Now.ToShortTimeString()), TimeSpan.FromMinutes(10))).ToShortTimeString());
            string sStartTime = StartTime.ToString("yyyy-MM-dd HH:mm:ss") + ".000";
            //Set endtime
            DateTime EndTime = Convert.ToDateTime(Convert.ToDateTime(GenericFunctions.RoundUp(DateTime.Parse(DateTime.Now.ToShortTimeString()), TimeSpan.FromMinutes(10))).ToShortTimeString());
            string sEndTime = EndTime.ToString("yyyy-MM-dd HH:mm:ss") + ".000";

            //Create dataset for query results
            var MyDataSet = new DataSet();

            //Find the subjects scheduled message from the database           
            string sSQL =
                "SELECT pe.name[Event], mt.name[Msg Name], ty.name[Msg Type] FROM React30..tbl_outbound_msg om " +
                "INNER JOIN React30..tbl_msg_template mt ON om.msg_template_id = mt.msg_template_id " +
                "INNER JOIN React30..tbl_actual_event ae ON ae.actual_event_id = om.actual_event_id " +
                "INNER JOIN React30..tbl_subject sub ON sub.subject_id = ae.subject_id " +
                "INNER JOIN React30..tbl_site si ON sub.site_id = si.site_id " +
                "INNER JOIN React30..tbl_study st ON st.study_id = si.study_id " +
                "INNER JOIN React30..tlkp_msg_type ty on mt.msg_type_id = ty.msg_type_id " +
                "INNER JOIN React30..tbl_planned_event pe on ae.planned_event_id = pe.planned_event_id " +
                "WHERE st.name = '" + sStudy + "' AND sub.user_subject_id = '" + sSubjectID + "' AND om.send_time BETWEEN '" + sStartTime + "' and '" + sEndTime + "' ORDER BY om.send_time;";
            var sScheduledMessages = GenericFunctions.QueryToDataSet(sSQL, ref MyDataSet);
            if (sScheduledMessages == true)
            {
                //Loop through messages
                DataTable dt = MyDataSet.Tables["SqlDataTable"];
                for (int row = 0; row < dt.Rows.Count; ++row)
                {
                    //Create dataset for new query results
                    var ds = new DataSet();

                    //Extract messages from dataset to populate variables
                    var Event = dt.Rows[row][0].ToString();
                    var msgName = dt.Rows[row][1].ToString();
                    var msgType = dt.Rows[row][2].ToString();

                    //Generate SQL for sent message query
                    string SQL =
                        "SELECT stat.name[React Status], CASE WHEN om.jit_status_id = 0 THEN 'Default'	WHEN om.jit_status_id = 1 THEN 'Placeholder' WHEN om.jit_status_id = -1 THEN 'Instant Placeholder' WHEN om.jit_status_id = 2 THEN 'Pending' WHEN om.jit_status_id = 3 THEN 'Sent' END [JIT Status] " +
                        "FROM React30..tbl_outbound_msg om " +
                        "INNER JOIN React30..tbl_msg_template mt ON om.msg_template_id = mt.msg_template_id " +
                        "INNER JOIN React30..tbl_actual_event ae ON ae.actual_event_id = om.actual_event_id " +
                        "INNER JOIN React30..tbl_subject sub ON sub.subject_id = ae.subject_id " +
                        "INNER JOIN React30..tbl_site si ON sub.site_id = si.site_id " +
                        "INNER JOIN React30..tbl_study st ON st.study_id = si.study_id " +
                        "INNER JOIN React30..tlkp_status stat on om.react_status_id = stat.status_id " +
                        "INNER JOIN React30..tbl_planned_event pe on ae.planned_event_id = pe.planned_event_id " +
                        "WHERE st.name = '" + sStudy + "' AND sub.user_subject_id = '" + sSubjectID + "' AND pe.name = '" + Event + "' AND mt.name = '" + msgName + "' ORDER BY om.send_time;";
                    var sSentMessages = GenericFunctions.QueryToDataSet(SQL, ref ds);
                    if (sSentMessages == true)
                    {
                        DataTable dt2 = ds.Tables["SqlDataTable"];

                        var reactStatus = dt2.Rows[0][0].ToString();
                        var jitStatus = dt2.Rows[0][1].ToString();
                        //Check email has been sent
                        if (msgType == "Email")
                        {
                            if (reactStatus == "Complete" && jitStatus == "Sent")
                            {
                                GenericFunctions.ReportTheTestPassed("The " + msgName + " email has been sent successfully");
                                return true;
                            }
                            else
                            {
                                GenericFunctions.ReportTheTestFailed("The " + msgName + " email has NOT been sent!");
                            }
                        }
                        //Check SMS has been sent
                        else if (msgType == "SMS")
                        {
                            if (reactStatus == "Scheduled" && jitStatus == "Sent")
                            {
                                GenericFunctions.ReportTheTestPassed("The " + msgName + " SMS has been sent successfully");
                                return true;
                            }
                            else
                            {
                                GenericFunctions.ReportTheTestFailed("The " + msgName + " SMS has NOT been sent!");
                            }
                        }

                    }
                }

            }
            return false;
        }

        #endregion


        #region Reporting

        public static void WriteReportHeader(IWebDriver driver)
        {
            var reportmsg = new System.IO.StreamWriter(_startUp.resultsFileName, true);
            reportmsg.WriteLine("<h2>");
            reportmsg.WriteLine("SmokeTest: React & ReactCell");
            reportmsg.WriteLine("</h2>");
            reportmsg.WriteLine("<h3>Test Report Summary</h3>");
            reportmsg.WriteLine("<br>");
            reportmsg.WriteLine("<strong>Environment: </strong>" + goAndGet("Envname"));
            reportmsg.WriteLine("<br>");
            reportmsg.WriteLine("<strong>Browser: </strong>" + goAndGet("SeleniumBrowser"));
            reportmsg.WriteLine("<br>");
            reportmsg.WriteLine("<strong>Date: </strong>" + System.DateTime.Now.ToShortDateString());
            reportmsg.WriteLine("<br>");
            reportmsg.WriteLine("<strong>Time: </strong>" + System.DateTime.Now.ToShortTimeString());
            reportmsg.WriteLine("<br>");
            reportmsg.Close();
        }

        public static void WriteTestNameToReport(string testcasename)
        {
            var reportmsg = new System.IO.StreamWriter(_startUp.resultsFileName, true);
            reportmsg.WriteLine("<h4>");
            reportmsg.WriteLine(testcasename);
            reportmsg.WriteLine("</h4>");
            reportmsg.Close();
            System.Console.WriteLine(testcasename);
        }

        public static void ReportTheTestPassed(string Message)
        {
            var reportmsg = new System.IO.StreamWriter(_startUp.resultsFileName, true);
            string Date = DateTime.Now.ToString("dd-MM-yyyy HH:mm.ss");


            reportmsg.WriteLine("<font face=\"verdana\" size=\"2\" color=\"#000000\">[" + Date + "]</font> <font face=\"verdana\" size=\"2\" color=\"#526F35\">Pass: " + Message + "</font><br>");
            reportmsg.Close();
            System.Console.WriteLine("Pass: " + Message);
        }

        public static void ReportSmokeTestPass()
        {
            System.IO.StreamWriter reportmsg = new System.IO.StreamWriter(_startUp.resultsFileName, true);
            reportmsg.WriteLine("<br><h4>");
            reportmsg.WriteLine("[" + System.DateTime.Now.ToShortDateString() + " " +
                                System.DateTime.Now.ToLongTimeString() + "] " +
                                "<span style=\"color: green;\"><strong>All of the smoke test passed</strong></span>");
            reportmsg.WriteLine("</h4><br>");
            reportmsg.WriteLine("<br>");
            reportmsg.Close();
            System.Console.WriteLine("All of the smoke test passed");
        }

        public static void ReportSmokeTestFailed()
        {
            System.IO.StreamWriter reportmsg = new System.IO.StreamWriter(_startUp.resultsFileName, true);
            reportmsg.WriteLine("<br><h4>");
            reportmsg.WriteLine("[" + System.DateTime.Now.ToShortDateString() + " " +
                                System.DateTime.Now.ToLongTimeString() + "] " +
                                "<span style=\"color: green;\"><strong>One or more tess failed in the smoke test.</strong></span>");
            reportmsg.WriteLine("<br>");
            reportmsg.WriteLine("[" + System.DateTime.Now.ToShortDateString() + " " +
                                System.DateTime.Now.ToLongTimeString() + "] " +
                                "<span style=\"color: green;\"><strong>Please check the smoke test output for clarification.</strong></span>");
            reportmsg.WriteLine("</h4><br>");
            reportmsg.WriteLine("<br>");
            reportmsg.Close();
            System.Console.WriteLine("One or more tests failed. Please check the results for clarification");
        }

        public static void createReportSummaryLayout()
        {
            var createReportSummary = new System.IO.StreamWriter(_startUp.testSummaryFileName, false);
            ;
            createReportSummary.WriteLine("<h1>Test Report Summary</h1>");
            createReportSummary.WriteLine("     ");
            createReportSummary.WriteLine("<strong>Project: </strong>" + "ReactCell"); //+ product);
            createReportSummary.WriteLine("     ");
            createReportSummary.WriteLine("<strong>Version: </strong>" + "V1.00"); //+ version);
            createReportSummary.WriteLine("     ");
            createReportSummary.WriteLine("<strong>Environment: </strong>" + "Sand"); //+ environment);
            createReportSummary.WriteLine("     ");
            createReportSummary.WriteLine("<strong>Browser: </strong>" + "FireFox"); //+ browser);
            createReportSummary.WriteLine("     ");
            createReportSummary.WriteLine("<strong>Date: </strong>" + System.DateTime.Now.ToShortDateString());
            createReportSummary.WriteLine("     ");
            createReportSummary.WriteLine("<strong>Time: </strong>" + System.DateTime.Now.ToShortTimeString());
            createReportSummary.WriteLine("     ");
            createReportSummary.Close();
        }

        public static void ReportTheTestFailed(string Message)
        {
            var reportmsg = new System.IO.StreamWriter(_startUp.resultsFileName, true);
            string Date = DateTime.Now.ToString("dd-MM-yyyy HH:mm.ss");


            reportmsg.WriteLine("<font face=\"verdana\" size=\"2\" color=\"#000000\">[" + Date + "]</font> <font face=\"verdana\" size=\"2\" color=\"#DB2929\">Fail: " + Message + "</font><br>");
            reportmsg.Close();
            System.Console.WriteLine("FAIL: " + Message);

        }

        public static void Warning(string Message)
        {
            var reportmsg = new System.IO.StreamWriter(_startUp.resultsFileName, true);
            string Date = DateTime.Now.ToString("dd-MM-yyyy HH:mm.ss");


            reportmsg.WriteLine("<font face=\"verdana\" size=\"2\" color=\"#000000\">[" + Date + "]</font> <font face=\"verdana\" size=\"2\" color=\"#FFA500\">Warning: " + Message + "</font><br>");
            reportmsg.Close();
            System.Console.WriteLine("FAIL: " + Message);
        }

        public static void GenerateGraph()
        {

            DateTime dt = DateTime.Now;

            string now = dt.ToString("dd/MM/yyyy hh:mm");
            string Results = "['" + now + "', " + _startUp.Passes + ", " + _startUp.Fails + ", " +
                             _startUp.Warnings + "],";
            string path = _startUp.RunEnv + "ResultsOverTime.txt";

            // This text is added only once to the File. 
            if (!File.Exists(path))
            {
                // Create a File to write to. 
                using (StreamWriter sw = File.CreateText(path))
                {
                    sw.WriteLine(Results);
                }
            }
            else
            {

                // This text is always added, making the File longer over time 
                // if it is not deleted. 
                using (StreamWriter sw = File.AppendText(path))
                {
                    sw.WriteLine(Results);
                  
                }

            }

        }

        public static void ClickOnButtonByXPath(string button, IWebDriver driver)
        {

            int counter = 1;


            while (!driver.FindElement(By.XPath(button)).Displayed)
            {

                Thread.Sleep(1000);

                if (counter == 100)
                {
                    break;
                }

            }

            driver.FindElement(By.XPath(button)).Click();

        }


        public static void GenerateLineGraph()
        {
            //https://developers.google.com/chart/interactive/docs/gallery/linechart
        }

        #endregion


        #region screenCapture

        public static void GetScreenShot(IWebDriver driver, string myFileName)
        {
            var date = DateTime.Now.ToString("dd-MM-yyyy HH_mm_ss");
            var myDate = ReplaceAllTextString(date);
            var Filename = myFileName
                              + myDate + DateTime.Now.ToString("ss") + ".png";

            // Specify a "currently active folder"
            const string activeDir = @"c:\";

            //Create a new subfolder inder the current active folder
            string newPath = System.IO.Path.Combine(activeDir, "myScreenShots");

            //Create the subfolder
            System.IO.Directory.CreateDirectory(newPath);

            //Set the path
            System.IO.Directory.SetCurrentDirectory(newPath);

            // newPath = System.IO.Path.Combine(newPath, FileName);

            if (!System.IO.File.Exists(newPath))
            {
                ((ITakesScreenshot)driver).GetScreenshot().SaveAsFile(Filename, OpenQA.Selenium.ScreenshotImageFormat.Png);
            }
            Functions.Reporting.TestStep("<br><a href=" + activeDir + Filename + " target=new><img src=" + activeDir + Filename + " width=500" + "Title=\"Click to view larger image\" border=0></a>");
        }


        public static void TakePicture(IWebDriver driver)
        {

            // Specify a "currently active folder"
            string activeDir = @"";

            ////Create a new subfolder inder the current active folder
            //string newPath = System.IO.Path.Combine(activeDir, "myScreenDumps");

            ////Create the subfolder
            //System.IO.Directory.CreateDirectory(newPath);

            ////Set the path
            //System.IO.Directory.SetCurrentDirectory(newPath);

            Random r = new Random();
            int y = r.Next(0000000, 99999999);
            string random = y.ToString(CultureInfo.InvariantCulture);
            string newFileName = ("itc" + y + ".png"); //Set our Filename

            Screenshot screenshot = ((ITakesScreenshot)driver).GetScreenshot();
            string ss = screenshot.AsBase64EncodedString;
            byte[] screenshotAsByteArray = screenshot.AsByteArray;

            screenshot.SaveAsFile(activeDir + newFileName, OpenQA.Selenium.ScreenshotImageFormat.Png);


            //_driver.GetScreenshot().SaveAsFile(newFileName, System.Drawing.Imaging.ImageFormat.Png);

            //Combine the new File name with the path
            //newPath = System.IO.Path.Combine(newPath, newFileName);

            //Check that the File has been created

            //if (!System.IO.File.Exists(newPath))
            //{
            //    using (System.IO.FileStream fs = System.IO.File.Create(newPath))
            //    {
            //        for (byte i = 0; i < 100; i++)
            //        {
            //            fs.WriteByte(i);
            //        }

        }

       

        
        #endregion
 
    }
}





