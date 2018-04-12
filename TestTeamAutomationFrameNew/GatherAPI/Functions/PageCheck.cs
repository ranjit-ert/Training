using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.IE;
using NUnit.Framework;
using OpenQA.Selenium.Support.Events;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using TestTeamAutomationFrameNew;
using TestTeamAutomationFrameNew.Functions;

namespace TestTeamAutomationFrameNew
{
    public class PageCheck
    {

        public static void CaptureVisibleElements(string sFileName, IWebDriver _driver)
        {
            string DebugMode = GenericFunctions.goAndGet("DebugMode");
            if (DebugMode.ToUpper() == "FALSE")
            {
                Console.WriteLine("CaptureVisibleElements NOT run:  DebugMode = " + DebugMode);
                return;
            }
            String sResults = Globals.Run.PsResultsFolder + sFileName;
            Console.WriteLine("CaptureVisibleElements Starting:  sFileName = " + sFileName);
            if (sFileName.Substring(sFileName.Length - 4, 4).ToUpper() != ".CSV") { sFileName = sFileName + ".csv"; }
            DataSet dsPageCheck = ListVisibleElementProperties(_driver);
            //string sResults = Globals.Run.psResultsFolder + sFileName;
            //Console.WriteLine("     dsPageCheck = " + dsPageCheck.Tables.Count);
            Console.WriteLine("     sResults = " + sResults);
            GenericFunctions.ConvertToCSV(dsPageCheck, "ElementProperties", sResults);
            Console.WriteLine("CaptureVisibleElements Finished:  ResultsFileName = " + sResults);
        }


        public static DataSet ListVisibleElementProperties(IWebDriver _driver)
        {
            string DebugMode = GenericFunctions.goAndGet("DebugMode");
            if (DebugMode.ToUpper() == "FALSE")
            {
                Console.WriteLine("ListVisibleElementProperties NOT run:  DebugMode = " + DebugMode);
                return null;
            }
            //Console.WriteLine("ListVisibleElementProperties Starting Function");
            DataSet MyDataSet = new DataSet("ListVisibleElementProperties");
            DataTable dtElementProperties = new DataTable("ElementProperties");
            dtElementProperties.Columns.Add("ID");
            dtElementProperties.Columns.Add("SELECTED");
            dtElementProperties.Columns.Add("DISPLAYED");
            dtElementProperties.Columns.Add("ENABLED");
            dtElementProperties.Columns.Add("TYPE");
            dtElementProperties.Columns.Add("NAME");
            dtElementProperties.Columns.Add("CLASS");
            dtElementProperties.Columns.Add("MAXLENGTH");
            dtElementProperties.Columns.Add("TEXT");

            string sXpath = ".//*[(self::input or self::span or self::a)]";
            long iElements = _driver.FindElements(By.XPath(sXpath)).LongCount();
            //Console.WriteLine("     iElements=" + iElements.ToString());
            if (iElements > 0)
            {
                for (int i = 0; i <= (iElements - 1); i++) // Go through all the error results and report each error
                {
                    string sID = _driver.FindElements(By.XPath(sXpath)).ElementAt(i).GetAttribute("id");
                    string sSelected = _driver.FindElements(By.XPath(sXpath)).ElementAt(i).Selected.ToString();
                    string sVisible = _driver.FindElements(By.XPath(sXpath)).ElementAt(i).Displayed.ToString();
                    string sEnabled = _driver.FindElements(By.XPath(sXpath)).ElementAt(i).Enabled.ToString();
                    string sType = _driver.FindElements(By.XPath(sXpath)).ElementAt(i).GetAttribute("type");
                    string sName = _driver.FindElements(By.XPath(sXpath)).ElementAt(i).GetAttribute("name");
                    string sClass = _driver.FindElements(By.XPath(sXpath)).ElementAt(i).GetAttribute("class");
                    string sLength = _driver.FindElements(By.XPath(sXpath)).ElementAt(i).GetAttribute("maxlength");
                    string sText = _driver.FindElements(By.XPath(sXpath)).ElementAt(i).Text.ToString();
                    sText = sText.Replace("\n", ";").Replace("\r", ";");
                    if (sType != "hidden" && (sID + sType + sName + sClass + sLength + sText) != "")
                    {
                        dtElementProperties.Rows.Add(sID, sSelected, sVisible, sEnabled, sType, sName, sClass, sLength, sText);
                        //Console.WriteLine("ID: {0}", sID);
                        //Console.WriteLine("Selected: {0}", sSelected);
                        //Console.WriteLine("Enabled: {0}", sEnabled);
                        //Console.WriteLine("Displayed: {0}", sVisible);
                        //Console.WriteLine("Type: {0}", sType);
                        //Console.WriteLine("Name: {0}", sName);
                        //Console.WriteLine("Class: {0}", sClass);
                        //Console.WriteLine("maxlength: {0}", sLength);
                        //Console.WriteLine("Text: {0}", sText);
                    }
                }
                //Return the list of attributes
                Console.WriteLine("ListVisibleElementProperties Finished Function:  True (" + iElements.ToString() + " elements found)");
                MyDataSet.Tables.Add(dtElementProperties);
                return MyDataSet;
            }
            //No elements found return null
            Console.WriteLine("ListElementCollectionsAttributeToStringList Function:  False (No elements found)");
            return MyDataSet;
        }




        public static List<string> ListElementCollectionsAttributeToStringList(string sXpath, string sAttribute, IWebDriver _driver)
        {
            string DebugMode = GenericFunctions.goAndGet("DebugMode");
            if (DebugMode.ToUpper() == "FALSE")
            {
                Console.WriteLine("ListElementCollectionsAttributeToStringList NOT run:  DebugMode = " + DebugMode);
                return null;
            }
            Console.WriteLine("ListElementCollectionsAttributeToStringList Starting Function:  sXpath=" + sXpath + ", " + sAttribute);
            List<string> slElementAttibuteValues = new List<string>();
            int iCount = _driver.FindElements(By.XPath(sXpath)).Count();
            if (iCount > 0)
            {
                for (int i = 0; i <= (iCount - 1); i++) // Go through all the error results and report each error
                {
                    slElementAttibuteValues.Add(_driver.FindElements(By.XPath(sXpath)).ElementAt(i).GetAttribute(sAttribute).ToString());
                }
                //Return the list of attributes
                Console.WriteLine("ListElementCollectionsAttributeToStringList Finished Function:  slElementAttibuteValues=" + string.Join(";", slElementAttibuteValues));
                return slElementAttibuteValues;
            }
            //No elements found return null
            slElementAttibuteValues = null;
            Console.WriteLine("ListElementCollectionsAttributeToStringList Finished Function:  slElementAttibuteValues=null");
            return slElementAttibuteValues;
        }


        public static List<string> ListElementCollectionsTextValuesToStringList(string sXpath, IWebDriver _driver)
        {
            string DebugMode = GenericFunctions.goAndGet("DebugMode");
            if (DebugMode.ToUpper() == "FALSE")
            {
                Console.WriteLine("ListElementCollectionsTextValuesToStringList NOT run:  DebugMode = " + DebugMode);
                return null;
            }
            Console.WriteLine("ListElementCollectionsTextValuesToStringList Starting Function:  sXpath=" + sXpath);
            List<string> slElementsTextValues = new List<string>();
            int iCount = _driver.FindElements(By.XPath(sXpath)).Count();
            if (iCount > 0)
            {
                for (int i = 0; i <= (iCount - 1); i++) // Go through all the elements and add each ones text value to the list
                {
                    slElementsTextValues.Add(_driver.FindElements(By.XPath(sXpath)).ElementAt(i).Text.ToString());
                }
                //Return the list of elements
                Console.WriteLine("ListElementCollectionsTextValuesToStringList Finished Function:  slElementsTextValues=" + string.Join(";", slElementsTextValues));
                return slElementsTextValues;
            }
            //No elements found return null
            slElementsTextValues = null;
            Console.WriteLine("ListElementCollectionsTextValuesToStringList Finished Function:  slElementsTextValues=null");
            return slElementsTextValues;
        }

    }
}
