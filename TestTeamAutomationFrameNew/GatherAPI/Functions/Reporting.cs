using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using TestTeamAutomationFrameNew;
using TestTeamAutomationFrameNew.ITC_SMK;
using System.Drawing.Imaging;
using System.Windows;
using System.Data.SqlClient;


namespace TestTeamAutomationFrameNew.Functions
{
 
    public class Reporting
    {

        public static IWebDriver driver { get; set; }
        public static AppiumDriver Appdriver { get; set; }
        public static string PerfGraph = null;


        public static void ReportHeader(string TestName)
        {

            // Ensure the report file is unique
            EnsureHeaderTitleIsUnique();

            //var mobile = GenericFunctions.goAndGet("MOBILE");
            var Envname = GenericFunctions.goAndGet("Envname");
            string showBrowser;
            var reportmsg = new System.IO.StreamWriter(@"C:\SVN\Automation\TestTeamAutomationFrameNew\bin\Debug\" + _startUp.resultsFileName, true);
            string Header = Windows.ReadToString(@"C:\SVN\Automation\TestTeamAutomationFrameNew\Reporting\design\header.txt");


            Header = Header.Replace("<REPORTTESTNAME>", TestName);
            Header = Header.Replace("<style></style>", "<style>" + Windows.ReadToString(@"C:\SVN\Automation\TestTeamAutomationFrameNew\Reporting\css\Layout.css") + "</style>");
            Header = Header.Replace("<script></script>", "<script>" + Windows.ReadToString(@"C:\SVN\Automation\TestTeamAutomationFrameNew\Reporting\Scripts\Scripts.js") + "</script>");

            try
            {
                Header = Header.Replace("<REPORTWIDTH>", GenericFunctions.goAndGet("REPORTWIDTH"));

            }

            catch (Exception)
            {
                Header = Header.Replace("<REPORTWIDTH>", "1000");

            }

            
            reportmsg.WriteLine(Header);
            reportmsg.WriteLine("<br><h2>Test Report Summary</h2><br>");
            reportmsg.WriteLine("<div id=\"Results\"><center><table class=\"matrix\">");
            reportmsg.WriteLine("<tr><td><strong>Environment:</strong></td><td>" + Envname + "</td></tr>");


            //C:\Reports
            //@"c:\";
            var driver = "*firefox";


            switch (driver)
            {
                case "*iexplore":
                    showBrowser = "Internet Explorer";
                    break;

                    case "*firefox":
                    showBrowser = "Firefox";
                    break;

                default:
                    showBrowser = "Chrome";
                    break;

            }
           
            reportmsg.WriteLine("<tr><td><strong>Browser:</strong></td><td>" + showBrowser + "</td></tr>");
            reportmsg.WriteLine("<tr><td><strong>Username:</strong></td><td>" + Windows.GetMyUsername(true) + "</td></tr>");

            try
            {
                reportmsg.WriteLine("<tr><td><strong>Test Package:</strong></td><td>" + GenericFunctions.goAndGet("TESTPACKAGE") + "</td></tr>");

            }

            catch (Exception)
            {

                Console.WriteLine("<TESTPACKAGE> not set in the paramters file. If you want the TP number to appear at the bottom, please add <TESTPACKAGE>TP00x</TESTPACKAGE> into the parameters file.");

            }

            reportmsg.WriteLine("<tr><td><strong>Start Time:</strong></td><td>" + DateTime.Now.ToString("dd-MM-yyyy HH:mm.ss") + "</td></tr>");
            reportmsg.WriteLine("<tr><td><strong>End Time:</strong></td><td></td></tr>");
            reportmsg.WriteLine("</table></center></div>");
            reportmsg.WriteLine("<br>");
            reportmsg.Close();
        }

        public static void EnsureHeaderTitleIsUnique()
        {
            int counter = 1;


            // While a report with the name exists
            while (System.IO.File.Exists(ITC_SMK._startUp.resultsFileName))
            {
                if (counter == 1)
                {
                    ITC_SMK._startUp.resultsFileName = ITC_SMK._startUp.resultsFileName.Replace(".html", "_" + counter + ".html");

                }
                else
                {

                    ITC_SMK._startUp.resultsFileName = ITC_SMK._startUp.resultsFileName.Replace("_" + (counter - 1) + ".html", "_" + counter + ".html");

                }
                counter++;

                if (counter == 100)
                {
                    break;

                }

            }

        }

        public static void ReportFooter()
        {
            var dt = DateTime.Now;
            var now = dt.ToString("ddMMyyyy");
            double totalPasses = Math.Round(System.Convert.ToDouble(_startUp.TestPasses));
            double totalFails = Math.Round(System.Convert.ToDouble(_startUp.TestFails));
            double totalWarns = Math.Round(System.Convert.ToDouble(_startUp.Warnings));
            double total = Math.Round(totalPasses + totalFails);
            double passPercent = Math.Round((totalPasses / total) * 100);
            double failPercent = Math.Round((totalFails / total) * 100);
            string TPNO = null;
            string Footer = null;

            // Try to get the Test Package number from the config file. If it's not there, then don't include it on the report.
            try { TPNO = GenericFunctions.goAndGet("TESTPACKAGE"); }

            catch (Exception)
            {

                Console.WriteLine("<TESTPACKAGE> not set in the parameters file. If you want the TP number to appear at the bottom, please add <TESTPACKAGE>TP00x</TESTPACKAGE> into the parameters file.");

            }

            var reportmsg = new StreamWriter(_startUp.resultsFileName, true);
            reportmsg.WriteLine("<br><br><p align=\"right\"><a href=#top><img src=\"data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABkAAAAZCAYAAADE6YVjAAAABGdBTUEAALGPC/xhBQAAAAlwSFlzAAAOwgAADsIBFShKgAAAABh0RVh0U29mdHdhcmUAcGFpbnQubmV0IDQuMC45bDN+TgAAAqJJREFUSEutlktoE0EcxiebTTZNTARFRc2xooiC4E0ExYuevHkQRRR8oHiQmoNe6qUtIqgX8d2b4M7sbpPWWLzFIoIHD1VasjMbia9YoVoRRao2Wf9JZmmymc2mwQ8+WPY/8/12duexyE+94yyhaPREkLAbMqE6UllWwjQTxPReGNPUeoMledOla+2YuSlI6LBE2DdELNvLAWz9ljEdjWv5HbxrZ4roZl+AsDlRqJcDmC2ENXZ7+1gpymO8JWPrviikUwNsImbMrOZxrQphelPUcamWCH26MZOP89hFRYl5HGFWEXXqxgpmV3l0XZvx1AoJs5KosduHXszYy0YKwlqj4bWVI5q1OBlC2BwSNXQ7NfnFXqhU7MeffnYEgtGoHAEfW6PTokaNdgCOsgCK+YACxCqvGrV6UTKd34pI+2+RmpxtAjjKlgBktAeFVHoEhVXznKjo+LwLkP8+z6/qqo2oDQh2ikEUUvOXRcWq3QD9ww87MfLGHpia5XfqajeisEqHEXpoXRMV+zwATl0EiopAmD6AmcUG3AU/gGM36JEAJGN2BymEnmq8ueXJW/tP2R/geNAFOgjrqLEOkP7qVp6UYKo1Fs68/GxXYCSGD8Dx0HQddMWcs4Nacy2smQdq6wRW+/PGQtX7n5XsRNof4HjvxMcWAOxh82uMQn2zjBms7TTu1nDYXa8Bqto3zhQYzWtRw24Neb/iemEDR9QVSxd3V085UYdu3KPTszy6WUA//T+2e1jld3flcjKPbVUPgGCb/ivq3IlDhN5CuWKEx3lLMQp74G/EFIV4Gd7CV0UvHEWXbInH+Kv6KwRPdVIm7JUo1HFQY+9llV5cmXm3jnftSoHlpLgNDrbDsmal4Jv1hwm7EFLZsQgxd8LDKLydhxD6B1tQKtjZs1AdAAAAAElFTkSuQmCC\" border=0 title=\"Back to Top\"></a></p><br><br>");

            Footer = Windows.ReadToString(_startUp.RunEnv + "../../Reporting/design/footer.txt");
            Footer = Footer.Replace("<YEAR>", dt.ToString("yyyy"));
            Footer = Footer.Replace("<TPNO>", (TPNO == null ? "" : "Test Package: " + TPNO));
            reportmsg.WriteLine(Footer);
            reportmsg.Close();
            _startUp.Contents = _startUp.Contents + "<td>Pass: <font color=\"#008000\">" + _startUp.TestPasses + " </font>(<font color=\"#008000\">" + passPercent + "%</font>).</td><td>Fail: <font color=\"#FF0000\">" + _startUp.TestFails + " </font>(<font color=\"#FF0000\">" + failPercent + "%</font>).</td></tr></table></center></div>";
            GenericFunctions.GenerateGraph();
            CopyTableIntoResults();
            Reporting.PerfGraph = null;
            updateEndTime();
            CleanUp();
            _startUp.TestID = 1;
        }

        public static void TestName(string testname, bool UseTestID = true, bool IncludePageBreak = false)
        {

            var reportmsg = new System.IO.StreamWriter(@"C:\SVN\Automation\TestTeamAutomationFrameNew\bin\Debug" + _startUp.resultsFileName, true);

            if (_startUp.TestID != 1)
            {
                reportmsg.WriteLine("<br><br><p align=\"right\"><a href=#top><img src=\"data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABkAAAAZCAYAAADE6YVjAAAABGdBTUEAALGPC/xhBQAAAAlwSFlzAAAOwgAADsIBFShKgAAAABh0RVh0U29mdHdhcmUAcGFpbnQubmV0IDQuMC45bDN+TgAAAqJJREFUSEutlktoE0EcxiebTTZNTARFRc2xooiC4E0ExYuevHkQRRR8oHiQmoNe6qUtIqgX8d2b4M7sbpPWWLzFIoIHD1VasjMbia9YoVoRRao2Wf9JZmmymc2mwQ8+WPY/8/12duexyE+94yyhaPREkLAbMqE6UllWwjQTxPReGNPUeoMledOla+2YuSlI6LBE2DdELNvLAWz9ljEdjWv5HbxrZ4roZl+AsDlRqJcDmC2ENXZ7+1gpymO8JWPrviikUwNsImbMrOZxrQphelPUcamWCH26MZOP89hFRYl5HGFWEXXqxgpmV3l0XZvx1AoJs5KosduHXszYy0YKwlqj4bWVI5q1OBlC2BwSNXQ7NfnFXqhU7MeffnYEgtGoHAEfW6PTokaNdgCOsgCK+YACxCqvGrV6UTKd34pI+2+RmpxtAjjKlgBktAeFVHoEhVXznKjo+LwLkP8+z6/qqo2oDQh2ikEUUvOXRcWq3QD9ww87MfLGHpia5XfqajeisEqHEXpoXRMV+zwATl0EiopAmD6AmcUG3AU/gGM36JEAJGN2BymEnmq8ueXJW/tP2R/geNAFOgjrqLEOkP7qVp6UYKo1Fs68/GxXYCSGD8Dx0HQddMWcs4Nacy2smQdq6wRW+/PGQtX7n5XsRNof4HjvxMcWAOxh82uMQn2zjBms7TTu1nDYXa8Bqto3zhQYzWtRw24Neb/iemEDR9QVSxd3V085UYdu3KPTszy6WUA//T+2e1jld3flcjKPbVUPgGCb/ivq3IlDhN5CuWKEx3lLMQp74G/EFIV4Gd7CV0UvHEWXbInH+Kv6KwRPdVIm7JUo1HFQY+9llV5cmXm3jnftSoHlpLgNDrbDsmal4Jv1hwm7EFLZsQgxd8LDKLydhxD6B1tQKtjZs1AdAAAAAElFTkSuQmCC\" border=0 title=\"Back to Top\"></a></p><br><br>");

            }
            reportmsg.WriteLine("<h3><a id=T" + _startUp.TestID + ">" + (IncludePageBreak == true ? "*PageBreak*" : "") + (UseTestID ? _startUp.TestID + ". " : "") + testname + "</a></h3>");
            reportmsg.Close();
            WriteTestToFile(_startUp.TestID, testname);
            _startUp.TestID++;

        }

        public static void TestSubHeading(string testname, bool IncludePageBreak = false)
        {

            var reportmsg = new System.IO.StreamWriter(_startUp.resultsFileName, true);

            reportmsg.WriteLine("<br><br><p align=\"right\"><a href=#top><img src=\"data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABkAAAAZCAYAAADE6YVjAAAABGdBTUEAALGPC/xhBQAAAAlwSFlzAAAOwgAADsIBFShKgAAAABh0RVh0U29mdHdhcmUAcGFpbnQubmV0IDQuMC45bDN+TgAAAqJJREFUSEutlktoE0EcxiebTTZNTARFRc2xooiC4E0ExYuevHkQRRR8oHiQmoNe6qUtIqgX8d2b4M7sbpPWWLzFIoIHD1VasjMbia9YoVoRRao2Wf9JZmmymc2mwQ8+WPY/8/12duexyE+94yyhaPREkLAbMqE6UllWwjQTxPReGNPUeoMledOla+2YuSlI6LBE2DdELNvLAWz9ljEdjWv5HbxrZ4roZl+AsDlRqJcDmC2ENXZ7+1gpymO8JWPrviikUwNsImbMrOZxrQphelPUcamWCH26MZOP89hFRYl5HGFWEXXqxgpmV3l0XZvx1AoJs5KosduHXszYy0YKwlqj4bWVI5q1OBlC2BwSNXQ7NfnFXqhU7MeffnYEgtGoHAEfW6PTokaNdgCOsgCK+YACxCqvGrV6UTKd34pI+2+RmpxtAjjKlgBktAeFVHoEhVXznKjo+LwLkP8+z6/qqo2oDQh2ikEUUvOXRcWq3QD9ww87MfLGHpia5XfqajeisEqHEXpoXRMV+zwATl0EiopAmD6AmcUG3AU/gGM36JEAJGN2BymEnmq8ueXJW/tP2R/geNAFOgjrqLEOkP7qVp6UYKo1Fs68/GxXYCSGD8Dx0HQddMWcs4Nacy2smQdq6wRW+/PGQtX7n5XsRNof4HjvxMcWAOxh82uMQn2zjBms7TTu1nDYXa8Bqto3zhQYzWtRw24Neb/iemEDR9QVSxd3V085UYdu3KPTszy6WUA//T+2e1jld3flcjKPbVUPgGCb/ivq3IlDhN5CuWKEx3lLMQp74G/EFIV4Gd7CV0UvHEWXbInH+Kv6KwRPdVIm7JUo1HFQY+9llV5cmXm3jnftSoHlpLgNDrbDsmal4Jv1hwm7EFLZsQgxd8LDKLydhxD6B1tQKtjZs1AdAAAAAElFTkSuQmCC\" border=0 title=\"Back to Top\"></a></p><br><br>");
            reportmsg.WriteLine("<h4>" + (IncludePageBreak == true ? "*PageBreak*" : "") + testname + "</h4>");
            reportmsg.Close();

        }

        // Add notes the the current test
        public static void TestNotes(string Notes)
        {

            System.IO.StreamWriter reportmsg = new System.IO.StreamWriter(ITC_SMK._startUp.resultsFileName, true);
            reportmsg.WriteLine("<h4>Notes:</h4><font face=\"verdana\" size=\"2\" color=\"#000000\">" + Notes + "</font>");
            reportmsg.Close();

        }

        public static void WriteTestToFile(int counter, string Test)
        {

            double TotalPasses = Math.Round(System.Convert.ToDouble(_startUp.TestPasses));
            double TotalFails = Math.Round(System.Convert.ToDouble(_startUp.TestFails));
            double Total = Math.Round(TotalPasses + TotalFails);
            double PassPercent = Math.Round((TotalPasses / Total) * 100);
            double FailPercent = Math.Round((TotalFails / Total) * 100);


            // Writing text to the File.
            if (_startUp.TestID == 1)
            {
                _startUp.Contents = "<h2>Contents</h2>";
                _startUp.Contents = _startUp.Contents + "<div id=\"Results\"><center><table class=\"matrix\"><tr><th valign=\"TOP\">ID</th><th valign=\"TOP\">Description</th><th colspan=\"2\">Pass/Fail</th></tr>";
            }

            if (_startUp.TestID != 1)
            {
                _startUp.Contents = _startUp.Contents + "<td>Pass: <font color=\"#008000\">" + _startUp.TestPasses + "</font> (<font color=\"#008000\">" + PassPercent + "%</font>).</td><td>Fail: <font color=\"#FF0000\">" + _startUp.TestFails + "</font> (<font color=\"#FF0000\">" + FailPercent + "%</font>).</td></tr>";
                _startUp.TestPasses = 0;
                _startUp.TestFails = 0;
            }
            _startUp.Contents = _startUp.Contents + "<tr><td valign=\"TOP\"><a href=#T" + _startUp.TestID + " title=\"Click to go to to the test\">" + _startUp.TestID + "</a></td><td valign=\"TOP\"><a href=#T" + _startUp.TestID + " title=\"Click to go to to the test\">" + Test + "</a></td>";
            TotalPasses = 0;
            TotalFails = 0;
            Total = 0;
            PassPercent = 0;
            FailPercent = 0;

        }
        
        public static void CleanUp()
        {
            string myDir = _startUp.resultsFileName;
            //string resultsfolder = myDir.Remove(44);
            string sourcePath = AppDomain.CurrentDomain.BaseDirectory.ToString();
            //string FileName = "\"Reporting/img/backtotop.png\"";
            //_startUp.RunEnv + "footer.txt"
            if (System.IO.Directory.Exists(sourcePath))
            {
                string[] FilePaths = Directory.GetFiles(sourcePath, "*.png");

                //    // Copy the Files and overwrite destination Files if they already exist.
                foreach (var File in FilePaths)
                {
                    string name = Path.GetFileName(File);
                    //         //string dest = Path.Combine(resultsfolder, FileName);
                    //    System.IO.File.Copy(File, dest, true);
                }
            }
        }

        public static void ReportScreenshot(string Filename, IWebDriver driver)
        {
            string activeDir = _startUp.resultsFileName;
            Filename = Filename.ToLower() + ".jpg";
            string newFilename = activeDir + Filename;
            string ScreenshotWidth = null;


            try { ScreenshotWidth = GenericFunctions.goAndGet("IMGWIDTH"); }
            catch (Exception) { Console.WriteLine("Couldn't grab IMGWIDTH from config file"); }


             ((ITakesScreenshot)driver).GetScreenshot()
                .SaveAsFile(Filename, OpenQA.Selenium.ScreenshotImageFormat.Jpeg);
             TestStep("<br><a href=" + Filename + " target=new><img src=\"" + Filename + "\" width=" + (ScreenshotWidth != null ? ScreenshotWidth : "500") +" border=0 Title=\"Click to view larger image\"></a><br>");

         }


        public static void ReportAppScreenshot(string Filename, AppiumDriver driver)
        {
            string activeDir = _startUp.resultsFileName;
            DateTime dt = DateTime.Now;
            string now = dt.ToString("ddMMyyyy-hhmmss");
            Filename = Filename.ToUpper() + "-" + now + ".jpeg";

            ((ITakesScreenshot)driver).GetScreenshot().SaveAsFile(Filename, OpenQA.Selenium.ScreenshotImageFormat.Jpeg);
            TestStep("<br><a href=" + activeDir + Filename + " target=new><img src=\"" + activeDir + Filename + "\" width=\"250\" Title=\"Click to view larger image\" border=0></a>");
        }

        public static void TestStep(string Message)
        {

            string Date = DateTime.Now.ToString("dd-MM-yyyy HH:mm.ss");


            System.IO.StreamWriter reportmsg = new System.IO.StreamWriter(_startUp.resultsFileName, true);
            reportmsg.WriteLine("<font face=\"verdana\" size=\"2\" color=\"#000000\">[" + Date + "] " + Message + "</font><br>");
            reportmsg.Close();            

        }

        // Check the Event log solution defined field.
        public static bool ReportDataFromTheDatabase(string Query, string Database, string Highlight = "", bool ShowQuery = false)
        {
            Console.WriteLine(Query);
            DataTable MyDataTable   = new DataTable("MyTable");
            string Table            = null;
            bool HeaderRow          = true;
            bool Data               = false;


            if (Query.Substring(0, 6).ToLower().Contains("select"))
            {

                // If we need to show the query on the report, show it.
                if (ShowQuery)
                {
                    Reporting.InformationBox(Query);

                }

                // Try connecting to the database
                try
                {

                    // Setup the connection and fill the data table with the data
                    SqlConnection Connection = new SqlConnection(GenericFunctions.goAndGet("DBCONNECTION").Replace("DBNAME", Database));
                    SqlCommand cmd = new SqlCommand(Query, Connection);
                    Connection.Open();
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(MyDataTable);
                    Connection.Close();

                    // For each row within the data table
                    foreach (DataRow row in MyDataTable.Rows)
                    {

                        // If this is the header row, add in the header formatting
                        if (HeaderRow)
                        {

                            Table = "<table class=\"matrix\"><tr>";

                            // For each column, add it into the table
                            foreach (DataColumn column in MyDataTable.Columns)
                            {

                                Table = Table + "<th>" + column.ColumnName + "</th>";

                            }

                            Table = Table + "</tr>";

                        }

                        // Switch the header row off
                        HeaderRow = false;

                        // Create a new row
                        Table = Table + "<tr>";

                        // For each column, add it into the report
                        foreach (DataColumn column in MyDataTable.Columns)
                        {

                            Table = Table + "<td>" + (Highlight.Contains(column.ColumnName) ? "<mark>" : "") + (row[column].ToString().Length != 0 ? row[column] : "NULL") + (Highlight.Contains(column.ColumnName) ? "</mark>" : "") + "</td>";
                            Data = true;

                        }

                        // Finish the table row
                        Table = Table + "</tr>";

                    }

                    // Finish off the table
                    Table = Table + "</table>";

                    // Destroy the datatable
                    da.Dispose();

                }

                // Catch any exceptions here
                catch (Exception e)
                {
                    Console.WriteLine("Something went wrong...\n" + e);
                }

                // Add the data into the report
                Reporting.TestStep("<br>" + (Data ? Table : "No data to display."));
            }

            return Data;

        }

        // Place the given JSON on the report
        public static void JSON(string Request, string Highlight, string XML = null)
        {

            Reporting.TestStep((XML == null ? "Get the JSON response:<br>" + HighlightJSON(Request, Highlight) : "Get the JSON response:<br>" + HighlightJSON(Request, "").Replace("<SolutionExport", "<xmp><SolutionExport").Replace("</SolutionExport>", "</SolutionExport></xmp>")) + "<br>");

        }


        // Highlight parts of the JSON string
        public static string HighlightJSON(string Request, string Highlight)
        {

            try
            {



                if (Highlight.Contains("Result"))
                {

                    Request = Request.Replace("\"Result\":\"OK\"", "<mark>\"Result\":\"OK\"</mark>");
                    Request = Request.Replace("\"Result\":\"FAIL\"", "<mark>\"Result\":\"FAIL\"</mark>");

                }

                if (Highlight.Contains("ResultCode"))
                {

                    Request = Request.Replace("\"ResultCode\":\"FAIL_NOT_FOUND\"", "<mark>\"ResultCode\":\"FAIL_NOT_FOUND\"</mark>");
                    Request = Request.Replace("\"ResultCode\":\"FAIL_INVALID_PARAMETERS\"", "<mark>\"ResultCode\":\"FAIL_INVALID_PARAMETERS\"</mark>");
                    Request = Request.Replace("\"ResultCode\":\"FAIL_AT_LEAST_ONE_PARAMETER_MUST_BE_DEFINED\"", "<mark>\"ResultCode\":\"FAIL_AT_LEAST_ONE_PARAMETER_MUST_BE_DEFINED\"</mark>");

                }

                if (Highlight.Contains("UserDefinedField"))
                {

                    Request = Request.Replace("\"UserDefinedField\"", "<mark>\"UserDefinedField\"</mark>");
                    Request = Request.Replace("\"UserDefinedField\":{}", "<mark>\"UserDefinedField\":{}</mark>");

                }

                if (Highlight.Contains("DeviceXml"))
                {

                    Request = Request.Replace("<device id", "<xmp><device id");
                    Request = Request.Replace("</device", "</device></xmp>");

                }

                if (Highlight.Contains("mDNATag"))
                {

                    Request = Request.Replace("\"mDNATag\"", "<mark>\"mDNATag\"</mark>");

                }

                if (Highlight.Contains("PublicId"))
                {

                    Request = Request.Replace("\"PublicId\"", "<mark>\"PublicId\"</mark>");

                }

                if (Highlight.Contains("State"))
                {

                    Request = Request.Replace("\"State\"", "<mark>\"State\"</mark>");

                }

                if (Highlight.Contains("Revision"))
                {

                    Request = Request.Replace("\"Revision\"", "<mark>\"Revision\"</mark>");

                }

                if (Highlight.Contains("SolutionName"))
                {

                    Request = Request.Replace("\"SolutionName\"", "<mark>\"SolutionName\"</mark>");

                }

                if (Highlight.Contains("ProvisioningUrl"))
                {

                    Request = Request.Replace("\"ProvisioningUrl\"", "<mark>\"ProvisioningUrl\"</mark>");

                }

                if (Highlight.Contains("ActivationCode"))
                {

                    Request = Request.Replace("\"ActivationCode\"", "<mark>\"ActivationCode\"</mark>");

                }

                if (Highlight.Contains("ActivationCodeNumeric"))
                {

                    Request = Request.Replace("\"ActivationCodeNumeric\"", "<mark>\"ActivationCodeNumeric\"</mark>");

                }

                if (Highlight.Contains("GroupId"))
                {

                    Request = Request.Replace("\"GroupId\"", "<mark>\"GroupId\"</mark>");

                }

                if (Highlight.Contains("InstanceId"))
                {

                    Request = Request.Replace("\"InstanceId\"", "<mark>\"InstanceId\"</mark>");

                }

                if (Highlight.Contains("Language"))
                {

                    Request = Request.Replace("\"Language\"", "<mark>\"Language\"</mark>");

                }

                if (Highlight.Contains("UserDefinedField"))
                {

                    Request = Request.Replace("\"UserDefinedField\"", "<mark>\"UserDefinedField\"</mark>");

                }

                Request = Request.Replace("{", "{<br>&nbsp;&nbsp;&nbsp;&nbsp;");
                Request = Request.Replace(",", ",<br>&nbsp;&nbsp;&nbsp;&nbsp;");
                Request = Request.Replace("}", "<br>}");

            }
            catch (Exception){


            }

            return Request;

        }


        public static void CopyTableIntoResults()
        {

            var resultsFile = Windows.ReadToString(_startUp.resultsFileName);
            var path = _startUp.RunEnv + "ResultsOverTime.txt";
            var FileContents = Windows.ReadToString(path);
            var lineGraph = "<h2>Results Over Time</h2><script type=\"text/javascript\">google.load(\"visualization\", \"1\", {packages:[\"corechart\"]});google.setOnLoadCallback(drawChart);function drawChart() {var data = google.visualization.arrayToDataTable([['Date', 'Pass', 'Fails', 'Warnings']," + FileContents +
                               "]);var options = {title: 'Software Results Over time',height: 500,width: 900,colors: ['Green','Red','Orange'],hAxis: {title: 'Date',  titleTextStyle: {color: '#333'}},vAxis: {minValue: 0}};var chart = new google.visualization.AreaChart(document.getElementById('chart_div2'));chart.draw(data, options);}</script> <div id=\"chart_div2\" style=\"width: 900px; height: 500px;\"></div>";
            resultsFile = resultsFile.Replace("<GRAPH>", "<h2>Results</h2><br><center><script type=\"text/javascript\" src=\"https://www.google.com/jsapi\"></script>    <script type=\"text/javascript\">      google.load(\"visualization\", \"1\", {packages:[\"corechart\"]});      google.setOnLoadCallback(drawChart);      function drawChart() {        var data = google.visualization.arrayToDataTable([          ['Task', 'Hours per Day'],        " +
                                                         "  ['Passes',     " + _startUp.Passes + "],         " +
                                                         " ['Failures',      " + _startUp.Fails + "],          ['Warnings',  " + _startUp.Warnings + "]]);       " +
                                                         " var options = {          title: 'Smoke Test Results', colors: ['Green','Red','Orange']        };      " +
                                                         "  var chart = new google.visualization.PieChart(document.getElementById('chart_div'));       " +
                                                         " chart.draw(data, options);      " +
                                                         "}    </script>  <div id=\"chart_div\" style=\"width: 900px; height: 500px;\"></div></center><br><br>");
            resultsFile = resultsFile.Replace("<LINEGRAPH>", lineGraph);
            resultsFile = resultsFile.Replace("<RESULTSTABLE>", _startUp.Contents).Replace("NaN%", "0%");

            if (PerfGraph != null)
            {

                resultsFile = resultsFile.Replace("<PERFHEADER>", "<li><a href=\"#tab3\">Performance</a></li>");
                resultsFile = resultsFile.Replace("<PERF>", "<div id=\"tab3\" class=\"tab\"><div id=\"Results\"><h2>Performance Stats</h2>" + PerfGraph + "</div></div>");

            }

            Windows.ClearFile(_startUp.resultsFileName);
            Windows.WriteToFile(resultsFile, _startUp.resultsFileName);

        }

        public static void updateEndTime()
        {
            
            var resultsFile = Windows.ReadToString(_startUp.resultsFileName);
            

            resultsFile = resultsFile.Replace("<tr><td><strong>End Time:</strong></td><td></td></tr>", "<tr><td><strong>End Time:</strong></td><td>" + DateTime.Now.ToString("dd-MM-yyyy HH:mm.ss") + "</td></tr>");
            Windows.ClearFile(_startUp.resultsFileName);
            Windows.WriteToFile(resultsFile, _startUp.resultsFileName);

        }

        public static void updateSmokeTestTitle(string smoketest)
        {
            var resultsFile = Windows.ReadToString(_startUp.resultsFileName);
            resultsFile = resultsFile.Replace("Engage", smoketest);
            Windows.ClearFile(_startUp.resultsFileName);
            Windows.WriteToFile(resultsFile, _startUp.resultsFileName);
        }

        public static void UpdatePerfGraph(int TestID, string Action, int Time)
        {

            bool GenPerf = false;


            // Try getting the option to gather performance results from the config file
            try { GenPerf = Convert.ToBoolean(GenericFunctions.goAndGet("PERFSTATS")); }

            catch (Exception) { }

            if (GenPerf)
            {

                string UniqueID = ("TEST" + TestID + "+" + Action).Replace(".mdna.excointouch.com/json/syncreply/", "").Replace(".services.excointouch.com/api/json/syncreply/", "")
                                    .Replace(" ", "").Replace("\"", "").Replace(".", "").Replace("http://", "").Replace("https://", "").Replace("/", "").ToUpper();


                // Trim the unique ID down to 50 characters if it is too long
                if (UniqueID.Length > 50)
                {

                    UniqueID = UniqueID.Substring(0, 50);

                }

                // If the graph hasn't been set up yet, set up the header row
                if (PerfGraph == null)
                {

                    PerfGraph = "<center><table class=\"matrix\"><tr><th>Test</th><th>Action</th><th>Time (ms)</th></tr></table></center>";

                }

                // Add a line to the table
                PerfGraph = PerfGraph.Replace("</table>", "<tr><td>" + TestID + "</td><td>" + Action + "</td><td><font color=\"#" + (Time > CalcAverageTime(UniqueID) ? "DB2929" : "526F35") + "\"><center>" + Time + "</center></font></td></tr></table>");

                // Write the value to the Database
                InsertIntoPerformanceTable(UniqueID, Time);
            }

        }

        // Looks at the performance table and calculates the average
        public static int CalcAverageTime(string UniqueID)
        {

            string Query    = "SELECT top %% AVG(Time) AS Average FROM Performance where UniqueId = '" + UniqueID + "' and Environment = '" + _startUp.RunEnv + "'";
            Random rnd      = new Random();
            int Average     = 0;
            int Limit       = 0;            


            //// Try getting the average limit from the config file. If it's not there, limit it to 30 results
            //try { Limit = Convert.ToInt32(GenericFunctions.goAndGet("AVERAGELIMIT")); }
            //catch (Exception) { Limit = 30; }

            //// Set up the database connection
            //using (SqlConnection conn = new SqlConnection(GenericFunctions.goAndGet("DBCONNECTION").Replace("DBNAME", "TEST")))
            //{

            //    // Attempt to open the connection and read the data.
            //    try
            //    {

            //        // Execute the query
            //        using (SqlCommand command = new SqlCommand(Query.Replace("%%", Limit.ToString())))
            //        {

            //            // Open the connection
            //            conn.Open();
            //            command.CommandType = System.Data.CommandType.Text;
            //            command.Connection = conn;
            //            Object obj = command.ExecuteScalar();

            //            // Grab the Average
            //            Average = Convert.ToInt32(obj);

            //            // Close the connection
            //            conn.Close();

            //        }

            //    }

            //    // If their were any errors, display them here
            //    catch (Exception e)
            //    {

            //        Console.WriteLine(e.ToString());
                    
            //    }

            //}

            //// Return the Average
            //return Average;          

            // Just return a random number for now as we need the new SAND database
            return Convert.ToInt32(rnd.Next(100, 999));

        }

        // Insert the data into the Performance table - Table doesn't exist yet
        public static void InsertIntoPerformanceTable(string UniqueID, int Time)
        {


        }

        public static void InformationBox(string Message, string Heading = "Information")
        {

            System.IO.StreamWriter reportmsg = new System.IO.StreamWriter(_startUp.resultsFileName, true);
            reportmsg.WriteLine("<div class=\"info\"><font face=\"verdana\" size=\"2\" color=\"#000000\"><u><strong>" + Heading + "</strong></u><br>" + Message + "</font></div>");
            reportmsg.Close();            

        }

        public static void WarningBox(string Message, string Heading = "Warning")
        {

            System.IO.StreamWriter reportmsg = new System.IO.StreamWriter(_startUp.resultsFileName, true);
            reportmsg.WriteLine("<div class=\"warning\"><font face=\"verdana\" size=\"2\" color=\"#000000\"><u><strong>" + Heading + "</strong></u><br>" + Message + "</font></div>");
            reportmsg.Close();

        }

    }

}

