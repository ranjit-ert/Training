using System;
using System.Data;
using System.Data.SqlClient;
using OpenQA.Selenium;
using TestTeamAutomationFrameNew.ITC_SMK;

namespace TestTeamAutomationFrameNew.Functions
{
    /// <summary>
    /// This class contains all of the database functions
    /// </summary>
    public class _dbActions
    {
        public static string GetQueryResult(string sSQL)
        {
            //This function can only be used to return single pieces of data from an SQL query.  IE.  it won't return multiple columns\rows.
            Console.WriteLine("GetQueryResult Starting Function:  sSQL=  " + sSQL);
            string sResult;
            using (var conn = new SqlConnection(GenericFunctions.goAndGet("DBCONNECTION")))
            {
                conn.Open();
                using (var command = new SqlCommand(sSQL))
                {
                    command.CommandType = System.Data.CommandType.Text;
                    command.Connection = conn;
                    var obj = command.ExecuteScalar();
                    sResult = obj.ToString();
                }
                conn.Close();
                 return sResult;
            }
        }
       public static bool QueryToDataSet(string sSQL, ref DataSet ds)
        {
            //This function will return query results as a dataset and can therefore return multiple columns\rows.
            bool bResult = false;
            var sqLconn = new SqlConnection(GenericFunctions.goAndGet("DBCONNECTION"));
            var sqLcommand = new SqlCommand(sSQL, sqLconn);
            var adapter = new SqlDataAdapter(sqLcommand);
            sqLconn.Close();
            var sqlDataSet = new DataSet();
             try
            {
                adapter.Fill(sqlDataSet, "SqlDataTable");
            }
            catch
            {
                //Fill error - abort and return false
                Console.WriteLine("QueryToArray:  Finised Function [FAIL - Fill Error]");
                return bResult;
            }
            if (sqlDataSet.Tables.Count == 0)
            {
                //NO data found - abort and return false
                Console.WriteLine("QueryToArray:  Finised Function [FAIL - Table Count = 0]");
                return false;
            }
           //Data found - return the DataSet
           ds = sqlDataSet;
           
           return true;
        }


        public static int GetRowIdForSubject(int subjectid, IWebDriver driver)
        {
            var returninternalsubjectid = 0;
            using (var conn = new SqlConnection(GenericFunctions.goAndGet("DBCONNECTION")))
            {
                conn.Open();
                using (var command = new SqlCommand("select subject_id from tbl_subject where user_subject_id = '" + subjectid + "' AND site_id = '" + GetSiteID(driver) + "'"))
                {
                    command.CommandType = System.Data.CommandType.Text;
                    command.Connection = conn;
                    var obj = command.ExecuteScalar();
                    returninternalsubjectid = Convert.ToInt32(obj.ToString());
                }

                conn.Close();
                return returninternalsubjectid;
            }
        }

        public static int GetLatestSubjectId(string studyName, string siteName, IWebDriver driver)
        {
            int tmpSubjectID;
            int SubjectID;
            //int x = 00000;
            using (var conn = new SqlConnection(GenericFunctions.goAndGet("DBCONNECTION")))
            {
                conn.Open();
                using (var command = new SqlCommand("select user_subject_id from tbl_subject where site_id = '" + GetSiteID(driver) + "' ORDER BY user_subject_id DESC"))
                {
                    command.CommandType = System.Data.CommandType.Text;
                    command.Connection = conn;
                    var obj = command.ExecuteScalar();
                    tmpSubjectID = Convert.ToInt32(obj.ToString());
                    //tmpSubjectID = Int32.TryParse(obj.ToString, out x());                    
                   SubjectID = tmpSubjectID + 1;
                }
                conn.Close();
                return SubjectID;

            }
        }

        public static void CheckTblSubject(int subjectId, string siteName, string studyName, IWebDriver driver)
        {
            using (SqlConnection conn = new SqlConnection(GenericFunctions.goAndGet("DBCONNECTION")))
            {

                conn.Open();

                using (SqlCommand command = new SqlCommand("select subject_id from tbl_subject where user_subject_id = '" + subjectId + "' AND Site_ID = '" + GetSiteID(driver) + "'"))
                {
                    command.CommandType = System.Data.CommandType.Text;
                    command.Connection = conn;

                    Object obj = command.ExecuteScalar();
                    subjectId = Convert.ToInt32(obj.ToString()); ;

                }


                System.Console.WriteLine("Checked tbl_subjects. Subject " + subjectId + " exists.");

                using (SqlCommand command = new SqlCommand("SELECT COUNT(outbound_msg_id) AS count FROM tbl_outbound_msg WHERE subject_id = '" + subjectId + "'"))
                {
                    command.CommandType = System.Data.CommandType.Text;
                    command.Connection = conn;



                    Object obj = command.ExecuteScalar();
                    int messages = Convert.ToInt32(obj.ToString());

                    if (messages == Convert.ToInt32(GenericFunctions.goAndGet("MESSAGES")))
                    {
                        GenericFunctions.ReportTheTestPassed("Subject " + subjectId + " contains the correct number of messages: " + messages);
                    }
                    else
                    {
                        GenericFunctions.ReportTheTestFailed("Subject " + subjectId + " contained " + messages + " messages, but " + Convert.ToInt32(GenericFunctions.goAndGet("MESSAGES")) + " were expected.");
                        _startUp.stepsFailed++;
                    }

                }

                conn.Close();

            }

        }

        public static int GetStudyId(IWebDriver driver)
        {
            using (SqlConnection conn = new SqlConnection(GenericFunctions.goAndGet("DBCONNECTION")))
            {

                conn.Open();
                int StudyID;
                using (var command = new SqlCommand("SELECT study_id from tbl_study where name = '" + GenericFunctions.goAndGet("STUDYID") + "'"))
                {
                    command.CommandType = System.Data.CommandType.Text;
                    command.Connection = conn;
                    var obj = command.ExecuteScalar();
                    StudyID = Convert.ToInt32(obj.ToString());
                }

                conn.Close();

                return StudyID;

            }
        }

        private static int GetSiteID(IWebDriver driver)
        {
            using (var conn = new SqlConnection(GenericFunctions.goAndGet("DBCONNECTION")))
            {

                conn.Open();

                int siteId;
                using (var command = new SqlCommand("SELECT site_id from tbl_site where study_id = '" + GetStudyId(driver) + "'"))
                {
                    command.CommandType = System.Data.CommandType.Text;
                    command.Connection = conn;

                    Object obj = command.ExecuteScalar();
                    siteId = Convert.ToInt32(obj.ToString());

                }
                conn.Close();

                return siteId;

            }
        }

        public static int CountSubjects(IWebDriver driver)
        {
            var studyName = GenericFunctions.goAndGet("STUDYID");
            var siteName = GenericFunctions.goAndGet("SITEID");


            using (SqlConnection conn = new SqlConnection(GenericFunctions.goAndGet("DBCONNECTION")))
            {

                conn.Open();


                int subjects;
                using (var command = new SqlCommand("SELECT COUNT(subject_id) AS Subjects FROM tbl_subject WHERE site_id = '" + GetSiteID(driver) + "'"))
                {
                    command.CommandType = System.Data.CommandType.Text;
                    command.Connection = conn;
                    var obj = command.ExecuteScalar();
                    subjects = Convert.ToInt32(obj.ToString());
                }

                conn.Close();
                return subjects;

            }
        }
        public static int CheckWithdrawn(int IntSubjectID, IWebDriver driver)
        {
            string studyName = GenericFunctions.goAndGet("STUDYID");
            string siteName = GenericFunctions.goAndGet("SITEID");
            int Messages;


            using (SqlConnection conn = new SqlConnection(GenericFunctions.goAndGet("DBCONNECTION")))
            {

                conn.Open();


                int subjects;
                using (SqlCommand command = new SqlCommand("SELECT COUNT(subject_id) AS Subjects FROM tbl_subject WHERE subject_id = '" + IntSubjectID + "'  AND site_id = '" + GetSiteID(driver) + "'"))
                {
                    command.CommandType = System.Data.CommandType.Text;
                    command.Connection = conn;

                    Object obj = command.ExecuteScalar();
                    subjects = Convert.ToInt32(obj.ToString());

                }
             using (var command = new SqlCommand("SELECT COUNT(outbound_msg_id) AS count FROM tbl_outbound_msg WHERE subject_id = '" + IntSubjectID + "'"))
                {
                    command.CommandType = System.Data.CommandType.Text;
                    command.Connection = conn;
                    var obj = command.ExecuteScalar();
                    Messages = Convert.ToInt32(obj.ToString());

                }

                conn.Close();

                return (subjects == 0 && Messages == 0 ? 1 : 0);

            }
        }

        public static string GetMessageStatus(string number, string date, IWebDriver driver)
        {
            using (var conn = new SqlConnection(GenericFunctions.goAndGet("ATLASDBCONNECTION")))
            {
                //date format for query must be yyyy-mm-dd hh:mm:ss.msmsms e.g. 2015-06-16 15:30:00.000
                conn.Open();

                string messageStatus;
                using (var command = new SqlCommand("SELECT a.name FROM [dbo].[tlkp_aggregator_status] AS a INNER JOIN [dbo].[tbl_sms_out] AS so ON a.status_id = so.aggregator_status_ind WHERE so.schedule_time_gmt = '" + date + "' AND dbo.func_aes256_decrypt(so.enc_msisdn_enc) = '" + number + "'"))
                {
                    command.CommandType = System.Data.CommandType.Text;
                    command.Connection = conn;

                    Object obj = command.ExecuteScalar();
                    messageStatus = obj.ToString();

                }
                conn.Close();

                return messageStatus;

            }
        }

    }
}
