using System;
using System.Text;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium;
using System.IO;
using OpenQA.Selenium.Chrome;
using System.Net;
using System.Data.SqlClient;
using System.Net.Sockets;
using System.Xml;
using Newtonsoft.Json.Linq;
using TestTeamAutomationFrameNew.Functions;
using TestTeamAutomationFrameNew.ITC_SMK;



namespace TestTeamAutomationFrameNew.GatherAPI
{
    class Functions
    {

        public static string Login (string Username, string Password)
        {

            string URL = "http://sand.mhealth.excointouch.com/gapi/api/User/Login";
            string Request = "{\"Username\":\"" + Username + "\", \"Password\":\"" + Password + "\"}";
            string Response = null;


            // Post the request
            //Response = mDNA.mDNA.PostRequest(URL, Request);

            // Return the Auth in the Response
            return GetAuthFromResponse(Response);

        }


        public static void ActivateStudyConfiguration(string StudyName, string ApiKey)
        {

            string URL = "http://sand.mhealth.excointouch.com/gapi/api/Study/ActivateStudyConfiguration";
            string Request = "{\"ActivationCode\":\"" + GetStudyActivationCode(StudyName) + "\"}";
            string Response = null;


            // Post the request
            //Response = mDNA.mDNA.PostJSONWithHeaders(URL, "", "", ApiKey, "", Request);

            Console.WriteLine("Response:\n" + Response);

        }

        public static void GetStudyConfiguration(string StudyName, string ApiKey)
        {

            string URL = "http://sand.mhealth.excointouch.com/gapi/api/Study/GetStudyConfiguration";
            string Request = "{\"StudyInstanceId\":\"" + GetStudyInstanceID(StudyName) + "\"}";
            string Response = null;


            // Post the request
            //Response = mDNA.mDNA.PostJSONWithHeaders(URL, "", "", ApiKey, "", Request);

            Console.WriteLine("Response:\n" + Response);

        }

        public static void ListPatientsBySite(string StudyName, string SiteName, string ApiKey, string State = "Design")
        {

            string URL = "http://sand.mhealth.excointouch.com/gapi/api/" + State + "/Subject/ListPatientsBySite";
            string Request = "{\"SiteId\":\"" + GetSiteInstanceID(StudyName, SiteName) + "\"}";
            string Response = null;


            // Post the request
            //Response = mDNA.mDNA.PostJSONWithHeaders(URL, "", "", ApiKey, "", Request);

            Console.WriteLine("Response:\n" + Response);

        }

        public static void RetrievePatientInformation(string StudyName, string SiteName, string SubjectID, string ApiKey, string State = "Design")
        {

            string URL = "http://sand.mhealth.excointouch.com/gapi/api/" + State + "/Subject/RetrievePatientInformation";
            string Request = "{\"Id\":\"" + GetSubjectID(StudyName, SiteName, SubjectID) + "\"}";
            string Response = null;


            // Post the request
            //Response = mDNA.mDNA.PostJSONWithHeaders(URL, "", "", ApiKey, "", Request);

            Console.WriteLine("Response:\n" + Response);

        }

        public static void RetrievePatientJourney(string StudyName, string SiteName, string SubjectID, int Type, string ApiKey, string State = "Design")
        {

            string URL = "http://sand.mhealth.excointouch.com/gapi/api/" + State + "/Subject/RetrievePatientJourney";
            string Request = "{\"PatientId\":\"" + GetSubjectID(StudyName, SiteName, SubjectID) + "\",\"Type\": " + Type + "}";
            string Response = null;


            // Post the request
            //Response = mDNA.mDNA.PostJSONWithHeaders(URL, "", "", ApiKey, "", Request);

            Console.WriteLine("Response:\n" + Response);

        }

        public static void ListPatientMessages(string StudyName, string SiteName, string SubjectID, string LanguageCode, string ApiKey, string State = "Design")
        {

            string URL = "http://sand.mhealth.excointouch.com/gapi/api/" + State + "/Subject/ListPatientMessages";
            string Request = "{\"PatientId\":\"" + GetSubjectID(StudyName, SiteName, SubjectID) + "\",\"LanguageId\":\"" + GetLanguageID(LanguageCode) + "\"}";
            string Response = null;


            // Post the request
           // Response = mDNA.mDNA.PostJSONWithHeaders(URL, "", "", ApiKey, "", Request);

            Console.WriteLine("Response:\n" + Response);

        }


        public static void ActivateSubject(string StudyName, string SiteName, string SubjectID, string ApiKey, string State = "Design")
        {

            string URL      = "http://sand.mhealth.excointouch.com/gapi/api/" + State + "/Subject/ActivateSubject";
            string Request  = "{\"StudyInstanceId\":\"" + GetStudyInstanceID(StudyName) + "\",\"ActivationCode\":\"" + GetSubjectActivationCode(StudyName, SiteName, SubjectID) + "\"}";
            string Response = null;


            // Post the request
           // Response = mDNA.mDNA.PostJSONWithHeaders(URL, "", "", ApiKey, "", Request);

            Console.WriteLine("Response:\n" + Response);

        }

        public static void ListQuestionnairesForStudy(string StudyName, string ApiKey, string State = "Design")
        {

            string URL = "http://sand.mhealth.excointouch.com/gapi/api/" + State + "/Assessment/ListQuestionnairesForStudy";
            string Request = "{\"StudyInstanceId\":\"" + GetStudyInstanceID(StudyName) + "\",\"Environment\": " + GetEnvID(State) + "}";
            string Response = null;


            // Post the request
            //Response = mDNA.mDNA.PostJSONWithHeaders(URL, "", "", ApiKey, "", Request);

            Console.WriteLine("Response:\n" + Response);

        }

        public static void RetrieveQuestionnaireDefinition(string StudyName, string QuestionnaireName, string ApiKey, string State = "Design")
        {

            string URL = "http://sand.mhealth.excointouch.com/gapi/api/" + State + "/Assessment/RetrieveQuestionnaireDefinition";
            string Request = "{\"Id\":\"" + GetQuestionnaireID(StudyName, QuestionnaireName, false) + "\"}";
            string Response = null;


            // Post the request
            //Response = mDNA.mDNA.PostJSONWithHeaders(URL, "", "", ApiKey, "", Request);

            Console.WriteLine("Response:\n" + Response);

        }


        public static void RetrievePatientJourney(string SubjectId, string StudyName, int Type, string ApiKey, string State = "Design")
        {

            string URL = "http://sand.mhealth.excointouch.com/gapi/api/" + State + "/Subject/RetrievePatientJourney";
            string Request = "{\"PatientId\":\"" +  SubjectId + "\",\"StudyInstanceId\":\"" + GetStudyInstanceID(StudyName) + "\",\"Type\": " + Type + "}";
            string Response = null;


            // Post the request
           // Response = mDNA.mDNA.PostJSONWithHeaders(URL, "", "", ApiKey, "", Request);

            Console.WriteLine("Response:\n" + Response);

        }

        #region Database Functions

        public static string GetSubjectActivationCode(string StudyName, string SiteName, string SubjectID, string State = "Design")
        {

            string Query            = "select Su.ShortCode from [Gather_R3_Design]..Subjects Su INNER JOIN[Gather_R3_Design]..SiteInstances Si ON Si.SiteInstanceId = Su.SiteInstanceId INNER JOIN[Gather_R3_Design]..StudyInstances StI ON StI.StudyInstanceId = Si.StudyInstanceId INNER JOIN[Gather_R3]..GatherStudies St ON St.GatherStudyId = StI.StudyId where St.StudyName = 'StudyName' and Si.SiteIdentifier = '" + SiteName + "' and Su.ClientFacingSubjectName = '" + SubjectID + "'"; 
            string ActivationCode   = null;


            try
            {
                // Setup the connection string
                using (SqlConnection conn = new SqlConnection(GenericFunctions.goAndGet("DBCONNECTION").Replace("DBNAME", "Gather_R3")))
                {
                    // Open the connection string
                    conn.Open();

                    // Read the data
                    using (SqlCommand command = new SqlCommand(Query))
                    {
                        command.CommandType = System.Data.CommandType.Text;
                        command.Connection  = conn;
                        Object obj          = command.ExecuteScalar();

                        // Grab the Activation Code
                        ActivationCode = obj.ToString();

                    }

                    // Close the connection
                    conn.Close();

                }
            }

            catch (Exception e)
            {


            }

            // Return the Activation Code
            return ActivationCode.ToUpper();

        }

        public static string GetStudyActivationCode(string StudyName)
        {

            string Query = "select Si.ShortCode from [Gather_R3_Design]..StudyInstances Si inner join[Gather_R3]..GatherStudies S ON S.GatherStudyId = SI.StudyId where S.StudyName = '" + StudyName + "'";
            string ActivationCode = null;


            try
            {
                // Setup the connection string
                using (SqlConnection conn = new SqlConnection(GenericFunctions.goAndGet("DBCONNECTION").Replace("DBNAME", "Gather_R3")))
                {
                    // Open the connection string
                    conn.Open();

                    // Read the data
                    using (SqlCommand command = new SqlCommand(Query))
                    {
                        command.CommandType = System.Data.CommandType.Text;
                        command.Connection  = conn;
                        Object obj          = command.ExecuteScalar();

                        // Grab the Activation Code
                        ActivationCode = obj.ToString();

                    }

                    // Close the connection
                    conn.Close();

                }
            }

            catch (Exception e)
            {


            }

            // Return the Activation Code
            return ActivationCode.ToUpper();

        }

        public static string GetStudyInstanceID(string StudyName)
        {

            string Query = "select Si.StudyInstanceID from [Gather_R3_Design]..StudyInstances Si inner join[Gather_R3]..GatherStudies S ON S.GatherStudyId = SI.StudyId where S.StudyName = '" + StudyName + "'";
            string InstanceID = null;


            // Setup the connection string
            using (SqlConnection conn = new SqlConnection(GenericFunctions.goAndGet("DBCONNECTION").Replace("DBNAME", "Gather_R3")))
            {
                // Open the connection string
                conn.Open();

                // Read the data
                using (SqlCommand command = new SqlCommand(Query))
                {
                    command.CommandType = System.Data.CommandType.Text;
                    command.Connection = conn;
                    Object obj = command.ExecuteScalar();

                    // Grab the Instance ID
                    InstanceID = obj.ToString();

                }

                // Close the connection
                conn.Close();

            }

            // Return the Instance ID
            return InstanceID.ToUpper();

        }


        public static string GetSiteID(string StudyName, string SiteName)
        {
            //Doesn't work yet

            string Query = "select Si.SiteId from Sites Si INNER JOIN GatherStudies St ON Si.StudyId = St.StudyId where St.ExcoStudyIdentifier = '" + StudyName + "' and Si.SiteIdentifier = '" + SiteName + "'";
            string SiteID = null;


            // Setup the connection string
            using (SqlConnection conn = new SqlConnection(GenericFunctions.goAndGet("DBCONNECTION").Replace("DBNAME", "Gather_R3")))
            {
                // Open the connection string
                conn.Open();

                // Read the data
                using (SqlCommand command = new SqlCommand(Query))
                {
                    command.CommandType = System.Data.CommandType.Text;
                    command.Connection = conn;
                    Object obj = command.ExecuteScalar();

                    // Grab the Site ID
                    SiteID = obj.ToString();

                }

                // Close the connection
                conn.Close();

            }

            // Return the Site ID
            return SiteID.ToUpper();

        }

        public static string GetSiteInstanceID(string StudyName, string SiteName, string State = "Design")
        {
            //Doesn't work yet

            string Query = "select Si.SiteInstanceId from [Gather_R3_" + State + "]..SiteInstances Si INNER JOIN[Gather_R3_" + State + "]..StudyInstances StI ON StI.StudyInstanceId = Si.StudyInstanceId INNER JOIN[Gather_R3]..GatherStudies St ON St.GatherStudyId = StI.StudyId where St.StudyName = '" + StudyName + "' and Si.SiteIdentifier = '" + SiteName + "'";
            string SiteInstanceID = null;


            // Setup the connection string
            using (SqlConnection conn = new SqlConnection(GenericFunctions.goAndGet("DBCONNECTION").Replace("DBNAME", "Gather_R3_" + State)))
            {
                // Open the connection string
                conn.Open();

                // Read the data
                using (SqlCommand command = new SqlCommand(Query))
                {
                    command.CommandType = System.Data.CommandType.Text;
                    command.Connection = conn;
                    Object obj = command.ExecuteScalar();

                    // Grab the Site Instance ID
                    SiteInstanceID = obj.ToString();

                }

                // Close the connection
                conn.Close();

            }

            // Return the Site Instance ID
            return SiteInstanceID.ToUpper();

        }

        public static string GetSubjectID(string StudyName, string SiteName, string SubjectID, string State = "Design")
        {
            //Doesn't work yet

            string Query = "select Su.SubjectId from [Gather_R3_Design]..Subjects Su INNER JOIN[Gather_R3_Design]..SiteInstances SiI ON Su.SiteInstanceId = SiI.SiteInstanceId INNER JOIN[Gather_R3_Design]..StudyInstances StI ON SiI.StudyInstanceId = StI.StudyInstanceId INNER JOIN[Gather_R3]..GatherStudies St ON StI.StudyId = St.StudyId where Su.ClientFacingSubjectName = '" + SubjectID + "' and SiI.SiteIdentifier = '" + SiteName + "' and St.ExcoStudyIdentifier = '" + StudyName + "'";
            string IntSubjectID = null;


            // Setup the connection string
            using (SqlConnection conn = new SqlConnection(GenericFunctions.goAndGet("DBCONNECTION").Replace("DBNAME", "Gather_R3" + "_" + State.ToLower())))
            {
                // Open the connection string
                conn.Open();

                // Read the data
                using (SqlCommand command = new SqlCommand(Query))
                {
                    command.CommandType = System.Data.CommandType.Text;
                    command.Connection = conn;
                    Object obj = command.ExecuteScalar();

                    // Grab the Subject ID
                    IntSubjectID = obj.ToString();

                }

                // Close the connection
                conn.Close();

            }

            // Return the SubjectID
            return IntSubjectID.ToUpper();

        }


        public static string GetLanguageID(string LanguageCode)
        {
            //Doesn't work yet

            string Query = "select LanguageId from Languages where Iso639 = '" + LanguageCode + "'";
            string LanguageID = null;


            // Setup the connection string
            using (SqlConnection conn = new SqlConnection(GenericFunctions.goAndGet("DBCONNECTION").Replace("DBNAME", "Gather_R3")))
            {
                // Open the connection string
                conn.Open();

                // Read the data
                using (SqlCommand command = new SqlCommand(Query))
                {
                    command.CommandType = System.Data.CommandType.Text;
                    command.Connection = conn;
                    Object obj = command.ExecuteScalar();

                    // Grab the Language ID
                    LanguageID = obj.ToString();

                }

                // Close the connection
                conn.Close();

            }

            // Return the Language ID
            return LanguageID.ToUpper();

        }

        public static string GetQuestionnaireID(string StudyName, string QuestionnaireName, bool PublishedID = true, string State = "Design")
        {
            //Doesn't work yet

            string Query = "select " + (PublishedID ? "PQ.ABPublishedQuestionnaireId" : "PQ.QuestionnaireGlobalIdentifier") + " from [Gather_R3_" + State + "]..ABPublishedQuestionnaires PQ INNER JOIN[Gather_R3_" + State + "]..StudyInstances StI ON StI.StudyInstanceId = PQ.StudyInstanceId INNER JOIN[Gather_R3]..GatherStudies St ON St.GatherStudyId = StI.StudyId where St.StudyName = '" + StudyName + "' and PQ.Name = '" + QuestionnaireName + "'"; 
            string QuestionnaireID = null;


            // Setup the connection string
            using (SqlConnection conn = new SqlConnection(GenericFunctions.goAndGet("DBCONNECTION").Replace("DBNAME", "Gather_R3")))
            {
                // Open the connection string
                conn.Open();

                // Read the data
                using (SqlCommand command = new SqlCommand(Query))
                {
                    command.CommandType = System.Data.CommandType.Text;
                    command.Connection = conn;
                    Object obj = command.ExecuteScalar();

                    // Grab the Questionnaire ID
                    QuestionnaireID = obj.ToString();

                }

                // Close the connection
                conn.Close();

            }

            // Return the Questionnaire ID
            return QuestionnaireID.ToUpper();

        }


        #endregion

        public static string GetAuthFromResponse(string Response)
        {

            string Auth = null;




            return Auth;
        }

        public static string GetEnvName(int ID)
        {

            string EnvName = null;


            switch (ID)
            {
                case 0:
                    EnvName = "Design";
                    break;

                case 1:
                    EnvName = "UAT";
                    break;

                case 2:
                    EnvName = "Live";
                    break;

                default:
                    break;
                    
            }

            return EnvName;

        }

        public static int GetEnvID(string State)
        {

            int EnvID = 0;


            switch (State)
            {
                case "Design":
                    EnvID = 0;
                    break;

                case "UAT":
                    EnvID = 1;
                    break;

                case "Live":
                    EnvID = 2;
                    break;

                default:
                    break;

            }

            return EnvID;

        }

    }

}
