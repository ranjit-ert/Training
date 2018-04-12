//using Selenium.PageTemplates;
using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using TestTeamAutomationFrameNew.Functions;


namespace TestTeamAutomationFrameNew.Functions
{
   public class Windows
    {
        public static void ClearFile(string FileName)
        {

            TextWriter tsw = new StreamWriter(FileName);
            tsw.WriteLine();
            tsw.Close();

        }

        public static string Tab()
        {

            return "&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;";

        }

        public static string ReadToString(string FilePath)
        {

            StreamReader streamReader = new StreamReader(FilePath, Encoding.UTF8);
            string text = streamReader.ReadToEnd();
            streamReader.Close();
            return text;

        }

        public static void CopyFile(string FileName, string sourceDir, string destinationDir)
        {

            string fileName = FileName;
            string sourcePath = sourceDir;
            string targetPath = destinationDir;
            string sourceFile = System.IO.Path.Combine(sourcePath, fileName);
            string destFile = System.IO.Path.Combine(targetPath, fileName);


            if (!System.IO.Directory.Exists(targetPath))
            {
                System.IO.Directory.CreateDirectory(targetPath);
            }


            System.IO.File.Copy(sourceFile, destFile, true);

            if (System.IO.Directory.Exists(sourcePath))
            {
                string[] Files = System.IO.Directory.GetFiles(sourcePath);

                // Copy the Files and overwrite destination Files if they already exist. 
                foreach (string s in Files)
                {
                    // Use static Path methods to extract only the File name from the path.
                    fileName = System.IO.Path.GetFileName(s);
                    destFile = System.IO.Path.Combine(targetPath, fileName);
                    System.IO.File.Copy(s, destFile, true);
                }
            }
            else
            {

                Console.WriteLine("Source path does not exist!");

            }

        }

        static public void CopyDir(string sourceDirectory, string destDirectory)
        {
            if (!Directory.Exists(destDirectory))
            {
                Directory.CreateDirectory(destDirectory);
            }
            string[] Files = Directory.GetFiles(sourceDirectory);
            foreach (var myfile in Files)
            {
                string name = Path.GetFileName(myfile);
                string dest = Path.Combine(destDirectory, name);
                Console.WriteLine("[{0} -> {1}]", name, dest);
                System.IO.File.Copy(myfile, dest);
            }
            string[] directories = Directory.GetDirectories(sourceDirectory);
            foreach (var directory in directories)
            {
                string name = Path.GetFileName(directory);
                string dest = Path.Combine(destDirectory, name);


                CopyDir(directory, dest);
            }
        }

        public static void MoveFile(string FileName, string sourceFile, string destinationFile)
        {
            //string sourceFile = @"C:\Users\Public\public\test.txt";
            //string destinationFile = @"C:\Users\Public\private\test.txt";

            // To move a File or folder to a new location:
            System.IO.File.Move(sourceFile + "\\" + FileName, destinationFile + "\\" + FileName);

            // To move an entire directory. To programmatically modify or combine 
            // path strings, use the System.IO.Path class.
            //System.IO.Directory.Move(@"C:\Users\Public\public\test\", @"C:\Users\Public\private");
        }

        public static void MoveDir(string sourceDir, string destinationDir)
        {

            System.IO.Directory.Move(sourceDir, destinationDir);

        }

        public static void StartBatchFile(string FileName, string Parameters)
        {

            Process proc = null;
            try
            {
                proc = new Process();
                proc.StartInfo.FileName = FileName;
                proc.StartInfo.Arguments = string.Format(Parameters);
                proc.StartInfo.CreateNoWindow = false;
                proc.Start();
                proc.WaitForExit();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception Occurred :{0},{1}", ex.Message, ex.StackTrace.ToString());
            }

        }

        public static void WaitForFileToExist(string FileName)
        {

            Console.WriteLine("Waiting for " + FileName + " to exist...");
            while (File.Exists(FileName) == false) {
               
                GenericFunctions.Wait(1);

            }
            Console.WriteLine(FileName + " exists! Carry on.");


        }

        public static void DeleteFile(string FileName)
        {
            
            if (File.Exists(FileName))
            {

                File.Delete(FileName);
                Console.WriteLine(File.Exists(FileName) ? "File was not deleted." : "File deleted successfully!");

            }
            else
            {

                Console.WriteLine("File wasn't there, so no need to delete anything.");

            }
          
        }

        public static void WriteToFile(string Text, string FileName)
        {

            TextWriter tsw = new StreamWriter(FileName, true);

            tsw.WriteLine(Text);
            tsw.Close();

        }

        public static void NewFile(string Text, string FileName)
        {


            // Write the string to a File.
            System.IO.StreamWriter File = new System.IO.StreamWriter(FileName);
            File.WriteLine(Text);

            File.Close();

        }

        public static void ChangeDateTimeOnDevice(string DateTime)
        {
            Process proc = null;
            try
            {
                proc = new Process();
                proc.StartInfo.FileName = @"%ANDROID_HOME%/tools/adb.exe";
                proc.StartInfo.Arguments = string.Format("shell date -s " + DateTime);
                proc.StartInfo.CreateNoWindow = true;
                proc.Start();
                proc.WaitForExit();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception Occurred :{0},{1}", ex.Message, ex.StackTrace.ToString());
            }
        }

        public static string GetMyUsername(bool Nice = false)
        {

            string Username = System.Security.Principal.WindowsIdentity.GetCurrent().Name.Replace(@"CORP\", "");;
            

            // We we need to convert the username from "firstname.lastname" to "Firstname Lastname"...
            if (Nice == true)
            {
                char[] delimiterChars = { '.' };
                string[] FirstLastName = Username.Split(delimiterChars);

                Username = FirstLastName[0] + " " + FirstLastName[1];
                Username = System.Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(Username.ToLower());
            }

            return Username;

        }

    }

}
