using System.Diagnostics;
using System.Net.Mime;

namespace TestTeamAutomationFrameNew.Functions
{
    public class Globals
    {
        public static class Run
        {
            public static readonly string PsEnvironment = GenericFunctions.goAndGet("RUNENVIRONMENT").ToUpper();
            public static readonly string PsRunTime = GenericFunctions.GetSortDate();
            public static readonly string PsResultsFolder = string.Format(@"C:\\Automation\\BOLT\\RUNS\\{0}\\{1}\\", PsEnvironment, PsRunTime);
            public static string PsAndriodResultsFolder = string.Format(@"C:\\Automation\\Android\\RUNS\\{0}\\{1}\\", PsEnvironment, PsRunTime);
            public static string PsDebugMode = GenericFunctions.goAndGet("DebugMode").ToUpper(); //this needs to be a boolean value
            public const string PsAppiumpath = @"/C node C:\Users\Wonderson.Chideya\Documents\Appium\AppiumForWindows-0.9.1\Appium\node_modules\appium";
            public const string PsAppiumserver = "127.0.0.1";
            public const string PsAppiumport = "4723";
            public static string PsOnGoing = "Yes";
        }
        public static class Itc
        {
            static Itc() { }
            public static readonly string Username = GenericFunctions.goAndGet("BoltUserName");
            public static readonly string Password = GenericFunctions.goAndGet("BoltPassword");
        }
    }

}
