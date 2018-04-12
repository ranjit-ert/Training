using System;
using System.Collections.Generic;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Chrome;
using TestTeamAutomationFrameNew.Functions;


namespace TestTeamAutomationFrameNew.Training
{

    public class TrainingTests
    {

        #region Setup

        public static IWebDriver driver;


        [SetUp]
        public void SetUp()
        {
            

        }

        #endregion

        [Test]
        public static void TrainingTest()
        {
            driver = new ChromeDriver();
            driver.Manage().Window.Maximize();

            // Load the test page
            driver.Navigate().GoToUrl("https://bbc.co.uk");
            GenericFunctions.ClickOnLink("Sport", driver);
            GenericFunctions.ClickOnLink("About the BBC", driver);
            //GenericFunctions.UnTickCheckBox
            GenericFunctions.Wait(5);
            driver.Quit();

            //Type
            //TickCheckBox
            //UnTickCheckbox - By.Id
            //SelectFromList
            //ClickOnButton
            //ClickOnLink
            //CheckTextIsOnPage
            //CheckTextIsNotOnPage

        }

        [Test]
        public static void LoginToGather()
        {
            driver = new ChromeDriver();
            driver.Manage().Window.Maximize();

            //Go to Gather
            driver.Navigate().GoToUrl("https://engagement.int.ert.com");

            //Check a few pieces of text on the login page
            GenericFunctions.CheckTextIsOnPageNew("Username", driver);
            GenericFunctions.CheckTextIsOnPageNew("Password", driver);
            GenericFunctions.CheckTextIsOnPageNew("Forgot Password", driver);
            GenericFunctions.CheckTextIsOnPageNew("Forgot Username", driver);
            GenericFunctions.CheckTextIsOnPageNew("LOGIN", driver);
            GenericFunctions.CheckTextIsOnPageNew("Terms & Conditions", driver);
            GenericFunctions.CheckTextIsOnPageNew("Privacy Policy", driver);
            
            //Enter login credentials
            GenericFunctions.Type("ccraiser.studyall", "Username", driver);
            GenericFunctions.Type("User1234", "Password", driver);
            GenericFunctions.Wait(3);

            //Login
            GenericFunctions.ClickOnButton("login-submit-button", driver);
            GenericFunctions.Wait(3);
            
            //Select the tabs at the top of the page.
            GenericFunctions.ClickOnLink("Reporting", driver);
            GenericFunctions.Wait(3);
            GenericFunctions.ClickOnLink("Change Control", driver);

            //Close Browser
            GenericFunctions.Wait(5);
            driver.Quit();

        }

        [Test]
        public static void NavigateAroundAnotherSite()
        {
            driver = new ChromeDriver();
            driver.Manage().Window.Maximize();

            //Go to Training HTML
            driver.Navigate().GoToUrl("file:///C:/Ranjit/Automation/TrainingTasks/IntroductionToAutomation/Testng.html");
            GenericFunctions.Wait(2);

            //Click on a hyperlink
            GenericFunctions.ClickOnLink("Click Me", driver);
            GenericFunctions.Wait(2);

            //Navigate Back
            driver.Navigate().Back();
            GenericFunctions.Wait(2);

            //Select a value from a Dropdown
            GenericFunctions.SelectFromList("Audi", "CarsList", driver);
            GenericFunctions.Wait(2);

            //Tick / Untick a checkbox
            GenericFunctions.TickCheckBox("Bike", driver);
            GenericFunctions.Wait(2);
            GenericFunctions.UnTickCheckBox("Bike", driver);
            GenericFunctions.Wait(2);

            //Check text is on a page.
            GenericFunctions.CheckTextIsOnPageNew("Login", driver);
            GenericFunctions.CheckTextIsOnPageNew("Username:", driver);
            GenericFunctions.CheckTextIsOnPageNew("Password:", driver);
            GenericFunctions.CheckTextIsOnPageNew("What transport do you have?", driver);
            GenericFunctions.CheckTextIsOnPageNew("What car do you have?", driver);
           
            //Check text is not on a page.
            GenericFunctions.CheckTextIsNotOnPageNew("Elephant", driver);
            GenericFunctions.CheckTextIsNotOnPageNew("Tiger", driver);
            GenericFunctions.CheckTextIsNotOnPageNew("Dog", driver);
            GenericFunctions.CheckTextIsNotOnPageNew("Cat", driver);

            //Close Browser
            GenericFunctions.Wait(2);
            driver.Quit();

        }

        [Test]
        public static void IFTest()
        {
            driver = new ChromeDriver();
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("file:///C:/Ranjit/Automation/TrainingTasks/AddingInLogic/TestPages/IF.html");


            string UserRole = driver.FindElement(By.Id("ShowText")).Text;

            if (UserRole == "Study User")
            {
                GenericFunctions.ClickOnLink("Link 1", driver);

            }
            else if (UserRole == "Site User")
            {
                GenericFunctions.ClickOnLink("Link 2", driver);
            }


            //Close Browser
            GenericFunctions.Wait(2);
            driver.Quit();

        }



        [Test]
        public static void SwitchTest()
        {
            driver = new ChromeDriver();
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("file:///C:/Ranjit/Automation/TrainingTasks/AddingInLogic/TestPages/Switch.html");


            string System = driver.FindElement(By.Id("ShowText")).Text;
            GenericFunctions.Wait(3);

            switch (System)
            {
                case "25%":
                    GenericFunctions.ClickOnLink("Link 1", driver);
                    break;

                case "50%":
                    GenericFunctions.ClickOnLink("Link 2", driver);
                    break;

                case "75%":
                    GenericFunctions.ClickOnLink("Link 3", driver);
                    break;

                case "100%":
                    GenericFunctions.ClickOnLink("Link 4", driver);
                    break;

                default:
                    break;
            }

            //Close Browser
            GenericFunctions.Wait(2);
            driver.Quit();

        }

        [Test]
        public static void WhileTest()
        {
            driver = new ChromeDriver();
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("file:///C:/Ranjit/Automation/TrainingTasks/AddingInLogic/TestPages/While.html");
            int counter = 0;


            while (!driver.FindElement(By.Id("TheImgId")).Displayed)
            {

                Console.WriteLine(counter.ToString());
                GenericFunctions.Wait(1);

                if (counter==60)
                {

                    break;
                }

                counter++;

            }

            GenericFunctions.ClickOnLink("Click me", driver);

            //Close Browser
            GenericFunctions.Wait(2);
            driver.Quit();
        }


        [Test]
        public static void GatherLogin()
        {
            driver = new ChromeDriver();
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("https://engagement.int.ert.com");

            TrainingFunctions.LoginToGather("ccraiser.studyall", "User1234", driver, "RanjitTest2");

            int Calc = TrainingFunctions.PreCalc(1, 10);

            Console.WriteLine(Calc);

            //Close Browser
            GenericFunctions.Wait(2);
            driver.Quit();
        }

        [Test]
        public static void CalculateTwoNumbers()
        {
            
            Console.WriteLine("Adding {0} and {1} gives {2}" , 5, 7, TrainingFunctions.Calculator(5, 7, "+"));
            Console.WriteLine("Subtracting {0} from {1} gives {2}", 6, 20, TrainingFunctions.Calculator(20, 6, "-"));
            Console.WriteLine("Dividing {0} by {1} gives {2}", 120, 12, TrainingFunctions.Calculator(120, 12, "/"));
            Console.WriteLine("Multiplying {0} by {1} gives {2}", 8, 9, TrainingFunctions.Calculator(8, 9, "*"));
            Console.WriteLine("Multiplying {0} by {1} gives {2}", 8, 9, TrainingFunctions.Calculator(8, 9, "@"));


        }

        [Test]
        public static void UseString()
        {

            if (TrainingFunctions.GetString() == "Go")
            { 
                Console.WriteLine("Expected string was returned");
            }
            else
            {
                Console.WriteLine("Expected string was not returned");
            }


        }

        [Test]
        public static void UseInteger()
        {
            int counter = 0;

            while (TrainingFunctions.GetInteger() == 5)
            {
                Console.WriteLine("Expected integer was returned");

                if (counter == 2)
                {

                    break;
                }

                counter++;
            }
        }


        [Test]
        public static void GoogleCalc()
        {

            driver = new ChromeDriver();
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("https://www.google.co.uk/search?&q=calc");
            int counter = 0;


            //Wait for the Google calculator to be visible
            while (!driver.FindElement(By.Id("cwmcwd")).Displayed)
            {
                Console.WriteLine(counter.ToString());
                GenericFunctions.Wait(1);

                if (counter == 60)
                {

                    break;
                }

                counter++;
            }

            // Type some numbers testing
            TrainingFunctions.CalculatorButton("4", driver, 0);
            TrainingFunctions.CalculatorButton("+", driver, 1);
            TrainingFunctions.CalculatorButton("8", driver, 2);
            TrainingFunctions.CalculatorButton("=", driver, 3);
            //TrainingFunctions.CalculatorButton("−", driver);
            //TrainingFunctions.CalculatorButton("1", driver);
            //TrainingFunctions.CalculatorButton("=", driver);
            //TrainingFunctions.CalculatorButton("+", driver);
            //TrainingFunctions.CalculatorButton("4", driver);
            //TrainingFunctions.CalculatorButton("=", driver);
            //TrainingFunctions.CalculatorButton("×", driver);
            //TrainingFunctions.CalculatorButton("5", driver);
            //TrainingFunctions.CalculatorButton("=", driver);

            
            // Report the result
            Console.WriteLine("Google Calc Answer: " + TrainingFunctions.GetGoogleCalcAnswer(driver));

            if (TrainingFunctions.GetExpectedAnswer()==9999)
            {
                Console.WriteLine("ERROR: Invalid Operator specified");

            }
            else
            {

                Console.WriteLine("Expected Answer: " + TrainingFunctions.GetExpectedAnswer());

                if (TrainingFunctions.GetGoogleCalcAnswer(driver)==TrainingFunctions.GetExpectedAnswer())
                {
                    Console.WriteLine("PASS: Google calculator answer agrees with Expected answer");
                }
                else
                {
                    Console.WriteLine("FAIL: Google calculator answer does not agree with Expected answer");
                }

            
            }

            TrainingFunctions.ClearKeyPressArray();


            //Close Browser
            GenericFunctions.Wait(3);
            driver.Quit();

        }

        
    

        [TearDown]
        public void TearDown()
        {

            //TrainingFunctions.ClearArray();
            

        }

    }

}