using System;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.PhantomJS;
using OpenQA.Selenium.Remote;
using TestTeamAutomationFrameNew.Functions;
using System.IO;
using System.Drawing;
using System.Threading;
using System.Linq;
using System.Diagnostics;
using SelectPdf;
using System.Data;
using System.Data.SqlClient;
using Newtonsoft.Json.Linq;

namespace TestTeamAutomationFrameNew.Training
{

    class TrainingFunctions
    {

        private static string[] StoreKeyPress = new string[4];

        //Function accepting string as a parameter as well as an optional parameter
        public static void LoginToGather(string Username, string Password, IWebDriver driver, string Org = null, string Study = null, string State = null)
        {
            GenericFunctions.Type(Username, "Username", driver);
            GenericFunctions.Type(Password, "Password", driver);
            GenericFunctions.ClickOnButton("login-submit-button", driver);

            Console.WriteLine("Org is " + Org);

        }

        //Function accepting int as a parameter
        public static int PreCalc(int Value1, int Value2)
        {
            int NewValue = Value1 + Value2;

            return NewValue; 

        }

        //Function returning a string
        public static string GetString()
        {

            return "Go";

        }

        //Function returning an integer
        public static int GetInteger()
        {

            return 5;

        }




        public static int Calculator(int number1, int number2, string operation)
        {
            int answer;

            //Console.WriteLine(RecordKeyPresses[0]);
            //Console.WriteLine(RecordKeyPresses[1]);
            //Console.WriteLine(RecordKeyPresses[2]);
            //Console.WriteLine(RecordKeyPresses[3]);

            //Console.WriteLine(number1);
            //Console.WriteLine(number2);
            //Console.WriteLine(operation);


            switch (operation)
            {
                case "+":
                    answer = number1 + number2;
                    //return answer;
                    //return answer.ToString();
                    break;

                case "-":
                    answer = number1 - number2;
                    //return answer;
                    //return answer.ToString();
                    break;

                case "/":
                    answer = number1 / number2;
                    //return answer;
                    //return answer.ToString();
                    break;

                case "÷":
                    answer = number1 / number2;
                    //return answer;
                    //return answer.ToString();
                    break;

                case "*":
                    answer = number1 * number2;
                    //return answer;
                    //return answer.ToString();
                    break;

                default:
                    //Console.WriteLine("======Invalid operator: " + operation + ", ERROR 99999:");
                    //return "ERROR: Invalid Operator specified";
                    return 9999;
                    //break;
            }

            //return answer.ToString();

            //reset the array
            //Array.Clear(RecordKeyPresses, 0, RecordKeyPresses.Length);
           //Console.WriteLine(RecordKeyPresses[1]);

            return answer;

        }

        //GenericFunctions.ClickOnButtonByXPath("//span[contains(text(), '" + input + "')]", driver);

        public static void CalculatorButton(string input, IWebDriver driver, int inputOrder)
        {
            
            GenericFunctions.ClickOnButtonByXPath("//*[@class='cwbts'][contains(text(),'" + input + "')]", driver);
            StoreKeyPress [inputOrder] = input;

        }

        public static int GetGoogleCalcAnswer(IWebDriver driver)
        {
            int answer = Convert.ToInt32(driver.FindElement(By.Id("cwos")).Text);


            return answer;
        }
       

        public static int GetExpectedAnswer()
        {

            //Console.WriteLine(RecordKeyPresses[1]);


            return Calculator(Convert.ToInt32(StoreKeyPress[0]), Convert.ToInt32(StoreKeyPress[2]), StoreKeyPress[1]);
            

            

        }

        public static void ClearKeyPressArray()
        {


            //for (int i=0; i<RecordKeyPresses.Length; i++)
            //{
            //    RecordKeyPresses[i] = "a";
            //    Console.WriteLine("Array" + i +": " + RecordKeyPresses[i]);


            //}

            Array.Clear(StoreKeyPress, 0, StoreKeyPress.Length);
            //Console.WriteLine("Array0: " + RecordKeyPresses[0]);
            //Console.WriteLine("Array1: " + RecordKeyPresses[1]);
            //Console.WriteLine("Array2: " + RecordKeyPresses[2]);
            //Console.WriteLine("Array3: " + RecordKeyPresses[3]);

        }




    }

    
       

}
