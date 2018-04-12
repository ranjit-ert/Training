using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using OpenQA.Selenium;
using OpenQA.Selenium.Remote;


namespace TestTeamAutomationFrameNew.Functions
{
    public class merckVRC
    {

        public static string enrolSubject(string loginPIN, string subjectID, string centreID, AndroidDriver driver)
        {
            string sTempPIN = "false";
            //login as site admin

            //standroid.GetScreenShot(driver, "test");
            //standroid.takeScreenshot(driver);



            driver.FindElementByClassName("android.widget.EditText").SendKeys(loginPIN);
            driver.FindElementByName("Next").Click();
            //select New Subject
            GenericFunctions.Wait(5);
            driver.FindElementByName("New Subject").Click();
            //enter new subject enrollment details
            GenericFunctions.Wait(1);

            //standroid.GetScreenShot(driver, "screenshot");

            orangutan.CaptureElementCollection("android.widget.EditText", driver);
            var fields = driver.FindElementsByClassName("android.widget.EditText");
            fields[0].SendKeys(centreID);
            fields[1].SendKeys(centreID);
            fields[2].SendKeys(subjectID);
            fields[3].SendKeys(subjectID);
            //flick screen to get next question in view
            var touch = new RemoteTouchScreen(driver);
            GenericFunctions.Wait(1); 
            touch.Flick(500, -1);
            GenericFunctions.Wait(1); 
            touch.Flick(500, -1);
            GenericFunctions.Wait(1); 
            //select 'No' to replacement phone question
            driver.FindElementByName("No").Click();
            //select Next to confirm enrolment form details
            driver.FindElementByName("Next").Click();
            GenericFunctions.Wait(1); 

            //Capture the temporary PIN and SUBMIT
            orangutan.CaptureElementCollection("android.widget.TextView", driver);
            fields = driver.FindElementsByClassName("android.widget.TextView");
            sTempPIN = fields[6].Text;
            //check it captured it correctly
            if (sTempPIN.Length != 4)
            {
                //didn't capture it correctly - return false
                sTempPIN = "false";
            }
            //Console.WriteLine("sTempPIN = " + sTempPIN);
            driver.FindElementByName("Submit").Click();
            //Select Next in the enrolment confirmation screen
            Thread.Sleep(1000);

            //standroid.GetScreenShot(driver, "screenshot");

            driver.FindElementByName("Next").Click();
            //Logout of the site admin side
            Thread.Sleep(1000);

            //standroid.GetScreenShot(driver, "screenshot");

            driver.FindElementByName("Logout").Click();

            //standroid.GetScreenShot(driver, "screenshot");

            return sTempPIN;
        }

        public static void SpecialCompleteScreensUntilSubjectMenu(AndroidDriver driver)
        {
            //complete all screens until the subject is returned to the LIVE Subject Menu
            bool bDone = false;
            do
            {
                //make sure screen is loaded
                GenericFunctions.Wait(1);
                //if this is the LIVE Subject Menu then set training to complete and exit do
                if (IssubjectMenu(driver) == true)
                {
                    bDone = true;
                }

                else
                {
                    //training isn't complete yet - check if this screen is special - if it's special it's dealt with by amIspecial
                    if (amIspecial(driver) == false)
                    {
                        //this screen aint nufin special... go ape
                        orangutan.goBananas(driver);
                        //navigate out of the screen
                        orangutan.RightTurnClyde("android.widget.Button", navigationTryOrder(), driver);
                    }
                }
            }
            while (bDone == false);
        }

        public static bool IssubjectMenu(AndroidDriver driver)
        {
            if (driver.FindElementsByName("Subject Menu").Any())
            {
                if (driver.FindElementByName("Medications").GetAttribute("enabled") == "true")
                {
                    return true;
                }
               }
            return false;
        }

        public static bool IsElementPresent(string element, AndroidDriver driver)
        {
            try
            {
                driver.FindElement(By.Name(element));
                return true; // Success!
            }
            catch (NoSuchElementException)
            {
                //Console.WriteLine("Error:" + element + " not displayed");
                return false; // Do not throw exception
            }
        }
    


        public static bool amIspecial(AndroidDriver driver)
        {
            bool bamSpecial = false;
            var vCollection = driver.FindElementsByClassName("android.widget.EditText");
            //int randomInt;
            //string randomText;
            //string myText;
            Random random = new Random();
            //check to see if screen is special

            //am i the new VRC screen?
            if (driver.FindElementsByName("New VRC").Any())
            {
                //Console.WriteLine("SPECIAL - New VRC start");

                //standroid.GetScreenShot(driver, "screenshot");

                //New VRC screen - select Yes then Next
                driver.FindElementByName("Yes").Click();
                Thread.Sleep(500);
                driver.FindElementByName("Next").Click();
                Thread.Sleep(500);
                //Console.WriteLine("SPECIAL - New VRC completed with YES");
            }


            //am i the Injection Location screen?
            if (IsElementPresent("InjectionSiteLocationActivity", driver))
            //if (driver.FindElementsByName("InjectionSiteLocationActivity").Count() > 0)
            {
                ScreenInjectionSiteLocation(driver);
                bamSpecial = true;
                return bamSpecial;
            }

            //am i the temperature screen?
            //if (driver.FindElementsByName("Oral Temperature").Count() > 0)
            if (IsElementPresent("TemperatureEntryActivity", driver))
            {
                screenTemperature(driver);
                //return special
               return true;
            }

            //am i the Training Complete screen?
            if (driver.FindElementsByName("The training module is now complete.").Count() > 0)
            {
                SpecialTrainingCompleteStartNewVrc(driver);
                //return special
               
                return true;
            }

            //am i an ER screen?
            if (driver.FindElementsByName("ER including urgent health centers").Count() > 0 && driver.FindElementsByClassName("android.widget.CheckBox").Count() == 3)
            {
                screenER(driver);
                return  true;
               
            }

            //am i a LOG screen?
            if (driver.FindElementsByName("TAP TO ADD").Count() > 0)
            {
                screenLog(driver);
               return true;
            }

            //am i the VSC instructions screen?
            if (IsElementPresent("VscInstructionsActivity", driver))
            {
                screenVSC(driver);
                return true;
            }
            //am i the ISRd6 instructions screen?
            if (IsElementPresent("Isrd6InstructionsActivity", driver))
            {
                screenISRD6(driver);
               return true;
            }
            // If OnGoing is not required the diary will be completed at random and could still have some ongoing events
            //Check if OnGoing is required - OnGoing set in globals            
            if (Globals.Run.PsOnGoing == "Yes")
            {
                //am i the Ongoing events screen?
                if (driver.FindElementsByName("Ongoing events").Any())
                {
                    screenOnGoing(driver);
                   return true;
                }
            }


            //I'm not special :[
            return false;
        }

        public static void SpecialTrainingCompleteStartNewVrc(AndroidDriver driver)
        {
            //Console.WriteLine("SPECIAL - specialTrainingCompleteStartNewVRC");
            //select 'Yes' training understood
            driver.FindElementByName("Yes").Click();
            //Select 'Next'
            driver.FindElementByName("Next").Click();
            GenericFunctions.Wait(1);
            //Eneter Live PIN and confirm it
            var vCollection = driver.FindElementsByClassName("android.widget.EditText");
            for (int i = 0; i <= (vCollection.Count() - 1); i++)
            {
                vCollection[i].SendKeys("1111");
            }
            //Select 'Next' to set PIN
            driver.FindElementByName("Next").Click();
            GenericFunctions.Wait(1);
            //enter 'Yes' to New VRC screen
            driver.FindElementByName("Yes").Click();
            //Select 'Next'
            driver.FindElementByName("Next").Click();
            GenericFunctions.Wait(1);
            //select injection site location and confirm in review screen
            ScreenInjectionSiteLocation(driver);
            GenericFunctions.Wait(1);
            driver.FindElementByName("Submit").Click();
            GenericFunctions.Wait(1);
        }

        public static void screenER(AndroidDriver driver)
        {
            //Console.WriteLine("SPECIAL SCREEN - ER");
            //semi-randomly select check boxes
            int randomInt;
            Random random = new Random();
            var vCollection = driver.FindElementsByClassName("android.widget.CheckBox");
            if (driver.FindElementsByClassName("android.widget.CheckBox").Count() == 3)
            {
                randomInt = random.Next(0, 4);
                switch (randomInt)
                {
                    case 0:
                        //select NONE
                        vCollection[0].Click();
                        //Console.WriteLine("ER Screen:  NONE");
                        break;
                    case 1:
                        //select ER
                        vCollection[1].Click();
                        //Console.WriteLine("ER Screen:  ER");
                        break;
                    case 2:
                        //select HOSPITALISED
                        vCollection[2].Click();
                        //Console.WriteLine("ER Screen:  HOSPITALISED");
                        break;
                    case 3:
                        //select ER and HOSPITALISED
                        vCollection[1].Click();
                        vCollection[2].Click();
                        //Console.WriteLine("ER Screen:  ER and HOSPITALISED");
                        break;
                }
                //select NEXT and return special
                driver.FindElementByName("Next").Click();
            }
         }

        public static void ScreenInjectionSiteLocation(AndroidDriver driver)
        {
            //Console.WriteLine("SPECIAL SCREEN - Injection Site Location");
            //semi-randomly select a location
            int randomInt;
            int i;
            var random = new Random();
            var vCollection = driver.FindElementsByClassName("android.widget.ToggleButton");
            if (driver.FindElementsByClassName("android.widget.ToggleButton").Count() == 3)
            {
                randomInt = random.Next(0, 3);

                switch (randomInt)
                {
                    case 0:
                        //select RIGHT
                        vCollection[0].Click();
                        //Console.WriteLine("Inejection Site Location Screen:  Left arm");
                        break;
                    case 1:
                        //select LEFT
                        vCollection[1].Click();
                        //Console.WriteLine("Inejection Site Location Screen:  Right arm");
                        break;
                    case 2:
                        //select OTHER
                        vCollection[2].Click();
                        Thread.Sleep(500);
                        string randomInjection = orangutan.randomSentence(30);
                        driver.FindElementByClassName("android.widget.EditText").SendKeys(randomInjection);
                        //Console.WriteLine("Inejection Site Location Screen:  Other(" + randomInjection + ")");
                        break;
                }
                //select NEXT
                driver.FindElementByName("Next").Click();
            }
            else 
                if (driver.FindElementsByClassName("android.widget.ToggleButton").Count() == 6)
            {
                randomInt = random.Next(0, 6);
                
                for (i = 0; i < 2; i++)
                {
                    //set int to odd if randomInt is even on first select and vice versa
                    if (i > 0 && randomInt % 2 == 0)
                    {
                        //Generate odd number
                        randomInt = random.Next(0 / 2, 6 / 2) * 2 + 1;
                    }
                    else if (i > 0 && randomInt % 2 == 1)
                    {
                        //Generate even number
                        randomInt = random.Next(0 / 2, 6 / 2) * 2;
                    }
                    
                    switch (randomInt)
                    {
                        case 0:
                            //select Left Arm
                            vCollection[0].Click();
                            //Console.WriteLine("Inejection Site Location Screen:  Left arm");
                            break;
                        case 1:
                            //select Right Arm
                            vCollection[1].Click();
                            //Console.WriteLine("Inejection Site Location Screen:  Right arm");
                            break;
                        case 2:
                            //select Left Thigh
                            vCollection[2].Click();
                            //Console.WriteLine("Inejection Site Location Screen:  Left thigh");
                            break;
                        case 3:
                            //select Right Thigh
                            vCollection[3].Click();
                            //Console.WriteLine("Inejection Site Location Screen:  Right thigh");
                            break;
                        case 4:
                            //select Other 1
                            vCollection[4].Click();
                            string randomInjection1 = orangutan.randomSentence(30);
                            driver.FindElementByClassName("android.widget.EditText").SendKeys(randomInjection1);
                            //Console.WriteLine("Inejection Site Location Screen:  Other(" + randomInjection1 + ")");
                            break;
                        case 5:
                            //select Other 2
                            vCollection[5].Click();
                            Thread.Sleep(500);
                            string randomInjection2 = orangutan.randomSentence(30);
                            //Check the number of text boxes on the screen - if greater than 1 select the second and populate with random string
                            if (driver.FindElementsByClassName("android.widget.EditText").Count() > 1)
                            {
                                var vColl = driver.FindElementsByClassName("android.widget.EditText");
                                vColl[1].SendKeys(randomInjection2);
                            }
                            //if only one exists populate with random string
                            else
                            {
                                driver.FindElementByClassName("android.widget.EditText").SendKeys(randomInjection2);
                            }                            
                            //Console.WriteLine("Inejection Site Location Screen:  Other(" + randomInjection2 + ")");
                            break;
                    }

                }
                    //select NEXT
                    driver.FindElementByName("Next").Click();
            }
            else
            {
                //Console.WriteLine("FAIL @ Injection Site Location Screen:  Missmatch in Collection Count:  Count = " + driver.FindElementsByClassName("android.widget.ToggleButton").Count().ToString());
            }
        }

        public static void screenTemperature(AndroidDriver driver)
        {
            //Console.WriteLine("SPECIAL - Temperature start");
            var vCollection = driver.FindElementsByClassName("android.widget.EditText");
            var random = new Random();
            string myText;
            int randomInt = random.Next(95, 105);
                //get a random decimal place
                myText = randomInt + "." + random.Next(0, 9);
      
                driver.Navigate().Back();
                GenericFunctions.Wait(1);

                vCollection[0].SendKeys(myText);
            //Console.WriteLine("SPECIAL - Temperature = " + myText);

                driver.Navigate().Back();
                Thread.Sleep(1000);

                vCollection[1].SendKeys(myText);
            //Console.WriteLine("SPECIAL - Temperature = " + myText);
            
            //close the keyboard
            driver.Navigate().Back();
            GenericFunctions.Wait(1);
            //select 'Next'
            driver.FindElementByName("Next").Click();
            GenericFunctions.Wait(1);
            driver.FindElementByName("Next").Click();
            GenericFunctions.Wait(1);
        }

        public static void screenLog(AndroidDriver driver)
        {
            //work out which log screen this is
            var vCollection = driver.FindElementsByClassName("android.widget.TextView");
            string screenName = vCollection[2].Text.ToString();
            switch (screenName)
            {
                case "Other Complaints or Illnesses:":
                    //Console.WriteLine("LOG SCREEN:  OCI");
                    driver.FindElementByName("TAP TO ADD").Click();
                    orangutan.goBananas(driver);
                    driver.FindElementByName("Next").Click();
                    merckVRC.screenER(driver);
                    driver.FindElementByName("Next").Click();
                    break;
                case "Non-study vaccinations:":
                    //Console.WriteLine("LOG SCREEN:  NSV");
                    driver.FindElementByName("TAP TO ADD").Click();
                    orangutan.goBananas(driver);
                    driver.FindElementByName("Next").Click();
                    driver.FindElementByName("Next").Click();
                    break;
                case "Medications":
                    //Console.WriteLine("LOG SCREEN:  MEDICATIONS");
                    driver.FindElementByName("TAP TO ADD").Click();
                    orangutan.goBananas(driver);
                    driver.FindElementByName("Next").Click();
                    driver.FindElementByName("Next").Click();
                    break;
                case "Injection Site Complaints":
                    //Console.WriteLine("LOG SCREEN:  ISCd6");
                    driver.FindElementByName("TAP TO ADD").Click();
                    orangutan.goBananas(driver);
                    merckVRC.screenER(driver);
                    driver.FindElementByName("Next").Click();
                    driver.FindElementByName("Next").Click();
                    break;
                case "Vaccine Specific Complaints:":
                    //Console.WriteLine("LOG SCREEN:  VSR");
                    driver.FindElementByName("TAP TO ADD").Click();
                    orangutan.goBananas(driver);
                    driver.FindElementByName("Next").Click();
                    orangutan.goBananas(driver);
                    // scroll down to ensure
                    RemoteTouchScreen touch = new RemoteTouchScreen(driver);
                    do {
                            Thread.Sleep(500);
                            touch.Flick(500, -1);
                            Thread.Sleep(500);
                    } while(IsElementPresent("Next", driver) == false);
                    driver.FindElementByName("Next").Click();
                    orangutan.goBananas(driver);
                    driver.FindElementByName("Next").Click();
                    orangutan.goBananas(driver);
                    // scroll down to ensure
                    //RemoteTouchScreen touch1 = new RemoteTouchScreen(driver);
                    do {
                            Thread.Sleep(500);
                            touch.Flick(500, -1);
                            Thread.Sleep(500);
                    } while(IsElementPresent("Next", driver) == false);
                    driver.FindElementByName("Next").Click();
                    orangutan.goBananas(driver);
                    driver.FindElementByName("Next").Click();
                    orangutan.goBananas(driver);
                    driver.FindElementByName("Next").Click();
                    screenVSC7(driver);
                    driver.FindElementByName("Next").Click(); 
                    break;
            }

        }

        public static void screenVSC(AndroidDriver driver)
        {
            //scroll to ensure next button is present
            RemoteTouchScreen touch = new RemoteTouchScreen(driver);
            do
            {
                GenericFunctions.Wait(0.5);
                touch.Flick(500, -1);
                Thread.Sleep(500);
            } while (IsElementPresent("Next", driver) == false);

            //select NEXT
            driver.FindElementByName("Next").Click();
        }

        public static void screenISRD6(AndroidDriver driver)
        {
            //scroll to ensure next button is present
            RemoteTouchScreen touch = new RemoteTouchScreen(driver);
            do
            {
                Thread.Sleep(500);
                touch.Flick(500, -1);
                Thread.Sleep(500);
            } while (IsElementPresent("Next", driver) == false);

            //select NEXT
            driver.FindElementByName("Next").Click();
        }

        public static void screenVSC7(AndroidDriver driver)
        {
            int i;
            int i2;
            // create collection of Radio Groups
            var vCollection = driver.FindElementsByClassName("android.widget.RadioGroup");
            var vCollRadio = driver.FindElementsByClassName("android.widget.RadioButton");
            //Loop through Radio Group collection
            for (i = 0; i < (vCollection.Count()); i++)
            {
                vCollRadio = vCollection[i].FindElements(By.ClassName("android.widget.RadioButton"));
                //Check the NO option is selected - If no select it
                for (i2 = 0; i2 < (vCollRadio.Count()); i2++)
                {
                    if (vCollRadio[i2].Text == "No" && vCollRadio[i2].GetAttribute("checked").ToString() == "false")
                    {
                        vCollRadio[i2].Click();
                    }
                }
            }
            //Set up touch screen scrolling
            RemoteTouchScreen touch = new RemoteTouchScreen(driver);
            do
            {
                //scroll to ensure next radio group is present
                touch.Flick(150, -1);
                Thread.Sleep(500);
                // create collection of Radio Groups
                vCollection = driver.FindElementsByClassName("android.widget.RadioGroup");
                //Loop through Radio Group collection
                for (i = 0; i < (vCollection.Count()); i++)
                {
                    vCollRadio = vCollection[i].FindElements(By.ClassName("android.widget.RadioButton"));
                    //Check the NO option is selected - If no select it
                    for (i2 = 0; i2 < (vCollRadio.Count()); i2++)
                    {
                        if (vCollRadio[i2].Text == "No" && vCollRadio[i2].GetAttribute("checked").ToString() == "false")
                        {
                            vCollRadio[i2].Click();
                        }
                    }
                    
                }
            } while (IsElementPresent("Next", driver) == false);

            //select NEXT
            driver.FindElementByName("Next").Click();   

        }

        //Function to ensure unresolved events remain unresolved for the duration of the diary entry
        public static void screenOnGoing(AndroidDriver driver)
        {
            //work out which log screen this is
            var vCollection = driver.FindElementsByClassName("android.widget.TextView");            
            string screenName = vCollection[2].GetAttribute("name").ToString();
            switch (screenName)
            {
                case "UnresolvedIsrSiteOneSwellingActivity":
                    driver.FindElementByName("Yes").Click();
                    driver.FindElementByName("Next").Click();
                    break;
                case "UnresolvedIsrSiteOneRednessActivity":
                    driver.FindElementByName("Yes").Click();
                    driver.FindElementByName("Next").Click();
                    break;
                case "UnresolvedIsrSiteOnePainActivity":
                    driver.FindElementByName("Yes").Click();
                    driver.FindElementByName("Next").Click();
                    break;
                case "UnresolvedIsrSiteOneOtherUpdateActivity":
                    driver.FindElementByName("Yes").Click();
                    driver.FindElementByName("Next").Click();
                    break;
                case "UnresolvedIsrSiteTwoSwellingActivity":
                    driver.FindElementByName("Yes").Click();
                    driver.FindElementByName("Next").Click();
                    break;
                case "UnresolvedIsrSiteTwoRednessActivity":
                    driver.FindElementByName("Yes").Click();
                    driver.FindElementByName("Next").Click();
                    break;
                case "UnresolvedIsrSiteTwoPainActivity":
                    driver.FindElementByName("Yes").Click();
                    driver.FindElementByName("Next").Click();
                    break;
                case "UnresolvedIsrSiteTwoOtherUpdateActivity":
                    driver.FindElementByName("Yes").Click();
                    driver.FindElementByName("Next").Click();
                    break;
                case "UnresolvedIsrInstructionsActivity":
                    driver.FindElementByName("Next").Click();
                    break;
            }
        }

        public static List<string> navigationTryOrder()
        {
            List<string> navigateOrder = new List<string>();
            navigateOrder.Add("Next");
            navigateOrder.Add("Submit");
            navigateOrder.Add("Close");
            navigateOrder.Add("Daily Questions");
            navigateOrder.Add("?");
            return navigateOrder;
        }


    }
}
