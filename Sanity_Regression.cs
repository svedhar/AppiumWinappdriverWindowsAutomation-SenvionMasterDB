using OpenQA.Selenium.Appium.Enums;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Appium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AventStack.ExtentReports;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium;
using System.Diagnostics;
using System.Xml.Linq;
using System.Diagnostics.Metrics;
using Microsoft.AspNetCore.Server.Kestrel;
using WindowsInput;

namespace SanityTest
{
    [TestFixture]
    public class Regression_Sanity : ReportsManager
    {
        static WindowsDriver<WindowsElement> driver;
        private ReportsManager reportHelperS;
        // private ExtentTest test;

        [OneTimeSetUp]
        public void Setup()
        {
            System.Diagnostics.Debug.WriteLine("****** Setp Method called ******");
            reportHelperS = new ReportsManager();
            reportHelperS.ExtentManager();
            var appiumOptions = new AppiumOptions();
            // appiumOptions.AddAdditionalCapability(MobileCapabilityType.App, @"C:\MasterDB\MasterDB.exe");
            appiumOptions.AddAdditionalCapability(MobileCapabilityType.App, @"C:\RE Technologies\MasterDB\MasterDB.exe");
            appiumOptions.AddAdditionalCapability(MobileCapabilityType.DeviceName, "WindowsPC");
            driver = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appiumOptions);
            // reportHelper.CreateTest(TestContext.CurrentContext.Test.Name);
        }

        [Test, Order(1)]
        public void LoginPage()
        {
            System.Diagnostics.Debug.WriteLine("*** Test Login Page called ***");
            try
            {
                //Invalid Credentials
                driver.FindElementByAccessibilityId("txtUserName").SendKeys("vedha");
                driver.FindElementByAccessibilityId("txtPassword").SendKeys("123456");
                driver.FindElementByAccessibilityId("btnLogin").Click();
                var okButn = driver.FindElementByName("OK");
                okButn.Click();
                //Empty fields
                driver.FindElementByAccessibilityId("btnLogin").Click();
                //CASE SENSITIVE and invalid password
                driver.FindElementByAccessibilityId("txtUserName").SendKeys("VEDHA.RAJESHWARI.EXT@SENVION.IN");
                driver.FindElementByAccessibilityId("txtPassword").SendKeys("123");
                driver.FindElementByAccessibilityId("btnLogin").Click();
                var okButn1 = driver.FindElementByName("OK");
                okButn1.Click();
                //Valid and click on Eye icon AutomationId="EyeIcon" 
                driver.FindElementByAccessibilityId("txtUserName").SendKeys("vedha.rajeshwari.ext@senvion.in");
                driver.FindElementByAccessibilityId("txtPassword").SendKeys("123456");
                driver.FindElementByAccessibilityId("EyeIcon").Click();
                driver.FindElementByAccessibilityId("btnLogin").Click();
                // reportHelper.LogTestStepPass(Status.Pass, "Test Pass");
                // Assert.Pass();
                reportHelperS.LogTestStep(Status.Pass, "Login Step");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"**** Parameters TCS fail executing****" + ex.Message);
                reportHelperS.LogTestStep(Status.Fail, "Login Step Failed");
                // reportHelperS.LogTestStepError(Status.Fail, + ex.Message);
                reportHelperS.LogScreenshot(Status.Fail, "Login Step Failed", "Screenshot of failure", driver);
            }
        }

        [Test, Order(2)]
        [Ignore("Ignoring for quick test")]
        public void ADD_EDIT_DELETE_Parameter_MainController()
        {
            try
            {
                Thread.Sleep(5000);
                var windowHandles = driver.WindowHandles;
                for (int attempt = 0; attempt < 5 && windowHandles.Count > 1; attempt++)
                {
                    Thread.Sleep(TimeSpan.FromSeconds(1));
                    windowHandles = driver.WindowHandles;
                }
                if (windowHandles.Count >= 0)
                {
                    driver.SwitchTo().Window(windowHandles[0]);
                }
                //2XM
                WindowsElement twoXMButton = driver.FindElementByName("2XM");
                twoXMButton.Click();
                twoXMButton.SendKeys(Keys.Down);
                Thread.Sleep(5000);
                twoXMButton.SendKeys(Keys.Enter + Keys.Enter);
               // Thread.Sleep(5000);
               // driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
               // driver.FindElementByName("View Parameters").Click();
               // Main Controller
               // var radioFirst1 = driver.FindElementByName("Main Controller");
               // System.Diagnostics.Debug.WriteLine($"***** Value of radio First: {radioFirst1.Selected}");
                //  radioFirst1.Click();
                Thread.Sleep(17000);
                var btnAdd = driver.FindElementByName("Add");
                btnAdd.Click();
                //Thread.Sleep(1000);
                // driver.FindElementByName("Open").Click();
                var combo = driver.FindElementByTagName("ComboBox");
                var open = combo.FindElementByName("Open");
                var listItems = combo.FindElementsByTagName("ListItem");
                Debug.WriteLine($"Before: Number of list items found: {listItems.Count}");
                combo.SendKeys(Keys.Down + Keys.Down);
                open.Click();
                listItems = combo.FindElementsByTagName("ListItem");
                Debug.WriteLine($"After: Number of list items found: {listItems.Count}");
                // maybe check number of elements in combo 
                //Assert.AreEqual(6, listItems.Count, "Combo box doesn't contain expected number of elements.");
                foreach (var comboKid in listItems)
                {
                    if (comboKid.Text == "Blade Sensors")
                    {
                        WebDriverWait wdv = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                        wdv.Until(x => comboKid.Displayed);
                        comboKid.Click();
                    }
                }
                var combo2 = driver.FindElementByAccessibilityId("ddlSystems");
                //  var combo2 = driver.FindElementByClassName("WindowsForms10.COMBOBOX.app.0.bb8560_r8_ad1");
                // var combo2 = driver.FindElementByTagName("ComboBox");
                var open2 = combo.FindElementByName("Open");
                var listItems2 = combo.FindElementsByTagName("ListItem");
                Debug.WriteLine($"Before: Number of list items found: {listItems2.Count}");
                combo2.SendKeys(Keys.Down + Keys.Down + Keys.Down);
                open2.Click();
                listItems2 = combo.FindElementsByTagName("ListItem");
                Debug.WriteLine($"After: Number of list items found: {listItems2.Count}");
                foreach (var comboKid in listItems2)
                {
                    if (comboKid.Text == "Gain correction")
                    {
                        WebDriverWait wdv = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                        wdv.Until(x => comboKid.Displayed);
                        comboKid.Click();
                    }
                }
                driver.FindElementByAccessibilityId("txtParameterName").SendKeys("System_Main");
                driver.FindElementByAccessibilityId("txtSystemName").SendKeys("Config.SessionTime_Main_Control");
                // Thread.Sleep(2000);
                driver.FindElementByAccessibilityId("txtMinValue").SendKeys("56");
                //  Thread.Sleep(2000);
                driver.FindElementByAccessibilityId("txtMaxValue").SendKeys("123");
                // Thread.Sleep(2000);
                driver.FindElementByAccessibilityId("txtValue").SendKeys("67");
                // Thread.Sleep(2000);
                driver.FindElementByAccessibilityId("txtUnit").SendKeys("h");
               // driver.FindElementByName("kW").Click();
                driver.FindElementByAccessibilityId("txtDescription").SendKeys("Controls Develop services");
                //  driver.FindElementByAccessibilityId("txtDescription").SendKeys("NULL Develop services with added value using your online access,search for correct information,check the contact details of an applicant,validate all of the data related to incoming and outgoing calls,etc.\r\nOur added value\r\nInfobel is a global platform with extensive search criteria allowing companies to search for individuals and professionals in a speedy,flexible and automated manner within an environment free of advertising.evelop services with added value using your online access,search for correct information,check the contact details of an applicant,validate all of the data related to incoming and outgoing calls,etc.\r\nOur added value\r\nInfobel is a global platform with extensive search criteria allowing companies to search for individuals and professionals in a speedy,flexible and automated");
                driver.FindElementByName("Add").Click();
                driver.FindElementByName("Yes").Click();
                driver.FindElementByName("OK").Click();
                //  driver.FindElementByAccessibilityId("txtMaxValue").Clear();
                //  driver.FindElementByAccessibilityId("txtMaxValue").SendKeys("345");
                //  driver.FindElementByAccessibilityId("txtValue").Clear();
                // driver.FindElementByAccessibilityId("txtValue").SendKeys("66");
                //   driver.FindElementByName("Add").Click();
                //   driver.FindElementByName("Yes").Click();
                //   driver.FindElementByName("OK").Click();
                Actions actsTree = new Actions(driver);
                var nodeWorld = driver.FindElementByName("Component Version Row 0");
                actsTree.MoveToElement(nodeWorld);
                actsTree.DoubleClick();
                actsTree.Perform();
               //  driver.FindElementByName("View Audit Logs").Click();
               // driver.FindElementByName("Close").Click();
                Thread.Sleep(1000);
                driver.FindElementByName("Close").Click();
                //EDIT
                var btnEdit1 = driver.FindElementByName("Edit");
                btnEdit1.Click();
                driver.FindElementByName("OK").Click();
                driver.FindElementByName("Component Version Row 0").Click();
                var btnEdit = driver.FindElementByName("Edit");
                btnEdit.Click();
                driver.FindElementByAccessibilityId("txtUnit").Clear();
                driver.FindElementByAccessibilityId("txtUnit").SendKeys("MW");
                driver.FindElementByName("Update").Click();
                // Thread.Sleep(2000);
                driver.FindElementByName("OK").Click();
                //  Thread.Sleep(2000);
                driver.FindElementByName("OK").Click();

                //DELETE

                driver.FindElementByName("Component Version Row 0").Click();

                var btnDel = driver.FindElementByName("Delete");
                btnDel.Click();

                driver.FindElementByName("Yes").Click();

                driver.FindElementByName("OK").Click();

                //LOG
                reportHelperS.LogTestStep(Status.Pass, "ADD, EDIT and DELETE Parameters - Main Controller");

            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"*****Exception in ADD Parameter tcs****" + driver + ex.Message);
                reportHelperS.LogTestStep(Status.Fail, "ADD, EDIT and DELETE Parameters - Main Controller Test Failed           " + ex.Message);
                reportHelperS.LogScreenshot(Status.Fail, "ADD, EDIT and DELETE Parameters - Main Controller Test Failed", "Screenshot of failure", driver);
            }
        }
        [Test, Order(3)]
       [Ignore("Ignoring for quick test")]
        public void ADD_EDIT_DELETE_Parameters_HubController()
        {
            try
            {
                // reportHelper.LogTestStep(Status.Pass, "Step 1: Add Parameter(Hub Controller)- MasterDB", driver);
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                //HUB Controller
                var radioFirst = driver.FindElementByName("Hub Controller");
                System.Diagnostics.Debug.WriteLine($"***** Value of radio First: {radioFirst.Selected}");
                radioFirst.Click();
                var btnAdd = driver.FindElementByName("Add");
                btnAdd.Click();
                var combofg = driver.FindElementByAccessibilityId("ddlFunctionalGroup");
                var listItems = combofg.FindElementsByTagName("ListItem");
                Debug.WriteLine($"Before: Number of list items found: {listItems.Count}");
                combofg.SendKeys(Keys.Down + Keys.Down + Keys.Down);
                combofg.Click();
                listItems = combofg.FindElementsByTagName("ListItem");
                Debug.WriteLine($"After: Number of list items found: {listItems.Count}");
                foreach (var comboKid in listItems)
                {
                    if (comboKid.Text == "Blades")
                    {
                        WebDriverWait wdv = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                        wdv.Until(x => comboKid.Displayed);
                        comboKid.Click();
                    }
                }
                Thread.Sleep(2000);
                var combo2 = driver.FindElementByAccessibilityId("ddlSystems");
               // var open2 = combo2.FindElementByName("Open");
              //  var listItems2 = combo2.FindElementsByTagName("ListItem");
               // Debug.WriteLine($"Before: Number of list items found: {listItems2.Count}");
                combo2.SendKeys(Keys.Down + Keys.Down + Keys.Enter);
               // open2.Click();
               // listItems2 = combo2.FindElementsByTagName("ListItem");
              //  Debug.WriteLine($"After: Number of list items found: {listItems2.Count}");
               // foreach (var comboKid in listItems2)
               // {
                 //   if (comboKid.Text == "Digital output")
                  //  {
                   //     WebDriverWait wdv = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                   //     wdv.Until(x => comboKid.Displayed);

                  //      comboKid.Click();
                 //   }
               // }
                driver.FindElementByAccessibilityId("txtParameterName").SendKeys("System_HUBCv");
                // Thread.Sleep(2000);
                driver.FindElementByAccessibilityId("txtSystemName").SendKeys("Config.SessionTime_HUB_v");
                // Thread.Sleep(2000);
                driver.FindElementByAccessibilityId("txtMinValue").SendKeys("4566");
                //  Thread.Sleep(2000);
                driver.FindElementByAccessibilityId("txtMaxValue").SendKeys("5675");
                // Thread.Sleep(2000);
                driver.FindElementByAccessibilityId("txtValue").SendKeys("5434");
                // Thread.Sleep(2000);
                driver.FindElementByAccessibilityId("txtUnit").SendKeys("kW");
                // driver.FindElementByName("kW").Click();
                Thread.Sleep(2000);
                driver.FindElementByAccessibilityId("txtDescription").SendKeys("Develop");
                //  driver.FindElementByAccessibilityId("txtDescription").SendKeys("Develop services with added value using your online access,search for correct information,check the contact details of an applicant,validate all of the data related to incoming and outgoing calls,etc.\r\nOur added value\r\nInfobel is a global platform with extensive search criteria allowing companies to search for individuals and professionals in a speedy,flexible and automated manner within an environment free of advertising.evelop services with added value using your online access,search for correct information,check the contact details of an applicant,validate all of the data related to incoming and outgoing calls,etc.\r\nOur added value\r\nInfobel is a global platform with extensive search criteria allowing companies to search for individuals and professionals in a speedy,flexible and automated");
                Thread.Sleep(2000);
                driver.FindElementByName("Add").Click();
                //  Thread.Sleep(2000);
                //  driver.FindElementByName("Yes").Click();
                // Thread.Sleep(2000);
                // driver.FindElementByName("OK").Click();
                //  driver.FindElementByAccessibilityId("txtValue").Clear();
                //  driver.FindElementByAccessibilityId("txtValue").SendKeys("5434");
                //  driver.FindElementByName("Add").Click();
                driver.FindElementByName("Yes").Click();
                driver.FindElementByName("OK").Click();
                //View
                Actions actsTree = new Actions(driver);
                var nodeWorld = driver.FindElementByName("Component Version Row 0");
                actsTree.MoveToElement(nodeWorld);
                actsTree.DoubleClick();
                actsTree.Perform();
               // driver.FindElementByName("View Audit Logs").Click();
               // driver.FindElementByName("Close").Click();
                // Thread.Sleep(2000);
                // driver.FindElementByName("OK").Click();
                Thread.Sleep(1000);
                driver.FindElementByName("Close").Click();

                //EDIT
                var btnEdit1 = driver.FindElementByName("Edit");
                btnEdit1.Click();
                driver.FindElementByName("OK").Click();
                driver.FindElementByName("Component Version Row 0").Click();
                var btnEdit = driver.FindElementByName("Edit");
                btnEdit.Click();
                driver.FindElementByAccessibilityId("txtUnit").Clear();
                driver.FindElementByAccessibilityId("txtUnit").SendKeys("MWt");
                driver.FindElementByName("Update").Click();
                driver.FindElementByName("OK").Click();
                driver.FindElementByName("OK").Click();

                //DELETE
                driver.FindElementByName("Component Version Row 0").Click();
                var btnDel = driver.FindElementByName("Delete");
                btnDel.Click();
                driver.FindElementByName("Yes").Click();
                driver.FindElementByName("OK").Click();

                //LOG
                reportHelperS.LogTestStep(Status.Pass, "ADD, EDIT and DELETE Parameters - Hub Controller");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"**** Parameters TCS fail executing****" + ex.Message);
                reportHelperS.LogTestStep(Status.Fail, "ADD, EDIT and DELETE Parameters - Hub Controller              " + ex.Message);
                reportHelperS.LogScreenshot(Status.Fail, "ADD, EDIT and DELETE Parameters - Hub Controller", "Screenshot of failure", driver);
            }
        }

        [Test, Order(4)]
       [Ignore("Ignoring for quick test")]
        public void Change_Requests()
        {
            try
            {
                Thread.Sleep(5000);
                var windowHandles = driver.WindowHandles;
                for (int attempt = 0; attempt < 5 && windowHandles.Count > 1; attempt++)
                {
                    Thread.Sleep(TimeSpan.FromSeconds(1));
                    windowHandles = driver.WindowHandles;
                }

                if (windowHandles.Count >= 0)
                {
                    driver.SwitchTo().Window(windowHandles[0]);
                }
                //2XM
                WindowsElement twoXMButton = driver.FindElementByName("2XM");
                twoXMButton.Click();
                twoXMButton.SendKeys(Keys.Down);
                Thread.Sleep(5000);
                twoXMButton.SendKeys(Keys.ArrowRight + Keys.Down);
                Thread.Sleep(5000);
                twoXMButton.SendKeys(Keys.Enter);


              //  WindowsElement Button = driver.FindElementByName("Parameters");
              //  Button.Click();
               // Button.SendKeys(Keys.Down + Keys.Down + Keys.Enter);
                // Button.Click();

                Thread.Sleep(5000);

                // driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                // driver.FindElementByName("Change Requests").Click();
                var radioFirst2 = driver.FindElementByName("Hub Controller");
                radioFirst2.Click();
                var radioFirst1 = driver.FindElementByName("Main Controller");
                radioFirst1.Click();
                driver.FindElementByName("Component Version Row 0").Click();
                driver.FindElementByName("Open").Click();
                driver.FindElementByName("Close").Click();
                driver.FindElementByName("Yes").Click();
                reportHelperS.LogTestStep(Status.Pass, "Change Requests - Parameters");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"**** Parameters TCS fail executing****" + ex.Message);
                reportHelperS.LogTestStep(Status.Fail, "Change Requests - Parameters                 " + ex.Message);
                reportHelperS.LogScreenshot(Status.Fail, "Change Requests - Parameters", "Screenshot of failure", driver);
            }
        }
        [Test, Order(5)]
      //  [Ignore("Ignoring for quick test")]
        public void Main_Controller_CRUD_Status_Code()
        {
            try
            {
                Thread.Sleep(5000);
                var windowHandles = driver.WindowHandles;
                for (int attempt = 0; attempt < 5 && windowHandles.Count > 1; attempt++)
                {
                    Thread.Sleep(TimeSpan.FromSeconds(1));
                    windowHandles = driver.WindowHandles;
                }

                if (windowHandles.Count >= 0)
                {
                    driver.SwitchTo().Window(windowHandles[0]);
                }

                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
               
                //2XM
                WindowsElement twoXMButton = driver.FindElementByName("2XM");
                twoXMButton.Click();
                twoXMButton.SendKeys(Keys.Down + Keys.Down + Keys.Enter);
                Thread.Sleep(5000);
              //  driver.FindElementByName("Status Code").Click();
                driver.FindElementByName("Add").Click();
                // driver.FindElementByAccessibilityId("textStatusCodeNo").SendKeys("00020122");
                driver.FindElementByAccessibilityId("textStatusCodeNo").SendKeys("01");
                //  driver.FindElementByAccessibilityId("ddlFunctionalGroup").Click();
                var FG_drpdwn2 = driver.FindElementByAccessibilityId("ddlFunctionalGroup");
                FG_drpdwn2.Click();
                FG_drpdwn2.SendKeys(Keys.Down + Keys.Down + Keys.Down + Keys.Down);
                FG_drpdwn2.Click();
                //  driver.FindElementByName("Brake").Click();
                //  driver.FindElementByAccessibilityId("ddlTypes").Click();
                //  driver.FindElementByName("Info").Click();
                var Type_drpdwn2 = driver.FindElementByAccessibilityId("ddlTypes");
                Type_drpdwn2.Click();
                Type_drpdwn2.SendKeys(Keys.Down + Keys.Down);
                Type_drpdwn2.Click();

                driver.FindElementByAccessibilityId("txtAvailabilityGroup").SendKeys("31");
                driver.FindElementByAccessibilityId("txtSetDelay").SendKeys("0ms");
                driver.FindElementByAccessibilityId("txtResetDelay").SendKeys("10ms");
                driver.FindElementByAccessibilityId("textDescription").SendKeys("System");
                driver.FindElementByAccessibilityId("txtDevelopment").SendKeys("10ms");
                driver.FindElementByAccessibilityId("txtSales").SendKeys("System");
                driver.FindElementByAccessibilityId("txtTCC").SendKeys("10ms");
                driver.FindElementByAccessibilityId("txtServices").SendKeys("System");
                driver.FindElementByAccessibilityId("txtTurbineOperator").SendKeys("10ms");
                driver.FindElementByAccessibilityId("txtGridOperator").SendKeys("System");
                driver.FindElementByAccessibilityId("txtServicePartner").SendKeys("10ms");
                driver.FindElementByAccessibilityId("txtCustomer").SendKeys("System");
                driver.FindElementByAccessibilityId("txtCause").SendKeys("Add_Status_Code_MainController");


                /*    WindowsElement elementToScroll = driver.FindElementByAccessibilityId("txtDirective"); // Replace with your element's name
                    Actions actions = new Actions(driver);
                    actions.MoveToElement(elementToScroll);
                    actions.SendKeys(Keys.PageDown + Keys.PageDown);
                    actions.Perform();

                    WindowsElement elementToScroll2 = driver.FindElementByAccessibilityId("btnAdd"); // Replace with your element's name
                    Actions actions2 = new Actions(driver);
                    actions2.MoveToElement(elementToScroll);
                    actions2.SendKeys(Keys.PageDown + Keys.PageDown);
                    actions2.Perform();
                */

                //   driver.FindElementByAccessibilityId("DownButton").Click();

                driver.FindElementByAccessibilityId("txtDirective").SendKeys("System Add Status Code MainController");
                //" Name="Position" AutomationId="ScrollbarThumb" //Name="Line down" AutomationId="DownButton" 45

                InputSimulator simulator = new InputSimulator();
                simulator.Mouse.VerticalScroll(-5);
                //     simulator.Mouse.HorizontalScroll(5);

                driver.FindElementByAccessibilityId("btnAdd").Click();
                driver.FindElementByName("OK").Click();
                driver.FindElementByName("OK").Click();

                //EDIT
                driver.FindElementByName("Status Code Number Row 1").Click();
                driver.FindElementByName("Edit").Click();
                //driver.FindElementByName("View Audit Logs").Click();
                //driver.FindElementByName("Close").Click();
                driver.FindElementByAccessibilityId("textDescription").Clear();
                driver.FindElementByAccessibilityId("textDescription").SendKeys("StatusCode");

                // WindowsElement elementToScroll3 = driver.FindElementByAccessibilityId("txtDirective"); // Replace with your element's name
                //  Actions actions3 = new Actions(driver);
                //  actions3.MoveToElement(elementToScroll3);
                // actions3.SendKeys(Keys.PageDown + Keys.PageDown);
                //actions3.Perform();
                // driver.FindElementByAccessibilityId("txtDirective").Clear();
                //  driver.FindElementByAccessibilityId("txtDirective").SendKeys("StatusCode directive");

                InputSimulator simulator1 = new InputSimulator();
                simulator1.Mouse.VerticalScroll(-15);

                Thread.Sleep(4000);
                driver.FindElementByName("Update").Click();
                driver.FindElementByName("OK").Click();
                driver.FindElementByName("OK").Click();

                //DELETE
                driver.FindElementByName("Status Code Number Row 1").Click();
                driver.FindElementByName("Delete").Click();
                driver.FindElementByName("Yes").Click();
                driver.FindElementByName("OK").Click();

                reportHelperS.LogTestStep(Status.Pass, "Add,Edit and Delete Status Code - Main Controller");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"**** Parameters TCS fail executing****" + ex.Message);
                reportHelperS.LogTestStep(Status.Fail, "Add,Edit and Delete Status Code - Main Controller                      " + ex.Message);
                reportHelperS.LogScreenshot(Status.Fail, "Add,Edit and Delete Status Code - Main Controller", "Screenshot of failure", driver);
            }
        }

        [Test, Order(6)]
        //[Ignore("Ignoring for quick test")]
        public void CRUD_Status_Code_HubController()
        {
            try
            {
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                var radioFirst2 = driver.FindElementByName("Hub Controller");
                radioFirst2.Click();

                driver.FindElementByName("Add").Click();
                //driver.FindElementByAccessibilityId("textStatusCodeNo").SendKeys("0002012289");
                driver.FindElementByAccessibilityId("textStatusCodeNo").SendKeys("1");
                // driver.FindElementByAccessibilityId("ddlFunctionalGroup").Click();
                var FG_drpdwn2 = driver.FindElementByAccessibilityId("ddlFunctionalGroup");
                FG_drpdwn2.Click();
                FG_drpdwn2.SendKeys(Keys.Down + Keys.Down + Keys.Down + Keys.Down);
                FG_drpdwn2.Click();
                //  driver.FindElementByName("Brake").Click();
                //  driver.FindElementByAccessibilityId("ddlTypes").Click();
                //  driver.FindElementByName("Info").Click();
                var Type_drpdwn2 = driver.FindElementByAccessibilityId("ddlTypes");
                Type_drpdwn2.Click();
                Type_drpdwn2.SendKeys(Keys.Down + Keys.Down);
                Type_drpdwn2.Click();
                driver.FindElementByAccessibilityId("txtAvailabilityGroup").SendKeys("31");
                driver.FindElementByAccessibilityId("txtSetDelay").SendKeys("0ms");
                driver.FindElementByAccessibilityId("txtResetDelay").SendKeys("10ms");
                driver.FindElementByAccessibilityId("textDescription").SendKeys("System");
                driver.FindElementByAccessibilityId("txtDevelopment").SendKeys("10ms");
                driver.FindElementByAccessibilityId("txtSales").SendKeys("System");
                driver.FindElementByAccessibilityId("txtTCC").SendKeys("10ms");
                driver.FindElementByAccessibilityId("txtServices").SendKeys("System");
                driver.FindElementByAccessibilityId("txtTurbineOperatorPackage").SendKeys("10ms");
                driver.FindElementByAccessibilityId("txtGridOperator").SendKeys("System");
                driver.FindElementByAccessibilityId("txtServicePartner").SendKeys("10ms");
                driver.FindElementByAccessibilityId("txtCustomer").SendKeys("System");
                driver.FindElementByAccessibilityId("txtCause").SendKeys("Add Status Code HubController");
                driver.FindElementByAccessibilityId("txtDirective").SendKeys("System Add Status Code HubController");
                InputSimulator simulator = new InputSimulator();
                simulator.Mouse.VerticalScroll(-5);
                driver.FindElementByAccessibilityId("btnAdd").Click();
                driver.FindElementByName("OK").Click();
                driver.FindElementByName("OK").Click();

                //EDIT
                driver.FindElementByName("Status Code Number Row 1").Click();
                driver.FindElementByName("Edit").Click();
                // driver.FindElementByName("View Audit Logs").Click();
                //  driver.FindElementByName("Close").Click();
                driver.FindElementByAccessibilityId("textDescription").Clear();
                driver.FindElementByAccessibilityId("textDescription").SendKeys("StatusCode - CURD operations");
                InputSimulator simulator1 = new InputSimulator();
                simulator1.Mouse.VerticalScroll(-19);
                Thread.Sleep(2000);
                driver.FindElementByName("Update").Click();
                driver.FindElementByName("OK").Click();
                driver.FindElementByName("OK").Click();

                //DELETE
                driver.FindElementByName("Status Code Number Row 1").Click();
                driver.FindElementByName("Delete").Click();
                driver.FindElementByName("Yes").Click();
                driver.FindElementByName("OK").Click();
                reportHelperS.LogTestStep(Status.Pass, "Add,Edit and Delete Status Code - Hub Controller");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"**** Parameters TCS fail executing****" + ex.Message);
                reportHelperS.LogTestStep(Status.Fail, "Add,Edit and Delete Status Code - Hub Controller             " + ex.Message);
                reportHelperS.LogScreenshot(Status.Fail, "Add,Edit and Delete Status Code - Hub Controller", "Screenshot of failure", driver);
            }
        }
        [Test, Order(7)]
      //  [Ignore("Ignoring for quick test")]
        public void StatusCode_Import_Export_SoftwareVersion_Filter_Type_Search_Clear()
        {
            try
            {
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                var SW_drpdwn2 = driver.FindElementByAccessibilityId("ddlVersion");
                SW_drpdwn2.Click();
                SW_drpdwn2.SendKeys(Keys.Down + Keys.Down + Keys.Down);
                SW_drpdwn2.Click();
                /* driver.FindElementByName("Import").Click();
                 var SWV_drpdwn2 = driver.FindElementByAccessibilityId("ddlVersion");
                 SWV_drpdwn2.Click();
                 SWV_drpdwn2.SendKeys(Keys.Down);
                 SWV_drpdwn2.Click();
                 driver.FindElementByName("Close").Click();
                */
                // driver.FindElementByName("Select Status Code Row 0").Click();
                // driver.FindElementByName("Add").Click();
                // driver.FindElementByName("OK").Click();
                // driver.FindElementByName("OK").Click();
                //Filter By 
                var FG_drpdwn2 = driver.FindElementByAccessibilityId("ddlFunctionalGroup");
                FG_drpdwn2.Click();
                FG_drpdwn2.SendKeys(Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down);
                FG_drpdwn2.Click();
                // Type
                var Type_drpdwn2 = driver.FindElementByAccessibilityId("ddlType");
                Type_drpdwn2.Click();
                Type_drpdwn2.SendKeys(Keys.Down + Keys.Down);
                Type_drpdwn2.Click();
                //Change to Hub controller
                var radioFirst2 = driver.FindElementByName("Hub Controller");
                radioFirst2.Click();
                //Filter By 
                var FG_drpdwn3 = driver.FindElementByAccessibilityId("ddlFunctionalGroup");
                FG_drpdwn3.Click();
                FG_drpdwn3.SendKeys(Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down);
                FG_drpdwn3.Click();
                // Type
                var Type1_drpdwn2 = driver.FindElementByAccessibilityId("ddlType");
                Type1_drpdwn2.Click();
                Type1_drpdwn2.SendKeys(Keys.Down + Keys.Down);
                Type1_drpdwn2.Click();
                //Clear
                driver.FindElementByName("Clear").Click();
                //Export
                var btnExport = driver.FindElementByName("Export");
                btnExport.Click();
                var btn_Export = driver.FindElementByAccessibilityId("btnExport");
                btn_Export.Click();
                var btn_Save = driver.FindElementByName("Save");
                btn_Save.Click();
                driver.FindElementByName("Yes").Click();
                driver.FindElementByName("OK").Click();
                //SEARCH
                var search = driver.FindElementByAccessibilityId("txtsearch");
                search.SendKeys("0");
                driver.FindElementByName("0 - System OK").Click();
                Thread.Sleep(2000);
                driver.FindElementByAccessibilityId("btnClear").Click();
                reportHelperS.LogTestStep(Status.Pass, "Status Code - Import, Export, SoftwareVersion, Filter, Type, Search and Clear");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"**** Parameters TCS fail executing****" + ex.Message);
                reportHelperS.LogTestStep(Status.Fail, "Status Code - Import, Export, SoftwareVersion, Filter, Type, Search and Clear               " + ex.Message);
                reportHelperS.LogScreenshot(Status.Fail, "Status Code - Import, Export, SoftwareVersion, Filter, Type, Search and Clear", "Screenshot of failure", driver);
            }
        }
        [Test, Order(8)]
       //   [Ignore("Ignoring for quick test")]
        public void Add_Variable_List_Main()
        {
            try
            {
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                Thread.Sleep(5000);
                var windowHandles = driver.WindowHandles;
                for (int attempt = 0; attempt < 5 && windowHandles.Count > 1; attempt++)
                {
                    Thread.Sleep(TimeSpan.FromSeconds(1));
                    windowHandles = driver.WindowHandles;
                }
                if (windowHandles.Count >= 0)
                {
                    driver.SwitchTo().Window(windowHandles[0]);
                }
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                //Name="Status Code" AutomationId="textDescription"
                driver.FindElementByName("Variable List").Click();
                driver.FindElementByName("Add").Click();
                var FG_drpdwn2 = driver.FindElementByAccessibilityId("ddlFunctionalGroups");
                FG_drpdwn2.Click();
                FG_drpdwn2.SendKeys(Keys.Down);
                FG_drpdwn2.Click();
                //ddlFunctionalSystems txtVariableName txtSystemName txtUnit checkBoxIsCustomerVisible txtDescription "btnAdd"
                var FS_drpdwn2 = driver.FindElementByAccessibilityId("ddlFunctionalSystems");
                FS_drpdwn2.Click();
                FS_drpdwn2.SendKeys(Keys.Down);
                FS_drpdwn2.Click();
                driver.FindElementByAccessibilityId("txtVariableName").SendKeys("CURD operations Varaible name");
                driver.FindElementByAccessibilityId("txtSystemName").SendKeys("System name");
                driver.FindElementByAccessibilityId("txtUnit").SendKeys("kNm");
                // driver.FindElementByAccessibilityId("checkBoxIsCustomerVisiblen").Click();
                driver.FindElementByAccessibilityId("txtDescription").SendKeys("Variable list - CURD operations");
                driver.FindElementByAccessibilityId("btnAdd").Click();
                driver.FindElementByName("OK").Click();
                driver.FindElementByName("OK").Click();
                //EDIT
                driver.FindElementByName("Functional Group Row 0").Click();
                driver.FindElementByName("Edit").Click();
                //  driver.FindElementByName("View Audit Logs").Click();
                //  driver.FindElementByName("Close").Click();
                driver.FindElementByAccessibilityId("txtVariableName").Clear();
                driver.FindElementByAccessibilityId("txtVariableName").SendKeys("new variable name");
                driver.FindElementByName("Update").Click();
                driver.FindElementByName("OK").Click();
                driver.FindElementByName("OK").Click();
                //DELETE
                driver.FindElementByName("Functional Group Row 0").Click();
                driver.FindElementByName("Delete").Click();
                driver.FindElementByName("Yes").Click();
                driver.FindElementByName("OK").Click();
                reportHelperS.LogTestStep(Status.Pass, "Variable List - CRUD operations - Main Controller");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"**** Parameters TCS fail executing****" + ex.Message);
                reportHelperS.LogTestStep(Status.Fail, "Variable List - Add, Edit and Delete - Main Controller" + ex.Message);

                reportHelperS.LogScreenshot(Status.Fail, "Variable List - Add, Edit and Delete - Main Controller", "Screenshot of failure", driver);
            }
        }
        [Test, Order(9)]
     //    [Ignore("Ignoring for quick test")]
        public void Add_Hub_Variable_List()
        {
            try
            {
                //Thread.Sleep(9000);
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                //Name="Status Code" AutomationId="textDescription"
                driver.FindElementByName("Variable List").Click();
                //Change to Hub controller
                var radioFirst2 = driver.FindElementByName("Hub Controller");
                radioFirst2.Click();
                driver.FindElementByName("Add").Click();

                var FG_drpdwn2 = driver.FindElementByAccessibilityId("ddlFunctionalGroups");
                FG_drpdwn2.Click();
                FG_drpdwn2.SendKeys(Keys.Down);
                FG_drpdwn2.Click();
                //ddlFunctionalSystems txtVariableName txtSystemName txtUnit checkBoxIsCustomerVisible txtDescription "btnAdd"
                var FS_drpdwn2 = driver.FindElementByAccessibilityId("ddlFunctionalSystems");
                FS_drpdwn2.Click();
                FS_drpdwn2.SendKeys(Keys.Down);
                FS_drpdwn2.Click();

                driver.FindElementByAccessibilityId("txtVariableName").SendKeys("Varaible name");
                driver.FindElementByAccessibilityId("txtSystemName").SendKeys("System name new");
                driver.FindElementByAccessibilityId("txtUnit").SendKeys("kNm");
                // driver.FindElementByAccessibilityId("checkBoxIsCustomerVisiblen").Click();
                driver.FindElementByAccessibilityId("txtDescription").SendKeys("Variable list - CURD operations");
                driver.FindElementByAccessibilityId("btnAdd").Click();
                driver.FindElementByName("OK").Click();
                driver.FindElementByName("OK").Click();

                //EDIT
                driver.FindElementByName("Functional Group Row 0").Click();
                driver.FindElementByName("Edit").Click();
                //  driver.FindElementByName("View Audit Logs").Click();
                //   driver.FindElementByName("Close").Click();
                driver.FindElementByAccessibilityId("txtVariableName").Clear();
                driver.FindElementByAccessibilityId("txtVariableName").SendKeys("new variable name for HubController");
                driver.FindElementByName("Update").Click();
                driver.FindElementByName("OK").Click();
                driver.FindElementByName("OK").Click();

                //DELETE
                driver.FindElementByName("Functional Group Row 0").Click();
                driver.FindElementByName("Delete").Click();
                driver.FindElementByName("Yes").Click();
                driver.FindElementByName("OK").Click();

                reportHelperS.LogTestStep(Status.Pass, "Variable List - Add, Edit and Delete - HUB Controller");
                // reportHelperS.LogTestStep(Status.Pass, "Variable_List");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"**** Parameters TCS fail executing****" + ex.Message);
                reportHelperS.LogTestStep(Status.Fail, "Variable_List - Add, Edit and Delete - HUB Controller" + ex.Message);
                reportHelperS.LogScreenshot(Status.Fail, "Variable_List  - Add, Edit and Delete - HUB Controller", "Screenshot of failure", driver);
            }
        }
        [Test, Order(10)]
      //     [Ignore("Ignoring for quick test")]
        public void VariableList_ImportExportFilterSearch()
        {
            try
            {
                // Thread.Sleep(9000);
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);

                //   driver.FindElementByName("Import").Click();
                //  var SWV_drpdwn2 = driver.FindElementByAccessibilityId("ddlVersion");
                //  SWV_drpdwn2.Click();
                //   SWV_drpdwn2.SendKeys(Keys.Down);
                //    SWV_drpdwn2.Click();
                //     driver.FindElementByName("Close").Click();

                //Filter By 
                var FG_drpdwn2 = driver.FindElementByAccessibilityId("ddlFunctionalGroup");
                FG_drpdwn2.Click();
                FG_drpdwn2.SendKeys(Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down);
                FG_drpdwn2.Click();

                // Type
                var Type_drpdwn2 = driver.FindElementByAccessibilityId("ddlFunctionalSystem");
                Type_drpdwn2.Click();
                Type_drpdwn2.SendKeys(Keys.Down + Keys.Down);
                Type_drpdwn2.Click();

                //Change to Hub controller
                var radioFirst2 = driver.FindElementByName("Hub Controller");
                radioFirst2.Click();
                //Filter By 
                var FG_drpdwn3 = driver.FindElementByAccessibilityId("ddlFunctionalGroup");
                FG_drpdwn3.Click();
                FG_drpdwn3.SendKeys(Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down);
                FG_drpdwn3.Click();

                // Type
                var Type1_drpdwn2 = driver.FindElementByAccessibilityId("ddlFunctionalSystems");
                Type1_drpdwn2.Click();
                Type1_drpdwn2.SendKeys(Keys.Down + Keys.Down);
                Type1_drpdwn2.Click();

                //Clear
                driver.FindElementByName("Clear").Click();

                //Export
                var btnExport = driver.FindElementByName("Export");
                btnExport.Click();
                var btn_Export = driver.FindElementByAccessibilityId("btnExport");
                btn_Export.Click();
                var btn_Save = driver.FindElementByName("Save");
                btn_Save.Click();
                // driver.FindElementByName("Yes").Click();
                driver.FindElementByName("OK").Click();


                //Change to Hub controller
                /*  var radioFirst33 = driver.FindElementByName("Main Controller");
                  radioFirst33.Click();
                   //SEARCH
                  var search = driver.FindElementByAccessibilityId("txtsearch");
                  search.SendKeys("Client IP - ActiveState.ClientIp");
                  driver.FindElementByName("Client IP - ActiveState.ClientIp").Click();
                  Thread.Sleep(3000);
                  driver.FindElementByAccessibilityId("btnClear").Click();*/

                reportHelperS.LogTestStep(Status.Pass, "Variable List - Import Export Filter Search");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"**** Parameters TCS fail executing****" + ex.Message);
                reportHelperS.LogTestStep(Status.Fail, "Variable List - Import Export Filter Search" + ex.Message);
                reportHelperS.LogScreenshot(Status.Fail, "Variable List - Import Export Filter Search", "Screenshot of failure", driver);
            }
        }



        [Test, Order(11)]
      //  [Ignore("Ignoring for quick test")]
        public void IO_CRUD()
        {
            try
            {
              
                var windowHandles = driver.WindowHandles;
                for (int attempt = 0; attempt < 5 && windowHandles.Count > 1; attempt++)
                {
                    Thread.Sleep(TimeSpan.FromSeconds(1));
                    windowHandles = driver.WindowHandles;
                }

                if (windowHandles.Count >= 0)
                {
                    driver.SwitchTo().Window(windowHandles[0]);
                }

                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                driver.FindElementByName("IO").Click();
                driver.FindElementByName("Add").Click();

                var FG_drpdwn2 = driver.FindElementByAccessibilityId("ddlIOType");
                FG_drpdwn2.Click();
                FG_drpdwn2.SendKeys(Keys.Down);
                FG_drpdwn2.Click();

                driver.FindElementByAccessibilityId("txtRevision").SendKeys("System");
                driver.FindElementByAccessibilityId("txtSlotNumber").SendKeys("11");
                driver.FindElementByAccessibilityId("txtModuleType").SendKeys("SY3429888");
                driver.FindElementByAccessibilityId("txtAssignmentEnglish").SendKeys("System");
                driver.FindElementByAccessibilityId("txtNumberCharactersEnglish").SendKeys("6");
                driver.FindElementByAccessibilityId("EditIO").SendKeys("Sy");
                driver.FindElementByAccessibilityId("txtRemardks").SendKeys("System");

                driver.FindElementByAccessibilityId("btnAdd").Click();
                driver.FindElementByName("OK").Click();
                driver.FindElementByName("OK").Click();
                //EDIT
                driver.FindElementByName("Revision Row 0").Click();
                driver.FindElementByName("Edit").Click();
                //driver.FindElementByName("View Audit Logs").Click();
                // driver.FindElementByName("Close").Click();
                driver.FindElementByAccessibilityId("txtCodeVerification").SendKeys("VED");
                driver.FindElementByAccessibilityId("txtModuleType").Clear();
                driver.FindElementByAccessibilityId("txtModuleType").SendKeys("IO241805");
                driver.FindElementByName("Update").Click();
                driver.FindElementByName("OK").Click();
                driver.FindElementByName("OK").Click();

                //DELETE
                driver.FindElementByName("Revision Row 0").Click();
                driver.FindElementByName("Delete").Click();
                driver.FindElementByName("Yes").Click();
                driver.FindElementByName("OK").Click();

                //Dropdown 
                var SW_drpdwn2 = driver.FindElementByAccessibilityId("ddlVersion");
                SW_drpdwn2.Click();
                SW_drpdwn2.SendKeys(Keys.Down);
                SW_drpdwn2.Click();

                var IO_drpdwn2 = driver.FindElementByAccessibilityId("ddlIOType");
                IO_drpdwn2.Click();
                IO_drpdwn2.SendKeys(Keys.Down);
                IO_drpdwn2.Click();

                //Export
                var btnExport = driver.FindElementByName("Export");
                btnExport.Click();
                var btn_Export = driver.FindElementByAccessibilityId("btnExport");
                btn_Export.Click();
                var btn_Save = driver.FindElementByName("Save");
                btn_Save.Click();
                driver.FindElementByName("Yes").Click();
                driver.FindElementByName("OK").Click();

                //SEARCH
                // var search = driver.FindElementByAccessibilityId("txtsearch");
                // search.SendKeys("0");
                // driver.FindElementByName("0 - System OK").Click();
                //  Thread.Sleep(2000);
                //  driver.FindElementByAccessibilityId("btnClear").Click();


                reportHelperS.LogTestStep(Status.Pass, "IO - Add, Edit,Delete and Export");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"**** Parameters TCS fail executing****" + ex.Message);
                reportHelperS.LogTestStep(Status.Fail, "IO - Add, Edit,Delete and Export" + ex.Message);
                reportHelperS.LogScreenshot(Status.Fail, "IO - Add, Edit,Delete and Export", "Screenshot of failure", driver);
            }
        }
        [Test, Order(12)]
     //  [Ignore("Ignoring for quick test")]
        public void SLC_Parameters()
        {
            try
            {
                // Thread.Sleep(9000);
                WindowsElement folderCompareButton = driver.FindElementByName("SLC");
                folderCompareButton.Click();
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                driver.FindElementByName("SLC Parameters").Click();

                driver.FindElementByName("Add").Click();

                var FG_drpdwn2 = driver.FindElementByAccessibilityId("ddlParaIdent");
                FG_drpdwn2.Click();
                FG_drpdwn2.SendKeys(Keys.Down);
                FG_drpdwn2.Click();

                driver.FindElementByAccessibilityId("txtParameterName").SendKeys("System SLC");
                driver.FindElementByAccessibilityId("txtValue").SendKeys("2");
                driver.FindElementByAccessibilityId("txtUnit").SendKeys("1");


                driver.FindElementByAccessibilityId("btnAdd").Click();
                driver.FindElementByName("Yes").Click();
                driver.FindElementByName("OK").Click();

                //EDIT
                driver.FindElementByName("SLCParameter Name Row 0").Click();
                driver.FindElementByName("Edit").Click();
                // driver.FindElementByName("View Audit Logs").Click();
                // driver.FindElementByName("Close").Click();
                driver.FindElementByAccessibilityId("txtParameterName").Clear();
                driver.FindElementByAccessibilityId("txtParameterName").SendKeys("System SLC new");
                driver.FindElementByAccessibilityId("txtValue").Clear();
                driver.FindElementByAccessibilityId("txtValue").SendKeys("22");
                driver.FindElementByName("Update").Click();
                driver.FindElementByName("OK").Click();
                driver.FindElementByName("OK").Click();

                //DELETE
                driver.FindElementByName("SLCParameter Name Row 0").Click();
                driver.FindElementByName("Delete").Click();
                driver.FindElementByName("Yes").Click();
                driver.FindElementByName("OK").Click();


                reportHelperS.LogTestStep(Status.Pass, "SLC Parameters - Add, Edit and Delete");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"**** Parameters TCS fail executing****" + ex.Message);
                reportHelperS.LogTestStep(Status.Fail, "SLC Parameters - Add, Edit and Delete" + ex.Message);
                reportHelperS.LogScreenshot(Status.Fail, "SLC Parameters - Add, Edit and Delete", "Screenshot of failure", driver);
            }
        }
        [Test, Order(13)]
      // [Ignore("Ignoring for quick test")]
        public void SLC_Parameters_ImportExport()
        {
            try
            {

                // Thread.Sleep(10000);
                /* driver.FindElementByName("Import").Click();

                     var SWV_drpdwn2 = driver.FindElementByAccessibilityId("ddlVersion");
                     SWV_drpdwn2.Click();
                     SWV_drpdwn2.SendKeys(Keys.Down);
                     SWV_drpdwn2.Click();
                     var FG_drpdwn2 = driver.FindElementByAccessibilityId("ddlParameterIdent");
                     FG_drpdwn2.Click();
                     FG_drpdwn2.SendKeys(Keys.Down + Keys.Down);
                     FG_drpdwn2.Click();
                     driver.FindElementByName("Close").Click(); */

                //Filter By ParameterIdent
                var P_drpdwn2 = driver.FindElementByAccessibilityId("ddlParaIdent");
                P_drpdwn2.Click();
                P_drpdwn2.SendKeys(Keys.Down + Keys.Down);
                P_drpdwn2.Click();


                //Change to Hub controller
                var radioFirst2 = driver.FindElementByName("Hub Controller");
                radioFirst2.Click();
                //Filter By 
                var FG_drpdwn3 = driver.FindElementByAccessibilityId("ddlFunctionalGroup");
                FG_drpdwn3.Click();
                FG_drpdwn3.SendKeys(Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down);
                FG_drpdwn3.Click();



                //Clear
                driver.FindElementByName("Clear").Click();

                //Export
                var btnExport = driver.FindElementByName("Export");
                btnExport.Click();
                var btn_Export = driver.FindElementByAccessibilityId("btnExport");
                btn_Export.Click();
                var btn_Save = driver.FindElementByName("Save");
                btn_Save.Click();
                driver.FindElementByName("Yes").Click();
                driver.FindElementByName("OK").Click();

                //SEARCH
                //  var search = driver.FindElementByAccessibilityId("txtsearch");
                //  search.SendKeys("0");
                //  driver.FindElementByName("0 - System OK").Click();
                //  Thread.Sleep(2000);
                // driver.FindElementByAccessibilityId("btnClear").Click();
                reportHelperS.LogTestStep(Status.Pass, "SLC Parameters - Import, Export and Filter functions");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"**** Parameters TCS fail executing****" + ex.Message);
                reportHelperS.LogTestStep(Status.Fail, "SLC Parameters - Import, Export and Filter functions" + ex.Message);
                reportHelperS.LogScreenshot(Status.Fail, "SLC Parameters - Import, Export and Filter functions", "Screenshot of failure", driver);
            }
        }

        [Test, Order(14)]
    //  [Ignore("Ignoring for quick test")]
        public void SLCVariables_Add_Edit_Delete()
        {
            try
            {
                WindowsElement Button = driver.FindElementByName("SLC");
                Button.Click();
                Button.SendKeys(Keys.Down + Keys.Down + Keys.Enter);
                // Button.Click();
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                Thread.Sleep(5000);
                driver.FindElementByName("Add").Click();
                var FG_drpdwn2 = driver.FindElementByAccessibilityId("ddlParaIdent");
                FG_drpdwn2.Click();
                FG_drpdwn2.SendKeys(Keys.Down);
                FG_drpdwn2.Click();
                driver.FindElementByAccessibilityId("txtVariableName").SendKeys("System SLC Variable");
                driver.FindElementByAccessibilityId("txtValue").SendKeys("2");
                driver.FindElementByAccessibilityId("txtUnit").SendKeys("1");
                driver.FindElementByAccessibilityId("txtIndex").SendKeys("1");
                driver.FindElementByAccessibilityId("btnAdd").Click();
                driver.FindElementByName("Yes").Click();
                driver.FindElementByName("OK").Click();
                //EDIT
                driver.FindElementByName("SLCVariable Name Row 0").Click();
                driver.FindElementByName("Edit").Click();
                // driver.FindElementByName("View Audit Logs").Click();
                //  driver.FindElementByName("Close").Click();
                driver.FindElementByAccessibilityId("txtVariableName").Clear();
                driver.FindElementByAccessibilityId("txtVariableName").SendKeys("System SLC Variable new");
                driver.FindElementByAccessibilityId("txtValue").Clear();
                driver.FindElementByAccessibilityId("txtValue").SendKeys("22");
                driver.FindElementByName("Update").Click();
                driver.FindElementByName("OK").Click();
                driver.FindElementByName("OK").Click();
                //DELETE
                driver.FindElementByName("SLCVariable Name Row 0").Click();
                driver.FindElementByName("Delete").Click();
                driver.FindElementByName("Yes").Click();
                driver.FindElementByName("OK").Click();

                reportHelperS.LogTestStep(Status.Pass, "SLC Variables - Add, Edit and Delete operations");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"**** Parameters TCS fail executing****" + ex.Message);
                reportHelperS.LogTestStep(Status.Fail, "SLC Variables - Add, Edit and Delete operations" + ex.Message);
                reportHelperS.LogScreenshot(Status.Fail, "SLC Variables - Add, Edit and Delete operations", "Screenshot of failure", driver);
            }
        }
        [Test, Order(15)]
      // [Ignore("Ignoring for quick test")]
        public void SLC_Variables_ImportExport()
        {
            try
            {
                driver.FindElementByName("Import").Click();
                var SWV_drpdwn2 = driver.FindElementByAccessibilityId("ddlVersion");
                SWV_drpdwn2.Click();
                SWV_drpdwn2.SendKeys(Keys.Down);
                SWV_drpdwn2.Click();
                var FG_drpdwn2 = driver.FindElementByAccessibilityId("ddlParameterIdent");
                FG_drpdwn2.Click();
                FG_drpdwn2.SendKeys(Keys.Down + Keys.Down);
                FG_drpdwn2.Click();
                driver.FindElementByName("Close").Click();
                //Filter By ParameterIdent
                var P_drpdwn2 = driver.FindElementByAccessibilityId("ddlParaIdent");
                P_drpdwn2.Click();
                P_drpdwn2.SendKeys(Keys.Down + Keys.Down);
                P_drpdwn2.Click();
                //Change to Hub controller
                var radioFirst2 = driver.FindElementByName("Hub Controller");
                radioFirst2.Click();
                //Filter By 
                var FG_drpdwn3 = driver.FindElementByAccessibilityId("ddlParaIdent");
                FG_drpdwn3.Click();
                FG_drpdwn3.SendKeys(Keys.Down + Keys.Down + Keys.Down);
                FG_drpdwn3.Click();
                //Clear
                driver.FindElementByName("Clear").Click();
                //Export
                var btnExport = driver.FindElementByName("Export");
                btnExport.Click();
                var btn_Export = driver.FindElementByAccessibilityId("btnExport");
                btn_Export.Click();
                var btn_Save = driver.FindElementByName("Save");
                btn_Save.Click();
                driver.FindElementByName("Yes").Click();
                driver.FindElementByName("OK").Click();

                reportHelperS.LogTestStep(Status.Pass, "SLC Variables -  Import and Export functions");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"**** Parameters TCS fail executing****" + ex.Message);
                reportHelperS.LogTestStep(Status.Fail, "SLC Variables -  Import and Export functions" + ex.Message);
                reportHelperS.LogScreenshot(Status.Fail, "SLC Variables -  Import and Export functions", "Screenshot of failure", driver);
            }
        }
        [Test, Order(16)]
     //  [Ignore("Ignoring for quick test")]
        public void PLC_Controller()
        {
            try
            {
                
                var windowHandles = driver.WindowHandles;
                for (int attempt = 0; attempt < 5 && windowHandles.Count > 1; attempt++)
                {
                    Thread.Sleep(TimeSpan.FromSeconds(1));
                    windowHandles = driver.WindowHandles;
                }

                if (windowHandles.Count >= 0)
                {
                    driver.SwitchTo().Window(windowHandles[0]);
                }
                WindowsElement Button = driver.FindElementByName("PLC");
                Button.Click();
                Button.SendKeys(Keys.Down + Keys.Enter);
                // Button.Click();
                Thread.Sleep(5000);
                driver.FindElementByAccessibilityId("textBoxName").SendKeys("Ved");
                //192.168.211.120
                var SWV_drpdwn2 = driver.FindElementByAccessibilityId("ddlIPAddress");
                SWV_drpdwn2.Click();
                SWV_drpdwn2.SendKeys(Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Enter);
                SWV_drpdwn2.Click();
                driver.FindElementByAccessibilityId("txtUserName").SendKeys("toby");
                driver.FindElementByAccessibilityId("txtPassword").SendKeys("test1234");
                driver.FindElementByAccessibilityId("buttonConnect").Click();

                reportHelperS.LogTestStep(Status.Pass, "PLC Controller Connection ");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"**** Parameters TCS fail executing****" + ex.Message);
                reportHelperS.LogTestStep(Status.Fail, "PLC Controller" + ex.Message);
                reportHelperS.LogScreenshot(Status.Fail, "PLC Controller", "Screenshot of failure", driver);
            }
        }
        [Test, Order(17)]
      //  [Ignore("Ignoring for quick test")]
        public void PLC_EditController()
        {
            try
            {
                WindowsElement Button = driver.FindElementByName("PLC");
                Button.Click();
                Button.SendKeys(Keys.Down + Keys.Down + Keys.Down + Keys.Enter);
                // Button.Click(); "  btnAdd
                driver.FindElementByName("Add").Click();
                driver.FindElementByAccessibilityId("txtName").SendKeys("Aar");
                driver.FindElementByAccessibilityId("txtIPAddress").SendKeys("192168211120");

                driver.FindElementByAccessibilityId("btnAdd").Click();
                driver.FindElementByName("OK").Click();
                driver.FindElementByName("OK").Click();
                //EDIT  
                driver.FindElementByName("NAME Row 4").Click();
                driver.FindElementByName("Edit").Click();
                // driver.FindElementByName("View Audit Logs").Click();
                //  driver.FindElementByName("Close").Click();
                driver.FindElementByAccessibilityId("txtName").Clear();
                driver.FindElementByAccessibilityId("txtName").SendKeys("Arya");
                driver.FindElementByAccessibilityId("txtIPAddress").Clear();
                driver.FindElementByAccessibilityId("txtIPAddress").SendKeys("192168211120");
                driver.FindElementByName("Update").Click();
                driver.FindElementByName("OK").Click();
                driver.FindElementByName("OK").Click();

                //DELETE
                driver.FindElementByName("NAME Row 4").Click();
                driver.FindElementByName("Delete").Click();
                driver.FindElementByName("Yes").Click();
                driver.FindElementByName("OK").Click();


                reportHelperS.LogTestStep(Status.Pass, "PLC - Edit Controller ");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"**** Parameters TCS fail executing****" + ex.Message);
                reportHelperS.LogTestStep(Status.Fail, "PLC - Edit Controller" + ex.Message);
                reportHelperS.LogScreenshot(Status.Fail, "PLC - Edit Controller", "Screenshot of failure", driver);
            }
        }
        [Test, Order(18)]
       // [Ignore("Ignoring for quick test")]
        public void PLC_Generate_DEF_file()
        {
            try
            {
                WindowsElement Button = driver.FindElementByName("PLC");
                Button.Click();
                Button.SendKeys(Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Enter + Keys.Enter);
                driver.FindElementByName("Add").Click();

                var ModuleType = driver.FindElementByAccessibilityId("ddlModuleType");
                ModuleType.Click();
                ModuleType.SendKeys(Keys.Down);
                ModuleType.Click();

                var DefType = driver.FindElementByAccessibilityId("ddlDefType");
                DefType.Click();
                DefType.SendKeys(Keys.Down);
                DefType.Click();

                driver.FindElementByAccessibilityId("txtSystemName").SendKeys("System");
                driver.FindElementByAccessibilityId("txtMinValue").SendKeys("0");
                driver.FindElementByAccessibilityId("txtMaxValue").SendKeys("10");
                driver.FindElementByAccessibilityId("txtValue").SendKeys("6");
                driver.FindElementByAccessibilityId("txtParameter").SendKeys("Ved Parameter");

                driver.FindElementByAccessibilityId("btnAdd").Click();
                driver.FindElementByName("Yes").Click();
                driver.FindElementByName("OK").Click();

                //Edit 
                driver.FindElementByName("Module Name Row 0").Click();
                driver.FindElementByName("Edit").Click();
                // driver.FindElementByName("View Audit Logs").Click();
                //  driver.FindElementByName("Close").Click();
                driver.FindElementByAccessibilityId("txtMinValue").Clear();
                driver.FindElementByAccessibilityId("txtMinValue").SendKeys("2");
                driver.FindElementByAccessibilityId("txtValue").Clear();
                driver.FindElementByAccessibilityId("txtValue").SendKeys("9");
                driver.FindElementByName("Update").Click();
                driver.FindElementByName("OK").Click();
                driver.FindElementByName("OK").Click();

                //DELETE
                driver.FindElementByName("Module Name Row 0").Click();
                driver.FindElementByName("Delete").Click();
                driver.FindElementByName("Yes").Click();
                driver.FindElementByName("OK").Click();

                //Generate Def file  
                driver.FindElementByAccessibilityId("btnGenerateDefFile").Clear();
                driver.FindElementByName("This PC").Click();
                driver.FindElementByName("Documents").Click();
                driver.FindElementByName("OK").Click();
                driver.FindElementByName("OK").Click();


                reportHelperS.LogTestStep(Status.Pass, "PLC Generate DEF file ");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"**** Parameters TCS fail executing****" + ex.Message);
                reportHelperS.LogTestStep(Status.Fail, "PLC Generate DEF file" + ex.Message);
                reportHelperS.LogScreenshot(Status.Fail, "PLC Generate DEF file", "Screenshot of failure", driver);
            }
        }
        [Test, Order(19)]
      //  [Ignore("Ignoring for quick test")]
        public void PLC_DEF_Test()
        {
            try
            {
                WindowsElement Button = driver.FindElementByName("PLC");
                Button.Click();
                Button.SendKeys(Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.ArrowRight + Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Enter);

                driver.FindElementByName("Add").Click();
                driver.FindElementByAccessibilityId("txtSystemName").SendKeys("System");
                driver.FindElementByAccessibilityId("txtMinValue").SendKeys("0");
                driver.FindElementByAccessibilityId("txtMaxValue").SendKeys("10");
                driver.FindElementByAccessibilityId("txtValue").SendKeys("6");
                driver.FindElementByAccessibilityId("txtParameter").SendKeys("Ved Parameter");

                driver.FindElementByAccessibilityId("btnAdd").Click();
                driver.FindElementByName("Yes").Click();
                driver.FindElementByName("OK").Click();

                //Edit 
                driver.FindElementByName("Module Name Row 0").Click();
                driver.FindElementByName("Edit").Click();
                // driver.FindElementByName("View Audit Logs").Click();
                //  driver.FindElementByName("Close").Click();
                driver.FindElementByAccessibilityId("txtMinValue").Clear();
                driver.FindElementByAccessibilityId("txtMinValue").SendKeys("2");
                driver.FindElementByAccessibilityId("txtValue").Clear();
                driver.FindElementByAccessibilityId("txtValue").SendKeys("9");
                driver.FindElementByName("Update").Click();
                driver.FindElementByName("OK").Click();
                driver.FindElementByName("OK").Click();

                //DELETE
                driver.FindElementByName("Module Name Row 0").Click();
                driver.FindElementByName("Delete").Click();
                driver.FindElementByName("Yes").Click();
                driver.FindElementByName("OK").Click();
                reportHelperS.LogTestStep(Status.Pass, "PLC - DEF Test ");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"**** Parameters TCS fail executing****" + ex.Message);
                reportHelperS.LogTestStep(Status.Fail, "PLC - DEF Test" + ex.Message);
                reportHelperS.LogScreenshot(Status.Fail, "PLC - DEF Test", "Screenshot of failure", driver);
            }
        }
        [Test, Order(20)]
       // [Ignore("Ignoring for quick test")]
        public void Tools_FunctionalGroups()
        {
            try
            {
                /* Thread.Sleep(5000);
 var windowHandles = driver.WindowHandles;
 for (int attempt = 0; attempt < 5 && windowHandles.Count > 1; attempt++)
 {
     Thread.Sleep(TimeSpan.FromSeconds(1));
     windowHandles = driver.WindowHandles;
 }

 if (windowHandles.Count >= 0)
 {
     driver.SwitchTo().Window(windowHandles[0]);
 }*/

                WindowsElement Button = driver.FindElementByName("Tools");
                Button.Click();
                Button.SendKeys(Keys.Down + Keys.Enter);
                Button.Click();

                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                Thread.Sleep(5000);
                driver.FindElementByName("Add").Click();

                driver.FindElementByAccessibilityId("txtName").SendKeys("System SLC Variable");
                var FG_drpdwn2 = driver.FindElementByAccessibilityId("ddlOwners");
                FG_drpdwn2.Click();
                FG_drpdwn2.SendKeys(Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down);
                FG_drpdwn2.Click();


                driver.FindElementByAccessibilityId("btnAdd").Click();
                driver.FindElementByName("OK").Click();
                driver.FindElementByName("OK").Click();

                //EDIT
                InputSimulator simulator1 = new InputSimulator();
                simulator1.Mouse.VerticalScroll(-15);

                /*
                driver.FindElementByName("Line down").Click();
                driver.FindElementByName("Line down").Click();
                driver.FindElementByName("Line down").Click();
                driver.FindElementByName("Line down").Click();
                driver.FindElementByName("Line down").Click();
                driver.FindElementByName("Line down").Click();
                driver.FindElementByName("Line down").Click();
                driver.FindElementByName("Line down").Click();
                driver.FindElementByName("Line down").Click();
                driver.FindElementByName("Line down").Click();
                driver.FindElementByName("Line down").Click();
                */

                driver.FindElementByName("Functional Group Name Row 32").Click();
                driver.FindElementByName("Edit").Click();
                // driver.FindElementByName("View Audit Logs").Click();
                // driver.FindElementByName("Close").Click();
                driver.FindElementByAccessibilityId("txtName").Clear();
                driver.FindElementByAccessibilityId("txtName").SendKeys("System SLC Variable new");
                driver.FindElementByName("Update").Click();
                driver.FindElementByName("OK").Click();
                driver.FindElementByName("OK").Click();

                //DELETE
                driver.FindElementByName("Functional Group Name Row 32").Click();
                driver.FindElementByName("Delete").Click();
                driver.FindElementByName("Yes").Click();
                driver.FindElementByName("OK").Click();


                reportHelperS.LogTestStep(Status.Pass, "Tools - Add,Edit,Delete - Functional Group");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"**** Parameters TCS fail executing****" + ex.Message);
                reportHelperS.LogTestStep(Status.Fail, "Tools - Add,Edit,Delete - Functional Group" + ex.Message);
                reportHelperS.LogScreenshot(Status.Fail, "Tools - Functional Group", "Screenshot of failure", driver);
            }
        }
        [Test, Order(21)]
      //  [Ignore("Ignoring for quick test")]
        public void Tools_FunctionalSystem()
        {
            try
            {
                WindowsElement Button = driver.FindElementByName("Tools");
                Button.Click();
                Button.SendKeys(Keys.Down + Keys.Down + Keys.Enter);
                Button.Click();

                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                Thread.Sleep(5000);
                driver.FindElementByName("Add").Click();

                var FG_drpdwn2 = driver.FindElementByAccessibilityId("ddlFunctionalGroups");
                FG_drpdwn2.Click();
                FG_drpdwn2.SendKeys(Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down);
                FG_drpdwn2.Click();
                driver.FindElementByAccessibilityId("txtSystemName").SendKeys("Ved");

                driver.FindElementByAccessibilityId("checkBoxIsStatusCode").Click();
                driver.FindElementByAccessibilityId("btnAdd").Click();
                driver.FindElementByName("OK").Click();
                driver.FindElementByName("OK").Click();

                //EDIT
                // Name = "Last"  Name="Line down AutomationId="txtsearch Name="ved  AutomationId="btnClear"
                driver.FindElementByAccessibilityId("txtsearch").SendKeys("Ved");
                driver.FindElementByName("Ved").Click();
                driver.FindElementByName("Functional Group Name Row 0").Click();

                /*  driver.FindElementByName("Last").Click();
                  driver.FindElementByName("Line down").Click();
                  driver.FindElementByName("Line down").Click();
                  driver.FindElementByName("Line down").Click();
                  driver.FindElementByName("Line down").Click();
                  driver.FindElementByName("Line down").Click();
                  driver.FindElementByName("Line down").Click();
                  driver.FindElementByName("Line down").Click();
                  driver.FindElementByName("Line down").Click();
                  driver.FindElementByName("Line down").Click();
                  driver.FindElementByName("Line down").Click();
                  driver.FindElementByName("Line down").Click();
                  driver.FindElementByName("Functional Group Name Row 33").Click();*/

                driver.FindElementByName("Edit").Click();
                // driver.FindElementByName("View Audit Logs").Click();
                // driver.FindElementByName("Close").Click();
                driver.FindElementByAccessibilityId("txtSystemName").Clear();
                driver.FindElementByAccessibilityId("txtSystemName").SendKeys("FunctionalGroups New");
                driver.FindElementByName("Update").Click();
                driver.FindElementByName("OK").Click();
                driver.FindElementByName("OK").Click();
                driver.FindElementByAccessibilityId("btnClear").Click();

                //DELETE
                driver.FindElementByName("Functional Group Name Row 33").Click();
                driver.FindElementByName("Delete").Click();
                driver.FindElementByName("Yes").Click();
                driver.FindElementByName("OK").Click();
                reportHelperS.LogTestStep(Status.Pass, "Tools - Add,Edit,Delete - Functional System ");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"**** Parameters TCS fail executing****" + ex.Message);
                reportHelperS.LogTestStep(Status.Fail, "Tools - Add,Edit,Delete - Functional System" + ex.Message);
                reportHelperS.LogScreenshot(Status.Fail, "Tools - Functional System", "Screenshot of failure", driver);
            }
        }
        [Test, Order(22)]
      //  [Ignore("Ignoring for quick test")]
        public void Tools_ComponentVersion()
        {
            try
            {
                WindowsElement Button = driver.FindElementByName("Tools");
                Button.Click();
                Button.SendKeys(Keys.Down + Keys.Down + Keys.Down + Keys.Enter);
                Button.Click();

                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                Thread.Sleep(5000);
                driver.FindElementByName("Add").Click();
                var FG_drpdwn2 = driver.FindElementByAccessibilityId("ddlFunctionalGroups");
                FG_drpdwn2.Click();
                FG_drpdwn2.SendKeys(Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down);
                FG_drpdwn2.Click();
                driver.FindElementByAccessibilityId("txtName").SendKeys("Ved");

                driver.FindElementByAccessibilityId("checkedListBoxFunctionalSystem").Click();
                //Park - Master
                driver.FindElementByAccessibilityId("btnAdd").Click();
                driver.FindElementByName("OK").Click();
                driver.FindElementByName("OK").Click();

                // Filter ddlFunctionalGroups
                var FG_drpdwn3 = driver.FindElementByAccessibilityId("ddlFunctionalGroups");
                FG_drpdwn3.Click();
                FG_drpdwn3.SendKeys(Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down);
                FG_drpdwn3.Click();
                driver.FindElementByName("Component Version Name Row 1").Click();
                //EDIT
                driver.FindElementByName("Edit").Click();
                driver.FindElementByAccessibilityId("txtName").Clear();
                driver.FindElementByAccessibilityId("txtName").SendKeys("NewVed");
                driver.FindElementByName("Update").Click();
                driver.FindElementByName("OK").Click();
                driver.FindElementByName("OK").Click();

                //DELETE
                var FG_drpdwn4 = driver.FindElementByAccessibilityId("ddlFunctionalGroups");
                FG_drpdwn4.Click();
                FG_drpdwn4.SendKeys(Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down);
                FG_drpdwn4.Click();
                driver.FindElementByName("Component Version Name Row 1").Click();
                driver.FindElementByName("Delete").Click();
                driver.FindElementByName("Yes").Click();
                driver.FindElementByName("OK").Click();

                reportHelperS.LogTestStep(Status.Pass, "Tools - Add,Edit - Component Version ");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"**** Parameters TCS fail executing****" + ex.Message);
                reportHelperS.LogTestStep(Status.Fail, "Tools - Add,Edit - Component Version" + ex.Message);
                reportHelperS.LogScreenshot(Status.Fail, "Tools - Add,Edit - Component Version", "Screenshot of failure", driver);
            }
        }
        [Test, Order(23)]
       // [Ignore("Ignoring for quick test")]
        public void Tools_AddSWVersion()
        {
            try
            {

                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                Thread.Sleep(5000);
                WindowsElement Button = driver.FindElementByName("Tools");
                Button.Click();
                Button.SendKeys(Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Enter);
                Button.Click();
                driver.FindElementByAccessibilityId("txtSoftwareVersion").SendKeys("010711");
                driver.FindElementByAccessibilityId("btnAdd").Click();
                driver.FindElementByName("Yes").Click();
                driver.FindElementByName("OK").Click();


                reportHelperS.LogTestStep(Status.Pass, "Tools - Add SW Version ");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"**** Parameters TCS fail executing****" + ex.Message);
                reportHelperS.LogTestStep(Status.Fail, "Tools - Add SW Version" + ex.Message);
                reportHelperS.LogScreenshot(Status.Fail, "Tools - Add SW Version", "Screenshot of failure", driver);
            }
        }
        [Test, Order(24)]
       // [Ignore("Ignoring for quick test")]
        public void Tools_AddSWPlatform()
        {
            try
            {
                // driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                //  Thread.Sleep(5000);
                WindowsElement Button = driver.FindElementByName("Tools");
                Button.Click();
                Button.SendKeys(Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Enter);
                Button.Click();
                driver.FindElementByAccessibilityId("txtPlatformName").SendKeys("1284");
                driver.FindElementByAccessibilityId("btnAdd").Click();
                driver.FindElementByName("OK").Click();
                driver.FindElementByName("OK").Click();

                reportHelperS.LogTestStep(Status.Pass, "Tools - Add SW Platform ");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"**** Parameters TCS fail executing****" + ex.Message);
                reportHelperS.LogTestStep(Status.Fail, "Tools - Add SW Platform" + ex.Message);
                reportHelperS.LogScreenshot(Status.Fail, "Tools - Add SW Platform", "Screenshot of failure", driver);
            }
        }
        [Test, Order(25)]
       // [Ignore("Ignoring for quick test")]
        public void Tools_UserManagement()
        {
            try
            {
                var FG_drpdwn2 = driver.FindElementByAccessibilityId("ddlUserRole");
                FG_drpdwn2.Click();
                FG_drpdwn2.SendKeys(Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Enter);
                FG_drpdwn2.Click();

                var UserStatus_drpdwn2 = driver.FindElementByAccessibilityId("ddlUserStatus");
                UserStatus_drpdwn2.Click();
                UserStatus_drpdwn2.SendKeys(Keys.Down + Keys.Down);
                UserStatus_drpdwn2.Click();

                driver.FindElementByName("Add").Click();
                //ddlUserRole txtFirstName txtLastName txtUserMail txtUserName ddlUserStatus- drop txtPassword txtConfirmPassword
                var UserRole = driver.FindElementByAccessibilityId("ddlUserRole");
                UserRole.Click();
                UserRole.SendKeys(Keys.Down + Keys.Down + Keys.Down + Keys.Down);
                UserRole.Click();
                driver.FindElementByAccessibilityId("txtFirstName").SendKeys("NewVed");
                driver.FindElementByAccessibilityId("txtLastNamee").SendKeys("System SLC Variable");
                driver.FindElementByAccessibilityId("txtUserMail").SendKeys("NewVed");
                driver.FindElementByAccessibilityId("txtUserName").SendKeys("NewVed");
                var UserStatus1_drpdwn2 = driver.FindElementByAccessibilityId("ddlUserStatus");
                UserStatus1_drpdwn2.Click();
                UserStatus1_drpdwn2.SendKeys(Keys.Down + Keys.Down);
                UserStatus1_drpdwn2.Click();
                driver.FindElementByAccessibilityId("txtPassword").SendKeys("NewVed");
                driver.FindElementByAccessibilityId("txtConfirmPassword").SendKeys("NewVed");
                driver.FindElementByAccessibilityId("btnAdd").Click();
                driver.FindElementByName("Yes").Click();
                driver.FindElementByName("OK").Click();

                reportHelperS.LogTestStep(Status.Pass, "Tools - Add,Edit,Delete - User Management ");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"**** Parameters TCS fail executing****" + ex.Message);
                reportHelperS.LogTestStep(Status.Fail, "Tools - Add,Edit,Delete - User Management" + ex.Message);
                reportHelperS.LogScreenshot(Status.Fail, "Tools - Add,Edit,Delete - User Management", "Screenshot of failure", driver);
            }
        }
        [Test, Order(26)]
       // [Ignore("Ignoring for quick test")]
        public void RegressionTesting()
        {
            try
            {
                //PLC Connection
                //PDF Generation
                //SLC Parameter 
                //SLC Variable
                //Para def
                //Email Notification
                /****************************************************PDF Generation*****************************/
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);

                Thread.Sleep(5000);
                var windowHandles = driver.WindowHandles;
                for (int attempt = 0; attempt < 5 && windowHandles.Count > 1; attempt++)
                {
                    Thread.Sleep(TimeSpan.FromSeconds(1));
                    windowHandles = driver.WindowHandles;
                }

                if (windowHandles.Count >= 0)
                {
                    driver.SwitchTo().Window(windowHandles[0]);
                }

                WindowsElement folderCompareButton = driver.FindElementByName("Parameters");
                folderCompareButton.Click();
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                driver.FindElementByName("View Parameters").Click();

                /* var btnImport = driver.FindElementByName("Import");
                 btnImport.Click();

                 var windowHandles = driver.WindowHandles;
                 for (int attempt = 0; attempt < 5 && windowHandles.Count > 1; attempt++)
                 {
                     Thread.Sleep(TimeSpan.FromSeconds(1));
                     windowHandles = driver.WindowHandles;
                 }

                 // Switch session to the only application top level window that is still active (an indication that this is the main window)
                 if (windowHandles.Count >= 0)
                 {
                     driver.SwitchTo().Window(windowHandles[0]);
                 }
                 // Assert.AreEqual("MasterDB", driver.Title, false,
                 //    $"Actual title doesn't match expected title: {driver.Title}");
                 // Thread.Sleep(1000);

                 var combobox1 = driver.FindElementByAccessibilityId("ddlVersion");
                 // var combo = driver.FindElementByTagName("ComboBox");
                 // var open = combobox1.FindElementByName("Open");

                 var listItems = combobox1.FindElementsByTagName("ListItem");
                 Debug.WriteLine($"Before: Number of list items found: {listItems.Count}");

                 combobox1.Click();
                 combobox1.SendKeys(Keys.Down + Keys.Down + Keys.Down);
                 // open.Click();

                 listItems = combobox1.FindElementsByTagName("ListItem");
                 Debug.WriteLine($"After: Number of list items found: {listItems.Count}");

                 foreach (var comboKid in listItems)
                 {
                     if (comboKid.Text == "SW_8_30_6_2.83MW")
                     {
                         WebDriverWait wdv = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                         wdv.Until(x => comboKid.Displayed);

                         comboKid.Click();
                     }
                 }
                 var combobox = driver.FindElementByAccessibilityId("ddlComponentVersion");
                 var open3 = combobox.FindElementByName("Open");

                 var listItems3 = combobox.FindElementsByTagName("ListItem");
                 Debug.WriteLine($"Before: Number of list items found: {listItems3.Count}");


                 combobox.SendKeys(Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down);
                 open3.Click();

                 listItems = combobox.FindElementsByTagName("ListItem");
                 Debug.WriteLine($"After: Number of list items found: {listItems3.Count}");

                 foreach (var comboKid in listItems3)
                 {
                     if (comboKid.Text == "Batteries 1.0")
                     {
                         WebDriverWait wdv = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                         wdv.Until(x => comboKid.Displayed);

                         comboKid.Click();
                     }
                 }
                 Thread.Sleep(1000);
                 var btnselect = driver.FindElementByName("Select Parameter Row 0");
                 btnselect.Click();
                 Thread.Sleep(1000);
                 var btnselectt = driver.FindElementByName("Select Parameter Row 1");
                 btnselectt.Click();
                 var btnAdd = driver.FindElementByName("Add");
                 btnAdd.Click();
                 // Thread.Sleep(2000);
                 driver.FindElementByName("OK").Click();
                 //  Thread.Sleep(2000);
                 driver.FindElementByName("OK").Click();
                 // Thread.Sleep(2000);
                 // driver.FindElementByName("Cancel").Click();
                 //  var btnClose = driver.FindElementByName("Close");
                 // btnClose.Click();
                 // Thread.Sleep(1000);
                */

                //Export - Parameters
                var btnExport = driver.FindElementByName("Export");
                btnExport.Click();

                //ddlFormat
                var doctype_drpdwn2 = driver.FindElementByAccessibilityId("ddlFormat");
                doctype_drpdwn2.Click();
                doctype_drpdwn2.SendKeys(Keys.Down + Keys.Down);
                doctype_drpdwn2.Click();
                var btn_Export = driver.FindElementByAccessibilityId("btnExport");
                btn_Export.Click();
                var btn_Save = driver.FindElementByName("Save");
                btn_Save.Click();
                Thread.Sleep(2000);
                // driver.FindElementByName("Yes").Click();
                driver.FindElementByName("OK").Click();

                //Export -Status Code 
                driver.FindElementByName("Status Code").Click();
                var btExport = driver.FindElementByName("Export");
                btnExport.Click();

                //ddlFormat
                var doctype_drpdwn = driver.FindElementByAccessibilityId("ddlFormat");
                doctype_drpdwn.Click();
                doctype_drpdwn.SendKeys(Keys.Down + Keys.Down);
                doctype_drpdwn.Click();
                var btn_Export1 = driver.FindElementByAccessibilityId("btnExport");
                btn_Export1.Click();
                var btn_Save1 = driver.FindElementByName("Save");
                btn_Save1.Click();
                Thread.Sleep(2000);
                // driver.FindElementByName("Yes").Click();
                driver.FindElementByName("OK").Click();

                //Export -Variable List
                driver.FindElementByName("Variable List").Click();
                var Export = driver.FindElementByName("Export");
                Export.Click();

                //ddlFormat
                var doctyp_drpdwn = driver.FindElementByAccessibilityId("ddlFormat");
                doctyp_drpdwn.Click();
                doctyp_drpdwn.SendKeys(Keys.Down + Keys.Down);
                doctyp_drpdwn.Click();
                var butn_Export1 = driver.FindElementByAccessibilityId("btnExport");
                butn_Export1.Click();
                var butn_Save1 = driver.FindElementByName("Save");
                butn_Save1.Click();
                Thread.Sleep(2000);
                driver.FindElementByName("OK").Click();
                reportHelperS.LogTestStep(Status.Pass, "Export PDF");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"**** Parameters TCS fail executing****" + ex.Message);
                reportHelperS.LogTestStep(Status.Fail, "Export PDF" + ex.Message);
                reportHelperS.LogScreenshot(Status.Fail, "Regression Testing", "Screenshot of failure", driver);
            }

        }
        [Test, Order(27)]
      // [Ignore("Ignoring for quick test")]
        public void Email_Notification()
        {
            try
            {
                Thread.Sleep(5000);
                var windowHandles = driver.WindowHandles;
                for (int attempt = 0; attempt < 5 && windowHandles.Count > 1; attempt++)
                {
                    Thread.Sleep(TimeSpan.FromSeconds(1));
                    windowHandles = driver.WindowHandles;
                }

                if (windowHandles.Count >= 0)
                {
                    driver.SwitchTo().Window(windowHandles[0]);
                }
                WindowsElement folderCompareButton = driver.FindElementByName("Parameters");
                folderCompareButton.Click();
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                driver.FindElementByName("View Parameters").Click();
                //Main Controller
                // var radioFirst1 = driver.FindElementByName("Main Controller");
                // System.Diagnostics.Debug.WriteLine($"***** Value of radio First: {radioFirst1.Selected}");
                //  radioFirst1.Click();
                // Thread.Sleep(1000);
                var btnAdd = driver.FindElementByName("Add");
                btnAdd.Click();
                //Thread.Sleep(1000);
                // driver.FindElementByName("Open").Click();
                var combo = driver.FindElementByTagName("ComboBox");
                var open = combo.FindElementByName("Open");
                var listItems = combo.FindElementsByTagName("ListItem");
                Debug.WriteLine($"Before: Number of list items found: {listItems.Count}");
                combo.SendKeys(Keys.Down + Keys.Down);
                open.Click();
                listItems = combo.FindElementsByTagName("ListItem");
                Debug.WriteLine($"After: Number of list items found: {listItems.Count}");
                // maybe check number of elements in combo 
                //Assert.AreEqual(6, listItems.Count, "Combo box doesn't contain expected number of elements.");
                foreach (var comboKid in listItems)
                {
                    if (comboKid.Text == "Blade Sensors")
                    {
                        WebDriverWait wdv = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                        wdv.Until(x => comboKid.Displayed);
                        comboKid.Click();
                    }
                }
                var combo2 = driver.FindElementByAccessibilityId("ddlSystems");
                //  var combo2 = driver.FindElementByClassName("WindowsForms10.COMBOBOX.app.0.bb8560_r8_ad1");
                // var combo2 = driver.FindElementByTagName("ComboBox");
                var open2 = combo.FindElementByName("Open");
                var listItems2 = combo.FindElementsByTagName("ListItem");
                Debug.WriteLine($"Before: Number of list items found: {listItems2.Count}");
                combo2.SendKeys(Keys.Down + Keys.Down + Keys.Down);
                open2.Click();
                listItems2 = combo.FindElementsByTagName("ListItem");
                Debug.WriteLine($"After: Number of list items found: {listItems2.Count}");
                foreach (var comboKid in listItems2)
                {
                    if (comboKid.Text == "Gain correction")
                    {
                        WebDriverWait wdv = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                        wdv.Until(x => comboKid.Displayed);
                        comboKid.Click();
                    }
                }
                driver.FindElementByAccessibilityId("txtParameterName").SendKeys("VED");
                driver.FindElementByAccessibilityId("txtSystemName").SendKeys("VED System");
                driver.FindElementByAccessibilityId("txtMinValue").SendKeys("1");
                driver.FindElementByAccessibilityId("txtMaxValue").SendKeys("1000");
                driver.FindElementByAccessibilityId("txtValue").SendKeys("67");
                driver.FindElementByAccessibilityId("txtUnit").SendKeys("kW");
                driver.FindElementByName("kW").Click();
                driver.FindElementByAccessibilityId("txtDescription").SendKeys("Controls Develop services");
                driver.FindElementByName("Add").Click();
                driver.FindElementByName("Yes").Click();
                driver.FindElementByName("OK").Click();
                var btnEdit1 = driver.FindElementByName("Edit");
                btnEdit1.Click();
                driver.FindElementByName("OK").Click();
                driver.FindElementByName("Component Version Row 0").Click();
                var btnEdit = driver.FindElementByName("Edit");
                btnEdit.Click();
                driver.FindElementByAccessibilityId("txtMinValue").Clear();
                driver.FindElementByAccessibilityId("txtMinValue").SendKeys("45");
                var btnEditMax = driver.FindElementByAccessibilityId("txtMaxValue");
                btnEditMax.Click();
                driver.FindElementByName("OK").Click();
                btnEditMax.Clear();
                btnEditMax.SendKeys("10001");
                // driver.FindElementByName("OK").Click();
                // Thread.Sleep(1000);
                driver.FindElementByAccessibilityId("txtValue").Click();
                driver.FindElementByName("OK").Click();
                driver.FindElementByAccessibilityId("txtValue").Clear();
                driver.FindElementByAccessibilityId("txtValue").SendKeys("448");
                driver.FindElementByAccessibilityId("txtUnit").SendKeys("MW");
                //  driver.FindElementByName("Update").Click();
                // Thread.Sleep(2000);
                // driver.FindElementByName("OK").Click();
                // Thread.Sleep(2000);
                driver.FindElementByAccessibilityId("txtComments").SendKeys("Changing the Min and MAX value");
                // Thread.Sleep(1000);checkBoxIsCustomerVisibility
                // driver.FindElementByAccessibilityId("checkBoxIsCustomerVisibility").Click();
                driver.FindElementByName("Update").Click();
                // Thread.Sleep(2000);
                driver.FindElementByName("OK").Click();
                //  Thread.Sleep(2000);
                driver.FindElementByName("OK").Click();

                //Navigate to Change Request
                WindowsElement Button = driver.FindElementByName("Parameters");
                Button.Click();
                Button.SendKeys(Keys.Down + Keys.Down + Keys.Enter);
                Thread.Sleep(5000);
                var radioFirst2 = driver.FindElementByName("Hub Controller");
                radioFirst2.Click();
                var radioFirst1 = driver.FindElementByName("Main Controller");
                radioFirst1.Click();
                WindowsElement Filter = driver.FindElementByAccessibilityId("ddlFilters");
                Filter.Click();
                Filter.SendKeys(Keys.Down + Keys.Down + Keys.Enter);
                driver.FindElementByName("Controller Type Row 0").Click();
                driver.FindElementByName("Open").Click();
                // driver.FindElementByName("Close").Click();
                driver.FindElementByAccessibilityId("txtApproverComment").SendKeys("Approved");
                driver.FindElementByAccessibilityId("btnApprove").Click();
                driver.FindElementByName("Yes").Click();
                driver.FindElementByName("OK").Click();
                WindowsElement Filter1 = driver.FindElementByAccessibilityId("ddlFilters");
                Filter1.Click();
                Filter1.SendKeys(Keys.Down + Keys.Enter);
                //View Parameters
                WindowsElement Buttonn = driver.FindElementByName("Parameters");
                Buttonn.Click();
                Buttonn.SendKeys(Keys.Down + Keys.Enter);
                //View
                Actions actsTree = new Actions(driver);
                var nodeWorld = driver.FindElementByName("Component Version Row 0");
                actsTree.MoveToElement(nodeWorld);
                actsTree.DoubleClick();
                actsTree.Perform();
                // driver.FindElementByName("View Audit Logs").Click();
                // driver.FindElementByName("Close").Click();
                // Thread.Sleep(2000);
                // driver.FindElementByName("OK").Click();
                Thread.Sleep(1000);
                driver.FindElementByName("Close").Click();
                //Delete
                driver.FindElementByName("Component Version Row 0").Click();
                var btnDel = driver.FindElementByName("Delete");
                btnDel.Click();
                driver.FindElementByName("Yes").Click();
                driver.FindElementByName("OK").Click();

                reportHelperS.LogTestStep(Status.Pass, "Email_Notification - change request(max/min value) ");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"**** Parameters TCS fail executing****" + ex.Message);
                reportHelperS.LogTestStep(Status.Fail, "Email_Notification - change request(max/min value) " + ex.Message);
                reportHelperS.LogScreenshot(Status.Fail, "Email_Notification - change request(max/min value) ", "Screenshot of failure", driver);
            }
        }
        [Test, Order(28)]
       // [Ignore("Ignoring for quick test")]
        public void About()
        {
            try
            {
                driver.FindElementByName("About").Click();
                var About = driver.FindElementByName("About");
                About.Click();
                About.SendKeys(Keys.Alt + Keys.F4);
                //driver.FindElementByName("Close").Click();
                reportHelperS.LogTestStep(Status.Pass, "About - Version Details  ");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"**** Parameters TCS fail executing****" + ex.Message);
                reportHelperS.LogTestStep(Status.Fail, "About - Version Details " + ex.Message);
                reportHelperS.LogScreenshot(Status.Fail, "About - Version Details ", "Screenshot of failure", driver);
            }
        }
        

        //Audit logs
        [Test, Order(29)]
        // [Ignore("Ignoring for quick test")]
        public void NewEnhancments()
        {
            try
            {
                // New Enhancements
                WindowsElement twoXMButton = driver.FindElementByName("2XM");
                twoXMButton.Click();
                Thread.Sleep(8000);
                twoXMButton.SendKeys(Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.ArrowRight);
                Thread.Sleep(8000);
                twoXMButton.SendKeys(Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.ArrowRight);
                Thread.Sleep(8000);
                twoXMButton.SendKeys(Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Down + Keys.Enter);
               // twoXMButton.Click();

                reportHelperS.LogTestStep(Status.Pass, "New Enhancements - Para.Def Generator - DEF module");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"**** Parameters TCS fail executing****" + ex.Message);
                reportHelperS.LogTestStep(Status.Fail, "About - Version Details " + ex.Message);
                reportHelperS.LogScreenshot(Status.Fail, "About - Version Details ", "Screenshot of failure", driver);
            }
        }
        

        [TearDown]
        public void ReprotsFlushed()
        {
            System.Diagnostics.Debug.WriteLine("***** Finial Method called *****");
            // driver.Quit();
            reportHelperS.FlushReport();
        }
    }
}


