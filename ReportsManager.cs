using AventStack.ExtentReports.Reporter;
using AventStack.ExtentReports;
using OpenQA.Selenium.Appium.Windows;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;

namespace SanityTest
{
    public class ReportsManager
    {
        private ExtentReports extent = new AventStack.ExtentReports.ExtentReports();
        private ExtentTest test;
        //  static WindowsDriver<WindowsElement> driver;
        public static String dir = AppDomain.CurrentDomain.BaseDirectory;
        public static String testResultPath = dir.Replace("bin\\Debug\\", "SanityTestResults");

        public void ExtentManager()
        {
            var img = new DirectoryInfo(System.IO.Directory.GetParent(@"../../../").FullName + Path.DirectorySeparatorChar
              + "Screenshots");
            FileInfo[] files = img.GetFiles();
            foreach (FileInfo file in files)
            {
                file.Delete();
            }
            string nameOS = Environment.OSVersion.ToString();
            string platform1 = Environment.OSVersion.Platform.ToString();
            System.Diagnostics.Debug.WriteLine("****** Reports started ******" + testResultPath);
            ExtentHtmlReporter htmlreport = new ExtentHtmlReporter(testResultPath);
            extent.AttachReporter(htmlreport);
            extent.AddSystemInfo("Application", "MasterDB");
            extent.AddSystemInfo("System", nameOS);
            extent.AddSystemInfo("Platform", platform1);
            extent.AddSystemInfo("Environment", "QA");
            extent.AddSystemInfo("Windows Automation", "Appium Winappdriver C#");
            test = extent.CreateTest("MasterDB Application - Reports with Screenshots");
            System.Diagnostics.Debug.WriteLine("****** Reports completed ****** testResultPath ********" + testResultPath);
        }

        public void CreateTest(string testName)
        {
            test = extent.CreateTest(testName);
        }

        public void LogTestStepPass(Status status, string stepName)
        {

            test.Log(status, stepName);

        }
        public void LogTestStep(Status status, string stepName)
        {
            System.Diagnostics.Debug.WriteLine($"*****Log test executing****");
            test.Log(status, stepName);
        }
        public void LogTestStepError(Status status)
        {
            System.Diagnostics.Debug.WriteLine($"*****Log test error executing****");
            test.Log(status);
        }
        public void LogScreenshot(Status status, string stepName, string details, WindowsDriver<WindowsElement> driver)
        {

            // test.Log(status, "LogScreenshot executing" + ex);
            System.Diagnostics.Debug.WriteLine($"*****LogScreenshot executing****" + driver);
            try
            {
                System.Diagnostics.Debug.WriteLine($"*****LogScreenshot executing inside try block****");
              //************************************************
                string screenshotpath = System.IO.Directory.GetParent(@"../../../").FullName + Path.DirectorySeparatorChar
           + "Screenshots" + Path.DirectorySeparatorChar + "Screenshots" + DateTime.Now.ToString("ddMMyyyy HHmmss");

                var screenShot = driver.GetScreenshot();
                System.Diagnostics.Debug.WriteLine($"*****LogScreenshot executing Got screenshot path****" + screenshotpath);
                screenShot.SaveAsFile(screenshotpath + ".png", ScreenshotImageFormat.Png);
                System.Diagnostics.Debug.WriteLine($"*****LogScreenshot executing File saving****" + screenshotpath);
                string path = screenshotpath + ".png";

                //    test.Log(status, stepName + " " + test.AddScreenCaptureFromPath(path));
                test.AddScreenCaptureFromPath(path);
                System.Diagnostics.Debug.WriteLine($"*****LogScreenshot executed");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"*****LogScreenshot executing exception****" + ex.Message);
                test.Log(status, "LogScreenshot" + " " + ex.Message);
            }
        }
        public void FlushReport()
        {
            System.Diagnostics.Debug.WriteLine($"*****Flush reports called****");
            extent.Flush();
        }
    }
}
