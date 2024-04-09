using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using DesktopApp.Helpers;

namespace DesktopApp.Tests
{
    [TestClass]
    public class AppTest
    {
        private const string WinAppDriverPath = @"C:\Program Files (x86)\Windows Application Driver\WinAppDriver.exe";
        private const string ScreenshotsDirectory = @"C:\Users\Starline\source\repos\CalculatorTests\CalculatorTests\Screenshots";

        [TestMethod]
        public void ValidatesTextInserted()
        {
            // Start WinAppDriver
            StartWinAppDriver();

            // Initialize App Session
            string app = "Notepad";
            using (WindowsDriver<WindowsElement> appSession = InitializeAppSession(app))
            {
                // Define expected result
                string expectedResult = "teste 0123456789 0.42,8.9,7:3/8* {ENTER} ";

                // Insert values inside the app
                WriteTest(appSession, expectedResult);

                // Capture and save screenshot with timestamp
                string screenshotPath = ScreenPrinter.CaptureAndSaveScreenshot(appSession, ScreenshotsDirectory);


                //Extracts result from calculator, using OCR and ROI coordinates
                string textExtracted = OCRTranslator.ExtractText(screenshotPath, 0, 45, 285, 20);
            
                Assert.AreEqual(expectedResult, textExtracted);
            }

            // Stop WinAppDriver
            StopWinAppDriver();
        }

        private void StartWinAppDriver()
        {
            // Start WinAppDriver process if not already running
            Process[] processes = Process.GetProcessesByName("WinAppDriver");
            if (processes.Length == 0)
            {
                Process.Start(WinAppDriverPath);
                // Wait for WinAppDriver to start
                Thread.Sleep(5000); // Adjust sleep time as needed
            }
        }

        private void StopWinAppDriver()
        {
            // Stop WinAppDriver process
            Process[] processes = Process.GetProcessesByName("WinAppDriver");
            foreach (Process process in processes)
            {
                process.Kill();
            }
        }

        private WindowsDriver<WindowsElement> InitializeAppSession(string app)
        {
            // Create session at root level
            AppiumOptions rootCapabilities = new AppiumOptions();
            rootCapabilities.AddAdditionalCapability("app", "Root");
            WindowsDriver<WindowsElement> winSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), rootCapabilities);

            var RootWindow = winSession.FindElementByClassName(app);
            // Create session by attaching to App top level window
            AppiumOptions appCapabilities = new AppiumOptions();
            var RootTopLevelWindowHandle = RootWindow.GetAttribute("NativeWindowHandle");
            RootTopLevelWindowHandle = (int.Parse(RootTopLevelWindowHandle)).ToString("x"); // Convert to Hex
            appCapabilities.AddAdditionalCapability("appTopLevelWindow", RootTopLevelWindowHandle);
            WindowsDriver<WindowsElement> appSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);

            // Maximize Window
            appSession.Manage().Window.Maximize();

            return appSession;
        }

        //Modify it after undestanding
        private void WriteTest(WindowsDriver<WindowsElement> appSession, string text)
        {
            // Wait for the calculator to load
            Thread.Sleep(500);

            // Enter num1 + num2 in the calculator
            appSession.FindElementByClassName("Edit").SendKeys(text);
        }
    }
}
