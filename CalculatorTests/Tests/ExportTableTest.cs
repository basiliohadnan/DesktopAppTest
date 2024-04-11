using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using DesktopApp.Helpers;
using WindowsInput;

namespace DesktopApp.Tests
{
    [TestClass]
    public class ExportTableTest
    {
        private const string app = "ExportTable";
        private const string appPath = $@"C:\\Users\\Starline\\Documents\\AppGrid\{app}.exe";
        private const string ScreenshotsDirectory = $@"C:\Users\Starline\source\repos\DesktopAppTest\CalculatorTests\Screenshots\{app}";
        private const string WinAppDriverPath = @"C:\Program Files (x86)\Windows Application Driver\WinAppDriver.exe";
        private static WindowsDriver<WindowsElement> appSession;

        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {
            StartWinAppDriver();
            LaunchTestedApp(appPath);
            //appSession = InitializeAppSession(appPath);
        }

        [ClassCleanup]
        public static void ClassCleanup()
        {
            CloseTestedApp();
            appSession?.Quit();
            StopWinAppDriver();
        }

        [TestMethod]
        public void ValidatesTextInserted()
        {
            // Define expected result
            //string expectedResult = "teste";

            //// Insert values inside the app
            //WriteTest(expectedResult);

            // Capture and save screenshot with timestamp
            string screenshotPath = ScreenPrinter.CaptureAndSaveScreenshot(appSession, ScreenshotsDirectory);

            //// Extract result from app using OCR
            //string textExtracted = OCRTranslator.ExtractText(screenshotPath, 0, 45, 32, 15, 150);

            //Assert.AreEqual(expectedResult, textExtracted);
        }

        private static void StartWinAppDriver()
        {
            // Start WinAppDriver process if not already running
            Process[] processes = Process.GetProcessesByName("WinAppDriver");
            if (processes.Length == 0)
            {
                Process.Start(WinAppDriverPath);
                // Wait for WinAppDriver to start
                Thread.Sleep(5000);
            }
        }

        private static void StopWinAppDriver()
        {
            // Stop WinAppDriver process
            Process[] processes = Process.GetProcessesByName("WinAppDriver");
            foreach (Process process in processes)
            {
                process.Kill();
            }
        }

        private static WindowsDriver<WindowsElement> LaunchTestedApp(string appPath)
        {
            Process.Start(appPath);

            // Wait 2sec
            Thread.Sleep(2000);

            // Press Enter
            InputSimulator inputSimulator = new InputSimulator();
            inputSimulator.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.RETURN);
            Thread.Sleep(2000);

            // Create session at root level
            AppiumOptions rootCapabilities = new AppiumOptions();
            rootCapabilities.AddAdditionalCapability("app", "Root");
            WindowsDriver<WindowsElement> winSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), rootCapabilities);

            // Launch the application using the provided appPath
            AppiumOptions appCapabilities = new AppiumOptions();
            appCapabilities.AddAdditionalCapability("app", appPath);
            WindowsDriver<WindowsElement> appSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);

            return appSession;

        }

        private static void CloseTestedApp()
        {
            Process[] processes = Process.GetProcessesByName(app.ToLower());
            foreach (Process process in processes)
            {
                process.Kill();
            }
        }

        private static WindowsDriver<WindowsElement> InitializeAppSession(string appPath)
        {
            // Create session at root level
            AppiumOptions rootCapabilities = new AppiumOptions();
            rootCapabilities.AddAdditionalCapability("app", "Root");
            WindowsDriver<WindowsElement> winSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), rootCapabilities);

            // Launch the application using the provided appPath
            AppiumOptions appCapabilities = new AppiumOptions();
            appCapabilities.AddAdditionalCapability("app", appPath);
            WindowsDriver<WindowsElement> appSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);
            
            // Wait 2sec
            Thread.Sleep(2000);

            // Press Enter
            InputSimulator inputSimulator = new InputSimulator();
            inputSimulator.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.RETURN);

            // Wait 3sec
            Thread.Sleep(3000);

            return appSession;
        }

        private void WriteTest(string text)
        {
            // Wait for the app to load
            Thread.Sleep(500);

            // Clear the content of the app by selecting all text and then deleting it
            appSession.FindElementByClassName("Edit").SendKeys(Keys.Control + "a");
            appSession.FindElementByClassName("Edit").SendKeys(Keys.Delete);

            // Enter values in app
            appSession.FindElementByClassName("Edit").SendKeys(text);
            // Press Enter
            appSession.FindElementByClassName("Edit").SendKeys(Keys.Enter);
        }
    }
}
