﻿using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using DesktopApp.Helpers;
using WindowsInput;

namespace DesktopApp.Tests
{
    [TestClass]
    public class AppGridTest
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
            appSession = InitializeAppSession(app);
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
            string expectedResult = "teste";

            // Insert values inside the app
            WriteTest(expectedResult);

            // Capture and save screenshot with timestamp
            string screenshotPath = ScreenPrinter.CaptureAndSaveScreenshot(appSession, ScreenshotsDirectory);

            // Extract result from app using OCR
            string textExtracted = OCRTranslator.ExtractText(screenshotPath, 0, 45, 32, 15, 150);

            Assert.AreEqual(expectedResult, textExtracted);
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

        private static void LaunchTestedApp(string appId)
        {
            Process.Start(appId);

            // Wait 2sec
            Thread.Sleep(2000);

            // Press Enter
            InputSimulator inputSimulator = new InputSimulator();
            inputSimulator.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.RETURN);

        }

        private static void CloseTestedApp()
        {
            Process[] processes = Process.GetProcessesByName(app.ToLower());
            foreach (Process process in processes)
            {
                process.Kill();
            }
        }

        private static WindowsDriver<WindowsElement> InitializeAppSession(string app)
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