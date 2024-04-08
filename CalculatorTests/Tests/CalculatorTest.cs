using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using CalculatorTests.Helpers;

namespace CalculatorTests.Tests
{
    [TestClass]
    public class CalculatorTest
    {
        private const string WinAppDriverPath = @"C:\Program Files (x86)\Windows Application Driver\WinAppDriver.exe";
        private const string CalculatorAppId = "Microsoft.WindowsCalculator_8wekyb3d8bbwe!App";
        private const string ScreenshotsDirectory = @"C:\Users\Starline\source\repos\CalculatorTests\CalculatorTests\Screenshots";

        [TestMethod]
        public void TestCalculator()
        {
            // Start WinAppDriver
            StartWinAppDriver();

            // Set up WinAppDriver connection and launch Calculator
            using (WindowsDriver<WindowsElement> calculatorSession = InitializeCalculatorSession())
            {
                //Define two integers (0~9)
                int num1 = 1;
                int num2 = 4;

                //Sum them
                int sum = num1 + num2;

                // Perform calculation
                AdditionTest(calculatorSession, num1, num2);

                // Capture and save screenshot with timestamp
                string screenshotPath = ScreenPrinter.CaptureAndSaveScreenshot(calculatorSession, ScreenshotsDirectory);

                // Define expected result
                string expectedResult = sum.ToString();

                //Extracts result from calculator, using OCR and ROI coordinates
                string calculatorResult = OCRTranslator.ExtractText(screenshotPath, 315, 167, 55, 55); 
             
                // Assert the result
                Assert.Equals(expectedResult, calculatorResult);
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

        private WindowsDriver<WindowsElement> InitializeCalculatorSession()
        {
            AppiumOptions appCapabilities = new AppiumOptions();
            appCapabilities.AddAdditionalCapability("app", CalculatorAppId);
            WindowsDriver<WindowsElement> calculatorSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);
            return calculatorSession;
        }

        private void AdditionTest(WindowsDriver<WindowsElement> calculatorSession, int num1, int num2)
        {
            // Wait for the calculator to load
            Thread.Sleep(2000);

            // Enter num1 + num2 in the calculator
            calculatorSession.FindElementByAccessibilityId($"num{num1}Button").Click();
            calculatorSession.FindElementByAccessibilityId("plusButton").Click();
            calculatorSession.FindElementByAccessibilityId($"num{num2}Button").Click();
            calculatorSession.FindElementByAccessibilityId("equalButton").Click();
        }
    }
}
