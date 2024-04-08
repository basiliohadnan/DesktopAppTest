using System.Diagnostics;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using CalculatorTests.Helpers;


namespace CalculatorTests.Tests
{
    public class CalculatorTest
    {
        private const string WinAppDriverPath = @"C:\Program Files (x86)\Windows Application Driver\WinAppDriver.exe";
        private const string CalculatorAppId = "Microsoft.WindowsCalculator_8wekyb3d8bbwe!App";

        public void RunTest()
        {
            // Start WinAppDriver
            StartWinAppDriver();

            // Set up WinAppDriver connection and launch Calculator
            using (WindowsDriver<WindowsElement> calculatorSession = InitializeCalculatorSession())
            {
                // Perform calculation
                AdditionTest(calculatorSession);

                // Define the directory path
                string directoryPath = "C:\\Users\\Starline\\source\\repos\\CalculatorTests\\CalculatorTests\\Screenshots";

                // Capture and save screenshot with timestamp
                ScreenPrinter.CaptureAndSaveScreenshotWithTimestamp(calculatorSession, directoryPath);
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

        private void AdditionTest(WindowsDriver<WindowsElement> calculatorSession, int num1 = 3, int num2 = 6)
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
