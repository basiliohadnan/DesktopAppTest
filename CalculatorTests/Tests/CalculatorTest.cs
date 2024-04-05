using System;
using System.IO;
using System.Threading;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;

namespace CalculatorTests.Tests
{
    public class CalculatorTest
    {
        public void RunTest()
        {
            // Set up WinAppDriver connection
            using (WindowsDriver<WindowsElement> calculatorSession = InitializeCalculatorSession())
            {
                // Wait for the calculator to load
                Thread.Sleep(2000);

                // Perform calculation
                AdditionTest(calculatorSession);

                // Capture and read the result using OCR
                string result = OCRHelper.ReadTextFromImage("screenshot.png");

                // Output the result
                Console.WriteLine("Result: " + result);
            }
        }

        private WindowsDriver<WindowsElement> InitializeCalculatorSession()
        {
            WindowsDriver<WindowsElement> calculatorSession = null;
            AppiumOptions appCapabilities = new AppiumOptions();
            appCapabilities.AddAdditionalCapability("app", "Microsoft.WindowsCalculator_8wekyb3d8bbwe!App");
            calculatorSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);
            return calculatorSession;
        }

        private void AdditionTest(WindowsDriver<WindowsElement> calculatorSession)
        {
            // Enter '2 + 2' in the calculator
            calculatorSession.FindElementByAccessibilityId("num2Button").Click();
            calculatorSession.FindElementByAccessibilityId("plusButton").Click();
            calculatorSession.FindElementByAccessibilityId("num2Button").Click();
            calculatorSession.FindElementByAccessibilityId("equalButton").Click();
        }
    }
}
