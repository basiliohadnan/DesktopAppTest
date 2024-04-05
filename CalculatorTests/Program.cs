using System;
using System.IO;
using System.Threading;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using Tesseract;

namespace CalculatorTest
{
    class Program
    {
        static void Main(string[] args)
        {
            // Set up WinAppDriver connection
            WindowsDriver<WindowsElement> calculatorSession = null;
            AppiumOptions appCapabilities = new AppiumOptions();
            appCapabilities.AddAdditionalCapability("app", "Microsoft.WindowsCalculator_8wekyb3d8bbwe!App");
            calculatorSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);

            // Wait for the calculator to load
            Thread.Sleep(2000);

            // Enter '2 + 2' in the calculator
            calculatorSession.FindElementByAccessibilityId("num2Button").Click(); 
            calculatorSession.FindElementByAccessibilityId("plusButton").Click(); 
            calculatorSession.FindElementByAccessibilityId("num2Button").Click(); 
            calculatorSession.FindElementByAccessibilityId("equalButton").Click();


            // Capture and read the result using OCR
            var screenshot = calculatorSession.GetScreenshot();
            string screenshotFilePath = "screenshot.png";
            screenshot.SaveAsFile(screenshotFilePath);
            string result = ReadTextFromImage(screenshotFilePath);

            // Output the result
            Console.WriteLine("Result: " + result);

            // Close the calculator
            calculatorSession.Close();
        }

        // Method to extract text from image using OCR
        static string ReadTextFromImage(string imagePath)
        {
            using (var engine = new TesseractEngine(@"C:\Users\Starline\source\repos\CalculatorTests\tessdata", "eng", EngineMode.Default))
            {
                using (var img = Pix.LoadFromFile(imagePath))
                {
                    using (var page = engine.Process(img))
                    {
                        return page.GetText();
                    }
                }
            }
        }
    }
}
