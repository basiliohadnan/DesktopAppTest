using OpenQA.Selenium;
using OpenQA.Selenium.Appium.Windows;

namespace CalculatorTests.Helpers
{
    public class ScreenPrinter
    {
        public static void CaptureAndSaveScreenshot(WindowsDriver<WindowsElement> calculatorSession, string directoryPath, string filename)
        {
            // Capture screenshot using ITakesScreenshot interface
            var screenshot = ((ITakesScreenshot)calculatorSession).GetScreenshot();

            // Create the directory if it doesn't exist
            if (!Directory.Exists(directoryPath))
            {
                Directory.CreateDirectory(directoryPath);
            }

            // Save screenshot to a file
            string screenshotPath = Path.Combine(directoryPath, filename);
            screenshot.SaveAsFile(screenshotPath, ScreenshotImageFormat.Png);

            // Output the result
            Console.WriteLine("Screenshot saved to: " + screenshotPath);
        }

        public static void CaptureAndSaveScreenshotWithTimestamp(WindowsDriver<WindowsElement> calculatorSession, string directoryPath)
        {
            // Generate filename with timestamp
            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
            string filename = $"screenshot_{timestamp}.png";

            // Call the CaptureAndSaveScreenshot method with the generated filename
            CaptureAndSaveScreenshot(calculatorSession, directoryPath, filename);
        }
    }
}
