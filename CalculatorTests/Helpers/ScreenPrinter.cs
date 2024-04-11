using OpenQA.Selenium;
using OpenQA.Selenium.Appium.Windows;

namespace DesktopAppTests.Helpers
{
    public class ScreenPrinter
    {
        public static string CaptureAndSaveScreenshot(WindowsDriver<WindowsElement> appSession, string directoryPath)
        {
            // Generate filename with timestamp
            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
            string filename = $"screenshot_{timestamp}.png";

            // Capture screenshot using ITakesScreenshot interface
            var screenshot = ((ITakesScreenshot)appSession).GetScreenshot();

            // Ensure the directory exists or create it
            Directory.CreateDirectory(directoryPath);

            // Save screenshot to a file
            string screenshotPath = Path.Combine(directoryPath, filename);
            screenshot.SaveAsFile(screenshotPath, ScreenshotImageFormat.Png);

            // Output the result
            Console.WriteLine($"Screenshot saved to: {screenshotPath}");

            return screenshotPath;
        }
    }
}
