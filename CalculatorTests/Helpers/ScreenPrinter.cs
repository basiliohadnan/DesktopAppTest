using OpenQA.Selenium;
using System.Drawing.Imaging;
using System.Drawing;
using System.Windows;

namespace Consinco.Helpers
{
    public class ScreenPrinter
    {
        public static string CaptureAndSaveScreenshot(string directoryPath, string testName)
        {
            // Generate filename with timestamp
            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
            string filename = $"screenshot_{timestamp}.png";

            // Capture screenshot using ITakesScreenshot interface
            var screenshot = ((ITakesScreenshot)Global.appSession).GetScreenshot();

            // Ensure the directory exists or create it
            directoryPath = directoryPath + Global.app + "\\" + testName; 
            Directory.CreateDirectory(directoryPath);

            // Save screenshot to a file
            string screenshotPath = Path.Combine(directoryPath, filename);
            screenshot.SaveAsFile(screenshotPath, ScreenshotImageFormat.Png);

            // Output the result
            Console.WriteLine($"Screenshot saved to: {screenshotPath}");

            return screenshotPath;
        }
        public static string CustomPrintScreen(string filename, bool full = true, int sleep = 0, int left = 0, int top = 0, int width = 0, int height = 0)
        {
            try
            {
                if (sleep > 0)
                {
                    Thread.Sleep(sleep * 1000);
                }

                if (File.Exists(filename))
                {
                    File.Delete(filename);
                }

                int widthPng;
                int heightPng;

                if (full)
                {
                    widthPng = (int)SystemParameters.FullPrimaryScreenWidth;
                    heightPng = (int)SystemParameters.FullPrimaryScreenHeight;
                }
                else
                {
                    widthPng = width;
                    heightPng = height;
                }

                var capturePng = new Bitmap(widthPng, heightPng, PixelFormat.Format32bppArgb);
                using (Graphics captureGraphic = Graphics.FromImage(capturePng))
                {
                    captureGraphic.CopyFromScreen(left, top, 0, 0, capturePng.Size);
                    capturePng.Save(filename, ImageFormat.Png);
                }

                return filename;
            }
            catch (Exception ex)
            {
                Global.processTest.Print("Exception at ProcessTest.Print", ex);
                return null;
            }
        }

    }
}
