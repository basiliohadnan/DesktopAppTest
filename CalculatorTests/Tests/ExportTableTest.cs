using DesktopAppTests.Helpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium.MultiTouch;
using OpenQA.Selenium.Appium.Windows;

namespace DesktopAppTests.Tests
{
    [TestClass]
    public class ExportTableTest : DesktopAppTest
    {
        const string app = "ExportTable";
        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {
            StartWinAppDriver();
            appSession = InitializeAppSession(@$"C:\Users\Starline\Documents\AppGrid\{app}.exe", app);
        }

        [ClassCleanup]
        public static void ClassCleanup()
        {
            CloseApp(app);
            appSession?.Quit();
            StopWinAppDriver();
        }

        [TestMethod]
        public void ValidatesTextInserted()
        {
            string expectedResult = "String1";

            // Capture and save screenshot with timestamp
            string screenshotPath = ScreenPrinter.CaptureAndSaveScreenshot(appSession, (ScreenshotsDirectory + "/" + app));

            string textExtracted = OCRTranslator.ExtractText(screenshotPath, 109, 197, 68, 15, 165);

            Assert.AreEqual(expectedResult, textExtracted);
        }

        [TestMethod]
        public void ClickAndVerifyEditableCellByLocation()
        {
            int left = 442;
            int top = 408;
            int right = 513;
            int bottom = 422;

            // Press the cell to give it focus
            TouchAction touchAction = new TouchAction(appSession);
            int centerX = (left + right) / 2;
            int centerY = (top + bottom) / 2;
            touchAction.Press(centerX, centerY).Release().Perform();

            // Find the cell element by class name
            WindowsElement cell = FindCellByClassName("Edit");

            // Verify if the cell is editable
            bool isEditable = IsCellEditable(cell);

            Assert.IsTrue(isEditable, "The cell is not editable after clicking.");
        }

        private WindowsElement FindCellByClassName(string className)
        {
            // Use FindElementByClassName method to locate the element by class name
            try
            {
                WindowsElement cell = appSession.FindElementByClassName(className);
                return cell;
            }
            catch (NoSuchElementException)
            {
                throw new Exception($"Cell element with class name '{className}' not found.");
            }
        }

        private bool IsCellEditable(WindowsElement cell)
        {
            // Send some keystrokes to the cell to simulate editing
            cell.SendKeys("Test");

            // Check if the cell value has been updated after sending keys
            return cell.Text.Contains("Test");
        }
    }
}
