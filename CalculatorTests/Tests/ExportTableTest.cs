using DesktopAppTests.Helpers;
using DesktopAppTests.Tests;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace DesktopAppTests.Tests
{
    [TestClass]
    public class ExportTableTest : DesktopAppTest
    {
        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {
            StartWinAppDriver();
            appSession = InitializeAppSession(@"C:\Users\Starline\Documents\AppGrid\ExportTable.exe", "ExportTable");
        }

        [ClassCleanup]
        public static void ClassCleanup()
        {
            CloseApp("ExportTable");
            appSession?.Quit();
            StopWinAppDriver();
        }

        [TestMethod]
        public void ValidatesTextInserted()
        {
            string expectedResult = "String1";

            // Capture and save screenshot with timestamp
            string screenshotPath = ScreenPrinter.CaptureAndSaveScreenshot(appSession, ScreenshotsDirectory);

            string textExtracted = OCRTranslator.ExtractText(screenshotPath, 109, 197, 68, 15, 165);

            Assert.AreEqual(expectedResult, textExtracted);
        }
    }
}
