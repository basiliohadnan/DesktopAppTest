using DesktopApp.Helpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace DesktopApp.Tests
{
    [TestClass]
    public class NoteAppTest : DesktopAppTest
    {
        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {
            StartWinAppDriver();
            appSession = InitializeAppSession(@"notepad.exe", "Notepad");
        }

        [ClassCleanup]
        public static void ClassCleanup()
        {
            CloseApp("notepad");
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
            string textExtracted = OCRTranslator.ExtractText(screenshotPath, 1, 49, 30, 15, 150);

            Assert.AreEqual(expectedResult, textExtracted);
        }
    }
}
