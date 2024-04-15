using DesktopAppTests.Helpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace DesktopAppTests.Tests
{
    [TestClass]
    public class NoteAppTest : DesktopAppTest
    {
        const string app = "Notepad";
        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {
            StartWinAppDriver();
            appSession = InitializeAppSession($@"{app.ToLower()}.exe", app);
        }

        [ClassCleanup]
        public static void ClassCleanup()
        {
            CloseApp(app.ToLower());
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
            string screenshotPath = ScreenPrinter.CaptureAndSaveScreenshot(appSession, (ScreenshotsDirectory + app));

            // Extract result from app using OCR
            string textExtracted = OCRTranslator.ExtractText(screenshotPath, 1, 50, 55, 20, 150);

            Assert.AreEqual(expectedResult, textExtracted);
        }
    }
}
