using Consinco.Helpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;

namespace Consinco.Tests
{
    [TestClass]
    public class NoteAppTest : WinAppDriver
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

        protected void WriteTest(string text)
        {
            // Wait for the app to load
            Thread.Sleep(500);

            // Clear the content of the app by selecting all text and then deleting it
            appSession.FindElementByClassName("Edit").SendKeys(Keys.Control + "a");
            appSession.FindElementByClassName("Edit").SendKeys(Keys.Delete);

            // Enter values in app
            appSession.FindElementByClassName("Edit").SendKeys(text);
            // Press Enter
            appSession.FindElementByClassName("Edit").SendKeys(Keys.Enter);
        }
    }
}
