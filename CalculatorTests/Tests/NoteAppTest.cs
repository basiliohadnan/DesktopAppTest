using Consinco.Helpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;

namespace Consinco.Tests
{
    [TestClass]
    public class NoteAppTest : WinAppDriver
    {
        const string app = "Notepad";
        const string appPath = "pat for Notepad";

        [ClassInitialize]
        public void ClassInitialize(TestContext context)
        {
            StartWinAppDriver();
            InitializeAppSession(appPath);
        }

        [ClassCleanup]
        public void ClassCleanup()
        {
            CloseApp(app.ToLower());
            Global.appSession?.Quit();
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
            string screenshotPath = Global.processTest.CaptureWholeScreen();

            // Extract result from app using OCR
            string textExtracted = OCRScanner.ExtractText(screenshotPath, 1, 50, 55, 20, 150);

            Assert.AreEqual(expectedResult, textExtracted);
        }

        protected void WriteTest(string text)
        {
            // Wait for the app to load
            Thread.Sleep(500);

            // Clear the content of the app by selecting all text and then deleting it
            Global.appSession.FindElementByClassName("Edit").SendKeys(Keys.Control + "a");
            Global.appSession.FindElementByClassName("Edit").SendKeys(Keys.Delete);

            // Enter values in app
            Global.appSession.FindElementByClassName("Edit").SendKeys(text);
            // Press Enter
            Global.appSession.FindElementByClassName("Edit").SendKeys(Keys.Enter);
        }
    }
}
