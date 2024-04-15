using DesktopAppTests.Helpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace DesktopAppTests.Tests
{
    [TestClass]
    public class MaxCompraTest : DesktopAppTest
    {
        const string app = "MaxCompra";
        const string user = "PS032528";

        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {
            StartWinAppDriver();
            appSession = InitializeAppSession(@$"C:\C5Client\Max\{app}.exe", app);
        }

        [ClassCleanup]
        public static void ClassCleanup()
        {
            CloseApp(app);
            appSession?.Quit();
            StopWinAppDriver();
        }

        [TestMethod]
        public void Login()
        {
            // Define the bounding rectangle of the field "Usuário"
            var userFieldRect = new ElementHandler.BoundingRectangle(390, 217, 514, 237);

            // Call the method to click on the item
            ClickOnItem(userFieldRect);

            // Fill username field with the information provided
            FillItemWithInformation(user);
            
            PressEnter();

            WaitSeconds(2);

            PressEnter();

            // Result
            //ScreenPrinter.CaptureAndSaveScreenshot(appSession, ScreenshotsDirectory + "/" + app);

        }
    }
}
