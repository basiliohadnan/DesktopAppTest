using Consinco.Helpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;

namespace Consinco.MaxCompra
{
    public class MaxCompraInit : WinAppDriver
    {
        const string user = "PS032528";

        public MaxCompraInit(IWebDriver driver) : base(driver)
        {
        }

        [TestInitialize]
        public static void ClassInitialize(TestContext context)
        {
            Global.app = "MaxCompra";
            StartWinAppDriver();
            var appSession = InitializeAppSession(@$"C:\C5Client\Max\{Global.app}.exe", Global.app);
        }

        public void Login()
        {
            // Define the bounding rectangle of the field "Usuário"
            var userFieldRect = new ElementHandler.BoundingRectangle(390, 217, 514, 237);

            // Call the method to click on the item
            ClickOn(appSession, userFieldRect);

            // Fill username field with the information provided
            FillField(appSession, user);

            PressEnter();

            PressEnter();

            // Create a window manager instance
            var windowManager = new WinAppDriver(appSession); // Pass appSession to the constructor
            WaitSeconds(2);

            // Switch to the window with the specified title
            windowManager.SwitchToWindowWithTitle("acrux mercari - Compras");

            // Capture and save screenshot with test method name
            ScreenPrinter.CaptureAndSaveScreenshot(appSession, ScreenshotsDirectory + "\\" + Global.app + "_01-Login.png");
        }

        public void OpenMenuItem(string menuItemName, string screenshotName)
        {
            var menuItem = appSession.FindElementByName(menuItemName);
            menuItem.Click();
            ScreenPrinter.CaptureAndSaveScreenshot(appSession, ScreenshotsDirectory + "\\" + Global.app + "_" + screenshotName + ".png");
        }

    }
}
