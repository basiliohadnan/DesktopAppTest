using Consinco.Helpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium.Windows;

namespace Consinco.MaxCompra
{
    public class MaxCompraInit : WinAppDriver
    {
        const string user = "PS032528";
        const string app = "MaxCompra";
        const string appName = "acrux mercari - Compras";
        protected static WindowsDriver<WindowsElement> appSession;

        [TestInitialize]
        public void Initialize()
        {
            Global.app = app;
            StartWinAppDriver();
            InitializeWinSession();
            InitializeAppSession(@$"C:\C5Client\Max\{Global.app}.exe");
        }

        public void Login()
        {
            // Define the bounding rectangle of the field "Usuário"
            var userFieldRect = new ElementHandler.BoundingRectangle(390, 217, 514, 237);

            // Call the method to click on the item
            ClickOn(userFieldRect);

            // Fill username field with the information provided
            FillField(user);

            PressEnter();

            PressEnter();

            // Set AppSession using classname
            SetAppSession("Centura:MDIFrame");

            // Capture and save screenshot with test method name
            ScreenPrinter.CaptureAndSaveScreenshot(Global.appSession, screenshotsDirectory + "\\" + Global.app + "\\" + "01-Login");

        }

        public void OpenMenu(string menuItemName, string testName)
        {
            var menuItem = Global.appSession.FindElementByName(menuItemName);
            menuItem.Click();
            ScreenPrinter.CaptureAndSaveScreenshot(Global.appSession, screenshotsDirectory + "\\" + Global.app + "\\" + testName);
        }
    }
}
