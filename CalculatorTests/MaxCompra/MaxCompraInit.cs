using Consinco.Helpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium.Windows;

namespace Consinco.MaxCompra
{
    public class MaxCompraInit : WinAppDriver
    {
        const string user = "PS032528";
        const string app = "MaxCompra";
        //const string appName = "acrux mercari - Compras";
        protected static WindowsDriver<WindowsElement> appSession;
        private ElementHandler elementHandler;

        public MaxCompraInit()
        {
            elementHandler = new ElementHandler();
        }

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
            FillField(user);
            PressEnter();
            PressEnter();

            // Set AppSession using classname
            SetAppSession("Centura:MDIFrame");
            ScreenPrinter.CaptureAndSaveScreenshot(Global.screenshotsDirectory, "01-Login");
        }

        public void OpenMenu(string menuItemName, string testName)
        {
            var menuItem = elementHandler.FindElementByName(menuItemName);
            menuItem.Click();
            ScreenPrinter.CaptureAndSaveScreenshot(Global.screenshotsDirectory, testName);
        }
    }
}
