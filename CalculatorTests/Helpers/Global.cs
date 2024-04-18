using OpenQA.Selenium.Appium.Windows;

namespace Consinco.Helpers
{
    public class Global
    {
        public static WindowsDriver<WindowsElement> winSession;
        public static WindowsElement mainElement;
        public static WindowsDriver<WindowsElement> appSession;
        public static string app;
        public static string screenshotsDirectory;
    }
}
