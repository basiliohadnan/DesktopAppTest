using OpenQA.Selenium;
using OpenQA.Selenium.Appium.Windows;

namespace DesktopAppTests.Helpers
{
    public class ElementHandler
    {
        public static WindowsElement FindElementByClassName(WindowsDriver<WindowsElement> appSession, string className)
        {
            int attempts = 0;
            WindowsElement element = null;
            while (attempts < 10)
            {
                try
                {
                    element = appSession.FindElementByClassName(className);
                    if (element != null)
                        break;
                }
                catch (NoSuchElementException)
                {
                    // Element not found, continue trying
                }
                attempts++;
            }
            return null;
        }
    }
}
