using OpenQA.Selenium;
using OpenQA.Selenium.Appium.Windows;

namespace Consinco.Helpers
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

        public struct BoundingRectangle
        {
            public int Left { get; set; }
            public int Top { get; set; }
            public int Right { get; set; }
            public int Bottom { get; set; }

            public BoundingRectangle(int left, int top, int right, int bottom)
            {
                Left = left;
                Top = top;
                Right = right;
                Bottom = bottom;
            }
        }
    }
}
