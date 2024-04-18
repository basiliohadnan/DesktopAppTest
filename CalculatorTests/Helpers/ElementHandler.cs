using OpenQA.Selenium;
using OpenQA.Selenium.Appium.Windows;
using System.Collections.ObjectModel;

namespace Consinco.Helpers
{
    public class ElementHandler
    {
        public ReadOnlyCollection<WindowsElement> FindElementsByClassName(string className)
        {
            const int maxAttempts = 10;
            int attempts = 0;

            while (attempts < maxAttempts)
            {
                try
                {
                    // Find elements by class name
                    var elements = Global.appSession.FindElementsByClassName(className);

                    // Check if any elements are found
                    if (elements != null && elements.Count > 0)
                    {
                        // Return the list of elements found
                        return elements;
                    }
                }
                catch (NoSuchElementException)
                {
                    // Element not found, continue trying
                    attempts++;
                }
                catch (Exception ex)
                {
                    // Log any other exceptions and retry
                    Console.WriteLine($"Exception occurred while finding elements by class name: {ex.Message}");
                    attempts++;
                }
            }

            // Elements not found after max attempts
            Console.WriteLine($"No elements with class name '{className}' found after {maxAttempts} attempts.");
            return null;
        }

        public WindowsElement FindElementByName(string name)
        {
            const int maxAttempts = 10;
            int attempts = 0;
            WindowsElement element = null;

            while (attempts < maxAttempts)
            {
                try
                {
                    element = Global.appSession.FindElementByName(name);
                    if (element != null)
                    {
                        // Element found, return it
                        return element;
                    }
                }
                catch (NoSuchElementException)
                {
                    // Element not found, continue trying
                    attempts++;
                }
                catch (Exception ex)
                {
                    // Log any other exceptions and retry
                    Console.WriteLine($"Exception occurred while finding element by name: {ex.Message}");
                    attempts++;
                }
            }

            // Element not found after max attempts
            Console.WriteLine($"Element with name '{name}' not found after {maxAttempts} attempts.");
            return null;
        }

        public WindowsElement FindElementByClassName(string className)
        {
            const int maxAttempts = 10;
            int attempts = 0;
            WindowsElement element = null;

            while (attempts < maxAttempts)
            {
                try
                {
                    element = Global.appSession.FindElementByClassName(className);
                    if (element != null)
                    {
                        // Element found, return it
                        return element;
                    }
                }
                catch (NoSuchElementException)
                {
                    // Element not found, continue trying
                    attempts++;
                }
                catch (Exception ex)
                {
                    // Log any other exceptions and retry
                    Console.WriteLine($"Exception occurred while finding element by class name: {ex.Message}");
                    attempts++;
                }
            }

            // Element not found after max attempts
            Console.WriteLine($"Element with class name '{className}' not found after {maxAttempts} attempts.");
            return null;
        }

        public WindowsElement FindElementByAutomationId(string automationId)
        {
            const int maxAttempts = 10;
            int attempts = 0;
            WindowsElement element = null;

            while (attempts < maxAttempts)
            {
                try
                {
                    element = Global.appSession.FindElementByWindowsUIAutomation(automationId);
                    if (element != null)
                    {
                        // Element found, return it
                        return element;
                    }
                }
                catch (NoSuchElementException)
                {
                    // Element not found, continue trying
                    attempts++;
                }
                catch (Exception ex)
                {
                    // Log any other exceptions and retry
                    Console.WriteLine($"Exception occurred while finding element by automationId: {ex.Message}");
                    attempts++;
                }
            }

            // Element not found after max attempts
            Console.WriteLine($"Element with automationID'{automationId}' not found after {maxAttempts} attempts.");
            return null;
        }

        public WindowsElement FindElementByLegacyIAccessiblePatternName(string name)
        {
            const int maxAttempts = 10;
            int attempts = 0;
            WindowsElement element = null;

            while (attempts < maxAttempts)
            {
                try
                {
                    element = Global.appSession.FindElementByAccessibilityId(name);
                    if (element != null)
                    {
                        // Element found, return it
                        return element;
                    }
                }
                catch (NoSuchElementException)
                {
                    // Element not found, continue trying
                    attempts++;
                }
                catch (Exception ex)
                {
                    // Log any other exceptions and retry
                    Console.WriteLine($"Exception occurred while finding element by LegacyIAccessiblePattern.Name: {ex.Message}");
                    attempts++;
                }
            }

            // Element not found after max attempts
            Console.WriteLine($"Element with LegacyIAccessiblePattern.Name '{name}' not found after {maxAttempts} attempts.");
            return null;
        }

        public WindowsElement FindElementByXPath(string xPath)
        {
            const int maxAttempts = 10;
            int attempts = 0;
            WindowsElement element = null;

            while (attempts < maxAttempts)
            {
                try
                {
                    element = Global.appSession.FindElementByXPath(xPath);
                    if (element != null)
                    {
                        // Element found, return it
                        return element;
                    }
                }
                catch (NoSuchElementException)
                {
                    // Element not found, continue trying
                    attempts++;
                }
                catch (Exception ex)
                {
                    // Log any other exceptions and retry
                    Console.WriteLine($"Exception occurred while finding element by XPath: {ex.Message}");
                    attempts++;
                }
            }

            // Element not found after max attempts
            Console.WriteLine($"Element with XPath '{xPath}' not found after {maxAttempts} attempts.");
            return null;
        }

        public WindowsElement FindElementByXPathPartialName(string partialName)
        {
            const int maxAttempts = 10;
            int attempts = 0;
            WindowsElement element = null;

            while (attempts < maxAttempts)
            {
                try
                {
                    element = Global.appSession.FindElement(By.XPath($"//*[contains(@Name, '{partialName}')]"));
                    if (element != null)
                    {
                        // Element found, return it
                        return element;
                    }
                }
                catch (NoSuchElementException)
                {
                    // Element not found, continue trying
                    attempts++;
                }
                catch (Exception ex)
                {
                    // Log any other exceptions and retry
                    Console.WriteLine($"Exception occurred while finding element by XPath : {ex.Message}");
                    attempts++;
                }
            }

            // Element not found after max attempts
            Console.WriteLine($"Element with partial name '{partialName}' not found after {maxAttempts} attempts.");
            return null;
        }

        public bool VerifyCheckBoxIsOn(WindowsElement checkbox)
        {
            return checkbox.Selected;
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
