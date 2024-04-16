using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Interactions;
using System.Diagnostics;
using WindowsInput;
using static Consinco.Helpers.ElementHandler;

namespace Consinco.Helpers
{
    public class WinAppDriver
    {
        protected const string LogonUser = "sv_pocqa3";
        protected const string ScreenshotsDirectory = @"C:\Users\" + LogonUser + @"\source\repos\DesktopAppTest\Screenshots\";
        protected static WindowsDriver<WindowsElement> appSession;
        private IWebDriver driver;

        public WinAppDriver(IWebDriver driver)
        {
            this.driver = driver ?? throw new ArgumentNullException(nameof(driver));
        }

        public void SwitchToWindowWithTitle(string windowTitle)
        {
            foreach (var handle in driver.WindowHandles)
            {
                driver.SwitchTo().Window(handle);
                if (driver.Title.Contains(windowTitle))
                {
                    return;
                }
            }
            throw new NoSuchWindowException($"Window with title '{windowTitle}' not found.");
        }

        protected static void StartWinAppDriver()
        {
            // Start WinAppDriver process if not already running
            if (Process.GetProcessesByName("WinAppDriver").Length == 0)
            {
                Process.Start(@"C:\Program Files (x86)\Windows Application Driver\WinAppDriver.exe");
                WaitSeconds(5);
            }
        }

        protected static void StopWinAppDriver()
        {
            // Stop WinAppDriver process if running
            foreach (Process process in Process.GetProcessesByName("WinAppDriver"))
            {
                process.Kill();
            }
        }

        protected static void CloseApp(string appName)
        {
            // Close application if running
            foreach (Process process in Process.GetProcessesByName(appName))
            {
                process.Kill();
            }
        }

        [TestCleanup]
        public static void ClassCleanup()
        {
            CloseApp(Global.app);
            appSession?.Quit();
            StopWinAppDriver();
        }

        protected static WindowsDriver<WindowsElement> InitializeAppSession(string appPath)
        {
            // Start the app's process
            Process process = Process.Start(appPath);
            WaitSeconds(2);

            // Get the window handle of the app's process
            nint mainWindowHandle = process.MainWindowHandle;

            // Identify the root level window of the app's process
            WindowsDriver<WindowsElement> appSession;
            AppiumOptions rootCapabilities = new AppiumOptions();

            // Use the window handle as the appTopLevelWindow capability
            rootCapabilities.AddAdditionalCapability("appTopLevelWindow", mainWindowHandle.ToInt64().ToString("x"));
            appSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), rootCapabilities);

            return appSession;
        }

        protected static void PressEnter()
        {
            InputSimulator inputSimulator = new InputSimulator();
            inputSimulator.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.RETURN);
        }

        protected static void WaitSeconds(int seconds)
        {
            Thread.Sleep(seconds * 1000);
        }

        protected void ClickOn(WindowsDriver<WindowsElement> driver, BoundingRectangle boundingRectangle)
        {
            // Extract coordinates from the bounding rectangle
            int centerX = (boundingRectangle.Left + boundingRectangle.Right) / 2;
            int centerY = (boundingRectangle.Top + boundingRectangle.Bottom) / 2;

            // Click on the center of the item
            new Actions(driver).MoveByOffset(centerX, centerY).Click().Perform();
        }

        protected void ClickOn(WindowsDriver<WindowsElement> appSession, WindowsElement element)

        {
            new Actions(appSession).MoveToElement(element).Click().Perform();
        }

        protected static void FillField(WindowsDriver<WindowsElement> appSession, string information)
        {
            appSession.Keyboard.SendKeys(information);
        }
    }
}
