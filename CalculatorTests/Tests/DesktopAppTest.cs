using System;
using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using WindowsInput;
using System.Threading;
using DesktopAppTests.Helpers;
using OpenQA.Selenium;

namespace DesktopAppTests.Tests
{
    public class DesktopAppTest
    {
        protected const string ScreenshotsDirectory = @"C:\Users\Starline\source\repos\DesktopAppTest\CalculatorTests\Screenshots\";
        protected static WindowsDriver<WindowsElement> appSession;

        protected static void StartWinAppDriver()
        {
            // Start WinAppDriver process if not already running
            if (Process.GetProcessesByName("WinAppDriver").Length == 0)
            {
                Process.Start(@"C:\Program Files (x86)\Windows Application Driver\WinAppDriver.exe");
                // Wait for WinAppDriver to start
                Thread.Sleep(5000);
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

        protected static WindowsDriver<WindowsElement> InitializeAppSession(string appPath, string appClassName)
        {
            // Start the app's process
            Process process = Process.Start(appPath);

            // Wait for a short duration to allow the process to initialize
            Thread.Sleep(2000);

            // Simulate pressing the Enter key
            InputSimulator inputSimulator = new InputSimulator();
            inputSimulator.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.RETURN);

            // Wait for another short duration after pressing Enter
            Thread.Sleep(2000);

            // Get the window handle of the app's process
            IntPtr mainWindowHandle = process.MainWindowHandle;

            // Identify the root level window of the app's process
            WindowsDriver<WindowsElement> winSession;
            AppiumOptions rootCapabilities = new AppiumOptions();
            // Use the window handle as the appTopLevelWindow capability
            rootCapabilities.AddAdditionalCapability("appTopLevelWindow", mainWindowHandle.ToInt64().ToString("x"));
            winSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), rootCapabilities);

            return winSession;
        }

        protected void WriteTest(string text)
        {
            // Wait for the app to load
            Thread.Sleep(500);

            // Clear the content of the app by selecting all text and then deleting it
            appSession.FindElementByClassName("Edit").SendKeys(Keys.Control + "a");
            appSession.FindElementByClassName("Edit").SendKeys(Keys.Delete);

            // Enter values in app
            appSession.FindElementByClassName("Edit").SendKeys(text);
            // Press Enter
            appSession.FindElementByClassName("Edit").SendKeys(Keys.Enter);
        }
    }
}