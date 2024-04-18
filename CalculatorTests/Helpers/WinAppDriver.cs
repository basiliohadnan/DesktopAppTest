﻿using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Appium;
using System.Diagnostics;
using WindowsInput;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium;
using DesktopAppTests.Helpers;
using WindowsInput.Native;
using OpenQA.Selenium.Support.UI;

namespace Consinco.Helpers
{
    public class WinAppDriver
    {
        protected const string logonUser = "sv_pocqa3";

        public WinAppDriver()
        {
            Global.screenshotsDirectory = @"C:\Users\" + logonUser + @"\source\repos\DesktopAppTest\CalculatorTests\Screenshots\";
        }

        protected static void StartWinAppDriver()
        {
            if (Process.GetProcessesByName("WinAppDriver").Length == 0)
            {
                Process.Start(@"C:\Program Files (x86)\Windows Application Driver\WinAppDriver.exe");
            }
        }

        protected static void InitializeWinSession()
        {
            AppiumOptions winCapabilities = new AppiumOptions();
            winCapabilities.AddAdditionalCapability("app", "Root");
            Global.winSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), winCapabilities);
            Global.mainElement = Global.winSession.FindElementByXPath("//*");
        }

        protected static void StopWinAppDriver()
        {
            foreach (Process process in Process.GetProcessesByName("WinAppDriver"))
            {
                process.Kill();
            }
        }

        protected static void CloseApp(string app)
        {
            foreach (Process process in Process.GetProcessesByName(app))
            {
                process.Kill();
            }
        }

        [TestCleanup]
        public void Cleanup()
        {
            CloseApp(Global.app);
            Global.appSession?.Quit();
            StopWinAppDriver();
        }

        protected static void InitializeAppSession(string appPath)
        {
            Process process = Process.Start(appPath);
            WaitSeconds(2);

            // Get the window handle of the app's process
            nint mainWindowHandle = process.MainWindowHandle;

            // Identify the root level window of the app's process
            AppiumOptions rootCapabilities = new AppiumOptions();

            // Use the window handle as the appTopLevelWindow capability
            rootCapabilities.AddAdditionalCapability("appTopLevelWindow", mainWindowHandle.ToInt64().ToString("x"));
            Global.appSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), rootCapabilities);
        }

        protected static void SetAppSession(string description)
        {
            WaitSeconds(1);
            var appWindow = Global.winSession.FindElementByClassName(description);
            AppiumOptions appCapabilities = new AppiumOptions();
            var rootTopLevelWindowHandle = appWindow.GetAttribute("NativeWindowHandle");
            rootTopLevelWindowHandle = (int.Parse(rootTopLevelWindowHandle)).ToString("x"); // Convert to Hex
            appCapabilities.AddAdditionalCapability("appTopLevelWindow", rootTopLevelWindowHandle);
            Global.appSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);
        }

        public static void PressEnter()
        {
            InputSimulator inputSimulator = new InputSimulator();
            inputSimulator.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.RETURN);
        }

        public static void WaitSeconds(int seconds)
        {
            Thread.Sleep(seconds * 1000);
        }

        public static void FillField(string information)
        {
            Global.appSession.Keyboard.SendKeys(information);
        }

        public static void SendKeys(KeyboardKey key)
        {
            InputSimulator inputSimulator = new InputSimulator();
            inputSimulator.Keyboard.KeyPress((VirtualKeyCode)key);
        }
        public static void SelectContentFromField()
        {
            Global.appSession.Keyboard.SendKeys(Keys.Control + "a");
        }

        public static void ClearField()
        {
            SelectContentFromField();
            Global.appSession.Keyboard.SendKeys(Keys.Delete);
        }

        public static void ClickOn(ElementHandler.BoundingRectangle boundingRectangle)
        {
            // Extract coordinates from the bounding rectangle
            int offsetX = (boundingRectangle.Left + boundingRectangle.Right) / 2;
            int offsetY = (boundingRectangle.Top + boundingRectangle.Bottom) / 2;

            Global.winSession.Mouse.MouseMove(Global.mainElement.Coordinates, offsetX, offsetY);
            Global.winSession.Mouse.Click(null);
        }

        public static void ClickOn(WindowsElement element)

        {
            new Actions(Global.appSession).MoveToElement(element).Click().Perform();
        }
        public static void Click()

        {
            Global.appSession.Mouse.Click(null);
        }

        public static void DoubleClick()

        {
            Global.appSession.Mouse.DoubleClick(null);
        }

        public static void DoubleClickOn(ElementHandler.BoundingRectangle boundingRectangle)
        {
            int offsetX = (boundingRectangle.Left + boundingRectangle.Right) / 2;
            int offsetY = (boundingRectangle.Top + boundingRectangle.Bottom) / 2;

            Global.winSession.Mouse.MouseMove(Global.mainElement.Coordinates, offsetX, offsetY);
            Global.winSession.Mouse.DoubleClick(null);
        }
                
        public static void WaitForElementVisibleByClassName(string className, int seconds)
        {
            var timeout = TimeSpan.FromSeconds(seconds);
            WebDriverWait wait = new WebDriverWait(Global.winSession, timeout);
            wait.Until(ExpectedConditions.ElementExists(By.ClassName(className)));
        }

        public static void WaitForElementVisibleByName(string name, int seconds)
        {
            var timeout = TimeSpan.FromSeconds(seconds);
            WebDriverWait wait = new WebDriverWait(Global.winSession, timeout);
            wait.Until(ExpectedConditions.ElementExists(By.Name(name)));
        }
    }
}
