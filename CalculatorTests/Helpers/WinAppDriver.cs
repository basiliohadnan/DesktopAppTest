﻿using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Appium;
using System.Diagnostics;
using WindowsInput;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium;
using WindowsInput.Native;
using OpenQA.Selenium.Appium.MultiTouch;
using System.Collections.ObjectModel;

namespace Consinco.Helpers
{
    public class WinAppDriver
    {

        protected void StartWinAppDriver()
        {
            if (Process.GetProcessesByName("WinAppDriver").Length == 0)
            {
                Process.Start(@"C:\Program Files (x86)\Windows Application Driver\WinAppDriver.exe");
            }
        }

        protected void InitializeWinSession()
        {
            AppiumOptions winCapabilities = new AppiumOptions();
            winCapabilities.AddAdditionalCapability("app", "Root");
            Global.winSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), winCapabilities);
            ElementHandler handler = new ElementHandler();
            Global.mainElement = Global.winSession.FindElementByXPath("//*");
        }

        protected void StopWinAppDriver()
        {
            foreach (Process process in Process.GetProcessesByName("WinAppDriver"))
            {
                process.Kill();
            }
        }

        protected void CloseApp(string app)
        {
            {

                WindowsElement exitButton = new ElementHandler().FindElementByName("Close", 3000);
                exitButton.Click();
            }
        }

        [TestCleanup]
        public void Cleanup()
        {
            CloseApp(Global.app);
            Global.appSession?.Quit();
            StopWinAppDriver();
        }

        protected void InitializeAppSession(string appPath)
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

        public void SetWindowSize(int width, int height)
        {
            IWebElement element = Global.appSession.FindElement(By.ClassName("Centura:AccFrame"));
            IJavaScriptExecutor jsExecutor = Global.appSession;
            jsExecutor.ExecuteScript($"arguments[0].style.width='{width}px'; arguments[0].style.height='{height}px';", element);

            //Global.appSession.Manage().Window.Size = new Size(width, height);
        }

        protected void SetAppSession(string className)
        {
            WaitSeconds(1);
            var appWindow = Global.winSession.FindElementByClassName(className);
            AppiumOptions appCapabilities = new AppiumOptions();
            var rootTopLevelWindowHandle = appWindow.GetAttribute("NativeWindowHandle");
            rootTopLevelWindowHandle = (int.Parse(rootTopLevelWindowHandle)).ToString("x"); // Convert to Hex
            appCapabilities.AddAdditionalCapability("appTopLevelWindow", rootTopLevelWindowHandle);
            Global.appSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);
        }

        public static void PressEnter()
        {
            InputSimulator inputSimulator = new InputSimulator();
            inputSimulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);
        }

        public static void PressKey(string key)
        {
            VirtualKeyCode keyCode;
            if (Enum.TryParse(key, out keyCode))
            {
                InputSimulator inputSimulator = new InputSimulator();
                inputSimulator.Keyboard.KeyPress(keyCode);
            }
            else
            {
                Console.WriteLine("Invalid key definition.");
            }
        }

        public static void WaitSeconds(int seconds)
        {
            Thread.Sleep(seconds * 1000);
        }

        public static void FillField(string information)
        {
            Global.appSession.Keyboard.SendKeys(information);
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

        public static void ClickOn(int offsetX, int offsetY)
        {
            Global.winSession.Mouse.MouseMove(Global.mainElement.Coordinates, offsetX, offsetY);
            Global.winSession.Mouse.Click(null);
        }

        public static void ClickOn(WindowsElement element)

        {
            new Actions(Global.appSession).MoveToElement(element).Click().Perform();
        }

        public static void DoubleClickOn(WindowsElement element)

        {
            new Actions(Global.appSession).MoveToElement(element).DoubleClick().Perform();
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

        public static void DoubleClickOn(int offsetX, int offsetY)
        {
            Global.winSession.Mouse.MouseMove(Global.mainElement.Coordinates, offsetX, offsetY);
            Global.winSession.Mouse.DoubleClick(null);
        }

        public static void DoubleTapOn(AppiumWebElement element)
        {
            TouchAction action = new TouchAction(Global.appSession);
            action.Tap(element).Perform();
            action.Tap(element).Perform();
        }

        public static void MaximizeWindow()
        {
            AppiumWebElement button = new ElementHandler().FindElementByName("Maximize");
            button.Click();
        }

        public static void RestoreWindow()
        {
            ReadOnlyCollection<WindowsElement> buttons = new ElementHandler().FindElementsByName("Restore");
            buttons[1].Click();
        }
    }
}
