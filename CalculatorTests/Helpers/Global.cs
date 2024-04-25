using OpenQA.Selenium.Appium.Windows;
using Starline;
using Microsoft.Extensions.Configuration;

namespace Consinco.Helpers
{
    public class Global
    {
        public static string logonUser = Environment.UserName;
        public static WindowsDriver<WindowsElement> winSession;
        public static WindowsElement mainElement;
        public static WindowsDriver<WindowsElement> appSession;
        public static string app;
        public static string screenshotsDirectory;
        public static IConfiguration Configuration { get; set; }
        public static ProcessTest processTest;
        public static string customerName = "Assai";

        public static void InitializeProcessTest(IConfiguration configuration)
        {
            processTest = new ProcessTest(configuration);
        }
    }
}
