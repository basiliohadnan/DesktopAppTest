using Consinco.Helpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Consinco.MaxCompra
{
    [TestClass]
    public class MaxCompraInit : WinAppDriver
    {
        const string user = "PS032528";
        const string app = "MaxCompra";
        const string appPath = @$"C:\C5Client\Max\{app}.exe";
        private ElementHandler elementHandler;
        private string suiteName = app;

        public MaxCompraInit()
        {
            elementHandler = new ElementHandler();
            Global.app = app;
        }

        [TestInitialize]
        public void Initialize()
        {
            StartWinAppDriver();
            InitializeWinSession();
            InitializeAppSession(appPath);
        }

        [TestMethod]
        public void Login()
        {
            FillField(user);
            PressEnter();
            SetAppSession("Centura:MDIFrame");

            //PressEnter();
            //ScreenPrinter.CaptureAndSaveScreenshot(Global.screenshotsDirectory, "01-Login");
        }

        public void OpenMenu(string menuItemName, string testName)
        {
            var menuItem = elementHandler.FindElementByName(menuItemName);
            menuItem.Click();
            ScreenPrinter.CaptureAndSaveScreenshot(Global.screenshotsDirectory, testName);
        }

        [TestMethod]
        public void LoginDummy()
        {
            Console.WriteLine("Global Variables");
            string scenarioName = "Relatorio - Realizar Login";
            int reportID = 1;
            int lgsID;
            string printFileName;
            // name

            Console.WriteLine("Test Variables");
            string testName = "Login";
            string testType = "Funcional";
            string analystName = "Hadnan Basilio";
            string testDesc = "Realizar login no app MaxCompra com matricula de analista";
            Global.processTest.StartTest(Global.customerName, suiteName, scenarioName, testName, testType, analystName, testDesc, reportID);

            Console.WriteLine("Test Details");
            string preCondition = "App instalado";
            string postCondition = "Nenhuma";
            string inputData = "Matricula valida";
            Global.processTest.DoTest(preCondition, postCondition, inputData);

            Console.WriteLine("Steps Definition");
            Global.processTest.DoStep("Abrir app", "Abertura do app com sucesso");
            Global.processTest.DoStep("Login do analista", "Login com sucesso");
            Global.processTest.DoStep("Tela final", "xx");

            Console.WriteLine("Steps Execution");
            lgsID = Global.processTest.StartStep("Abrir app", logMsg: "Tentando abrir app", paramName: "appPath", paramValue: appPath);
            // Step actions
            // Test Initialize runs
            // Opens app
            var windowName = "Conexão de Sistemas Consinco";
            var welcomeWindow = elementHandler.FindElementByName(windowName);

            printFileName = Global.processTest.PrintScreen();
            if (welcomeWindow != null)
            {
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: "app aberto");
            }
            else
            {
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "erro na abertura do app");
            }

            lgsID = Global.processTest.StartStep("Login do analista", logMsg: "Tentando login", paramName: "matricula", paramValue: user);
            Login();
            var warningText = "Não foi definido a versão do módulo no BANCO DE DADOS, para o sistema de Segurança";
            var databaseWarning = elementHandler.FindElementByXPathPartialName(warningText);

            printFileName = Global.processTest.PrintScreen();
            if (databaseWarning != null)
            {
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: "login efetuado");
            }
            else
            {
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "erro no login");
            }

            lgsID = Global.processTest.StartStep("Tela final", logMsg: "Tentando acessar tela principal", paramName: "", paramValue: "");
            databaseWarning.Click();
            
            PressEnter();

            var mainWindowName = "acrux mercari - Compras";
            var mainWindow = elementHandler.FindElementByXPathPartialName(mainWindowName);

            printFileName = Global.processTest.PrintScreen();
            if (mainWindow != null)
            {
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: "tela principal exibida");
            }
            else
            {
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "erro na exibição da tela principal");
            }

            Console.WriteLine("Teardown function");
            Global.processTest.EndTest(reportID);

            //FillField(user);
            //PressEnter();
            //PressEnter();

            //SetAppSession("Centura:MDIFrame");
            //ScreenPrinter.CaptureAndSaveScreenshot(Global.screenshotsDirectory, "01-Login");
        }


    }
}
