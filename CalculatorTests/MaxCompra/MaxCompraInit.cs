using Consinco.Helpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium.Windows;

namespace Consinco.MaxCompra
{
    [TestClass]
    public class MaxCompraInit : WinAppDriver
    {
        const string app = "MaxCompra";
        protected string matricula = "PS032528";
        private ElementHandler elementHandler;
        private string suiteName = app;
        protected string appPath;

        public MaxCompraInit()
        {
            elementHandler = new ElementHandler();
            Global.app = app;
            appPath = @$"C:\C5Client\Max\{app}.exe";
        }

        //[TestInitialize]
        private void Initialize()
        {
            StartWinAppDriver();
            InitializeWinSession();
            InitializeAppSession(appPath);
        }

        private void Login()
        {
            FillField(matricula);
            PressEnter();
        }

        private void SetAppSession()
        {
            string className = "Centura:MDIFrame";
            SetAppSession(className);
        }

        public void OpenMenu(string menuItemName, string testName)
        {
            var menuItem = elementHandler.FindElementByName(menuItemName);
            menuItem.Click();
            ScreenPrinter.CaptureAndSaveScreenshot(Global.screenshotsDirectory, testName);
        }

        [TestMethod]
        public void LoginWithReport()
        {
            Console.WriteLine("Global Variables");
            string scenarioName = "Relatorio - Realizar Login";
            int reportID = 1;
            int lgsID;
            string printFileName;

            //Console.WriteLine("Test Variables");
            string welcomeWindowName = "Conexão de Sistemas Consinco";
            string mainWindowName = "acrux mercari - Compras";
            string testName = "Login";
            string testType = "Funcional";
            string analystName = "Hadnan Basilio";
            string testDesc = "Realizar login no app MaxCompra com matricula de analista";
            Global.processTest.StartTest(Global.customerName, suiteName, scenarioName, testName, testType, analystName, testDesc, reportID);

            //Console.WriteLine("Test Details");
            string preCondition = "App instalado";
            string postCondition = "Nenhuma";
            string inputData = "Matricula valida";
            Global.processTest.DoTest(preCondition, postCondition, inputData);

            //Console.WriteLine("Steps Definition");
            Global.processTest.DoStep("Abrir app", "Abertura do app com sucesso");
            Global.processTest.DoStep("Login do analista", "Login com sucesso");
            Global.processTest.DoStep("Tela final", "Tela principal exibida com sucesso");

            //Console.WriteLine("Steps Execution");
            lgsID = Global.processTest.StartStep("Abrir app", logMsg: "Tentando abrir app", paramName: "appPath", paramValue: appPath);
            Initialize();
            printFileName = Global.processTest.PrintScreen();

            WindowsElement welcomeWindow = elementHandler.FindElementByName(welcomeWindowName);
            if (welcomeWindow != null)
            {
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: "app aberto");
            }
            else
            {
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "erro na abertura do app");
            }

            lgsID = Global.processTest.StartStep("Login do analista", logMsg: "Tentando login", paramName: "matricula", paramValue: matricula);
            Login();
            SetAppSession();
            printFileName = Global.processTest.PrintScreen();

            string databaseWarningName = "Não foi definido a versão do módulo no BANCO DE DADOS, para o sistema de Segurança";
            WindowsElement databaseWarning = elementHandler.FindElementByXPathPartialName(databaseWarningName);
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
            //Console.WriteLine("Teardown function");
            Global.processTest.EndTest(reportID);
        }

        [TestMethod]
        public void OpenMenuGerenciadorDeComprasWithReport()
        {
            LoginWithReport();

            Console.WriteLine("Global Variables");
            string scenarioName = "Abrir menu - Gerenciador de Compras";
            int reportID = 2;
            int lgsID;
            string printFileName;

            Console.WriteLine("Test Variables");
            List<string> menus = new List<string>
        {
            "Administração",
            "Compras",
            "Gerenciador de Compras"
        };
            string testName = "Menu Gerenciador de Compras";
            string testType = "Funcional";
            string analystName = "Hadnan Basilio";
            string testDesc = "Abrir menu: Gerenciador de Compras";

            Global.processTest.StartTest(Global.customerName, suiteName, scenarioName, testName, testType, analystName, testDesc, reportID);

            Console.WriteLine("Test Details");
            string preCondition = "App iniciando normalmente";
            string postCondition = "Painel gerenciador de compras aberto";
            string inputData = "";
            Global.processTest.DoTest(preCondition, postCondition, inputData);

            Console.WriteLine("Steps Definition");
            Global.processTest.DoStep("Abrir menu Administracao", "Abertura do menu Administracao: sucesso");
            Global.processTest.DoStep("Abrir submenu Compras", "Abertura do submenu Compras: sucesso");
            Global.processTest.DoStep("Abrir submenu Gerenciador de Compras", "Abertura do submenu Gerenciador de Compras: sucesso");

            menus.ForEach(menuName =>
            {
                Console.WriteLine("Steps Execution");
                lgsID = Global.processTest.StartStep($"Abrir menu {menuName}", logMsg: $"Tentando menu {menuName}", paramName: "menuName", paramValue: menuName);
                var menuItem = elementHandler.FindElementByName(menuName);

                printFileName = Global.processTest.PrintScreen();
                if (menuItem != null)
                {
                    Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"menu {menuName} aberto");
                    menuItem.Click();
                }
                else
                {
                    Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"erro na abertura do {menuName}");
                }

            });

            Console.WriteLine("Teardown function");
            Global.processTest.EndTest(reportID);

        }
    }
}
