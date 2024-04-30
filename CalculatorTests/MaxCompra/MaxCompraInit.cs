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
        protected ElementHandler elementHandler;
        protected string suiteName = app;
        protected string appPath;

        public MaxCompraInit()
        {
            elementHandler = new ElementHandler();
            Global.app = app;
            appPath = @$"C:\C5Client\Max\{app}.exe";
        }

        //[TestInitialize]
        protected void Initialize()
        {
            StartWinAppDriver();
            InitializeWinSession();
            InitializeAppSession(appPath);
        }

        protected void Login()
        {
            FillField(matricula);
            PressEnter();
        }

        protected void SetAppSession()
        {
            string className = "Centura:MDIFrame";
            SetAppSession(className);
        }

        protected void OpenMenu(string menuName)
        {
            WindowsElement menuItem = elementHandler.FindElementByName(menuName);
            menuItem.Click();
        }

        [TestMethod]
        public void LoginWithReport()
        {
            //Console.WriteLine("Global Variables");
            string scenarioName = "Relatorio - Realizar Login";
            int reportID = 1;
            int lgsID;
            string printFileName;

            //Console.WriteLine("Test Variables");
            string welcomeWindowName = "Conexão de Sistemas Consinco";
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

            //Console.WriteLine("Steps Definition");
            Global.processTest.DoStep("Abrir app", "Abertura do app com sucesso");
            Global.processTest.DoStep("Login do analista", "Login com sucesso");
            Global.processTest.DoStep("Tela final", "Tela principal exibida com sucesso");
            Global.processTest.DoStep("Abrir menu Administração", "Menu Administracao aberto com sucesso");
            Global.processTest.DoStep("Abrir menu Compras", "Menu Compras aberto com sucesso");
            Global.processTest.DoStep("Abrir menu Gerenciador de Compras", "Menu Gerenciador de Compras aberto com sucesso");

            //Console.WriteLine("Steps Execution");
            lgsID = Global.processTest.StartStep("Abrir app", logMsg: "Tentando abrir app", paramName: "appPath", paramValue: appPath);
            Initialize();
            printFileName = Global.processTest.CaptureWholeScreen();

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
            printFileName = Global.processTest.CaptureWholeScreen();

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

            string mainWindowClassName = "Centura:MDIFrame";
            WindowsElement mainWindow = elementHandler.FindElementByClassName(mainWindowClassName);
            printFileName = Global.processTest.CaptureWholeScreen();
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
        public void ReadsExcel()
        {
            string filePath = $"C:\\Users\\{Global.user}\\source\\repos\\DesktopAppTest\\Dataset\\GerenciadordeCompras.xlsx";
            ExcelReader excelReader = new ExcelReader(filePath);
            string columnName = "Report ID";
            int rowNumber = 2; // Assuming the data starts from the second row
            string cellValue = excelReader.ReadCellValue(columnName, rowNumber);
            Console.WriteLine($"Value in column '{columnName}' at row {rowNumber}: {cellValue}");
        }
    }
}