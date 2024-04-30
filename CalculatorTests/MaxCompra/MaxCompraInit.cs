using Consinco.Helpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium.Windows;

namespace Consinco.MaxCompra
{
    [TestClass]
    public class MaxCompraInit : WinAppDriver
    {
        const string app = "MaxCompra";
        private string filePath = $"C:\\Users\\{Global.user}\\source\\repos\\DesktopAppTest\\Dataset\\GerenciadordeCompras.xlsx";
        protected string matricula;
        protected ElementHandler elementHandler;
        protected ExcelReader excelReader;
        protected string suiteName = app;
        protected string appPath;

        public MaxCompraInit()
        {
            elementHandler = new ElementHandler();
            Global.app = app;
            excelReader = new ExcelReader(filePath);
            appPath = @$"C:\C5Client\Max\{app}.exe";
        }

        //[TestInitialize]
        protected void Initialize()
        {
            StartWinAppDriver();
            InitializeWinSession();
            InitializeAppSession(appPath);
        }

        protected void Login(string matricula)
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
            // Global Variables
            int rowNumber = 2;
            string matricula = excelReader.ReadCellValue("Matricula", rowNumber);
            string scenarioName = excelReader.ReadCellValue("scenarioName", rowNumber);
            int reportID = int.Parse(excelReader.ReadCellValue("reportID", rowNumber));
            int lgsID;
            string printFileName;

            // Test Variables
            string welcomeWindowName = excelReader.ReadCellValue("welcomeWindowName", rowNumber);
            string testName = excelReader.ReadCellValue("testName", rowNumber);
            string testType = excelReader.ReadCellValue("testType", rowNumber);
            string analystName = excelReader.ReadCellValue("analystName", rowNumber);
            string testDesc = excelReader.ReadCellValue("testDesc", rowNumber);
           
            Global.processTest.StartTest(Global.customerName, suiteName, scenarioName, testName, testType, analystName, testDesc, reportID);

            // Test Details
            string preCondition = excelReader.ReadCellValue("preCondition", rowNumber);
            string postCondition = excelReader.ReadCellValue("postCondition", rowNumber);
            string inputData = excelReader.ReadCellValue("inputData", rowNumber);
            Global.processTest.DoTest(preCondition, postCondition, inputData);

            // Steps Definition
            Global.processTest.DoStep("Abrir app", "Abertura do app com sucesso");
            Global.processTest.DoStep("Login do analista", "Login com sucesso");
            Global.processTest.DoStep("Tela final", "Tela principal exibida com sucesso");

            // Steps Execution
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
            Login(matricula);
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
            // Teardown function");
            Global.processTest.EndTest(reportID);
        }
    }
}