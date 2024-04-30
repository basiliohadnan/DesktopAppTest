using Consinco.Helpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OpenQA.Selenium.Appium.Windows;

namespace Consinco.MaxCompra
{
    [TestClass]
    public class MaxCompraInit : WinAppDriver
    {
        const string app = "MaxCompra";
        protected string excelFilePath = $"C:\\Users\\{Global.logonUser}\\source\\repos\\DesktopAppTest\\Dataset\\GerenciadordeCompras.xlsx";
        protected string matricula;
        protected ElementHandler elementHandler;
        protected ExcelReader excelReader;
        protected string suiteName = app;
        protected string appPath;

        public MaxCompraInit()
        {
            elementHandler = new ElementHandler();
            Global.app = app;
            excelReader = new ExcelReader();
            appPath = @$"C:\C5Client\Max\{app}.exe";
        }

        protected void Initialize()
        {
            StartWinAppDriver();
            InitializeWinSession();
            InitializeAppSession(appPath);
        }

        protected void Authenticate(string matricula)
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
        public void Login()
        {
            // Global Variables
            int rowNumber = 2;
            string worksheetName = "MaxComprasInit";
            ExcelWorksheet worksheet = excelReader.OpenWorksheet(excelFilePath, worksheetName);

            // Test Variables
            int lgsID;
            string printFileName;
            int reportID = int.Parse(excelReader.ReadCellValue(worksheet, "reportID", rowNumber));
            string scenarioName = excelReader.ReadCellValue(worksheet, "scenarioName", rowNumber);
            string testName = excelReader.ReadCellValue(worksheet, "testName", rowNumber);
            string testType = excelReader.ReadCellValue(worksheet, "testType", rowNumber);
            string analystName = excelReader.ReadCellValue(worksheet, "analystName", rowNumber);
            string testDesc = excelReader.ReadCellValue(worksheet, "testDesc", rowNumber);
           
            Global.processTest.StartTest(Global.customerName, suiteName, scenarioName, testName, testType, analystName, testDesc, reportID);

            // Test Details
            string preCondition = excelReader.ReadCellValue(worksheet, "preCondition", rowNumber);
            string postCondition = excelReader.ReadCellValue(worksheet, "postCondition", rowNumber);
            string inputData = excelReader.ReadCellValue(worksheet, "inputData", rowNumber);
            Global.processTest.DoTest(preCondition, postCondition, inputData);

            // Steps Definition
            Global.processTest.DoStep("Abrir app", "Abertura do app com sucesso");
            Global.processTest.DoStep("Login do analista", "Login com sucesso");
            Global.processTest.DoStep("Tela final", "Tela principal exibida com sucesso");

            // Steps Execution
            lgsID = Global.processTest.StartStep("Abrir app", logMsg: "Tentando abrir app", paramName: "appPath", paramValue: appPath);
            Initialize();
            printFileName = Global.processTest.CaptureWholeScreen();
            string welcomeWindowName = excelReader.ReadCellValue(worksheet, "welcomeWindowName", rowNumber);
            WindowsElement welcomeWindow = elementHandler.FindElementByName(welcomeWindowName);
            if (welcomeWindow != null)
            {
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: "app aberto");
            }
            else
            {
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "erro na abertura do app");
            }

            string matricula = excelReader.ReadCellValue(worksheet, "matricula", rowNumber);
            lgsID = Global.processTest.StartStep("Login do analista", logMsg: "Tentando login", paramName: "matricula", paramValue: matricula);
            Authenticate(matricula);
            SetAppSession();
            printFileName = Global.processTest.CaptureWholeScreen();
            string databaseWarningName = excelReader.ReadCellValue(worksheet, "databaseWarningName", rowNumber);
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
            string mainWindowClassName = excelReader.ReadCellValue(worksheet, "mainWindowClassName", rowNumber);
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