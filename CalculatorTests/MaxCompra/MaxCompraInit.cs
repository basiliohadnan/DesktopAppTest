﻿using Consinco.Helpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OpenQA.Selenium.Appium.Windows;
using Starline;

namespace Consinco.MaxCompra
{
    [TestClass]
    public class MaxCompraInit : WinAppDriver
    {
        protected const string app = "MaxCompra";
        //protected string appPath = @$"C:\Users\{Global.logonUser}\Desktop\SM_MAXCOMP_014\{app}.exe"; // v24.00.014 | release candidate
        //protected string appPath = @$" C:\C5Client\Max\{app}.exe"; // v23.00.036 | funcional no testes Login e Loja a loja
        protected string appPath = @$" C:\Users\sv_pocqa3\Desktop\MAXST_COMPRA_012\{app}.exe"; // v24.00.012 | Versão prod
        protected string excelFilePath = $"C:\\Users\\{Global.logonUser}\\source\\repos\\DesktopAppTest\\Dataset\\GerenciadordeCompras.xlsx";
        protected string matricula;
        protected ElementHandler elementHandler;
        protected ExcelReader excelReader;
        protected string suiteName = app;

        public MaxCompraInit()
        {
            elementHandler = new ElementHandler();
            Global.app = app;
            excelReader = new ExcelReader();
        }

        protected void Initialize()
        {
            StartWinAppDriver();
            InitializeWinSession();
            InitializeAppSession(appPath);
        }

        protected void Authenticate(string matricula, string loja = "000 - MATRIZ")
        {
            FillField(matricula);
            if (loja != "000 - MATRIZ")
            {
                WindowsElement lojasButton = elementHandler.FindElementByName("Open");
                lojasButton.Click();
                FillField(loja);
            }
            PressEnter();
        }

        protected void SetAppSession()
        {
            string className = "Centura:MDIFrame";
            SetAppSession(className);
        }

        protected void ConfirmDatabaseWarning(WindowsElement warning)
        {
            warning.Click();
            PressEnter();
        }

        protected void OpenMenu(string menuName)
        {
            WindowsElement menuItem = elementHandler.FindElementByName(menuName);
            menuItem.Click();
        }

        protected void Login(ExcelWorksheet worksheet, int rowNumber)
        {
            int lgsID;
            string printFileName;

            lgsID = Global.processTest.StartStep("Abrir app", logMsg: "Tentando abrir app", paramName: "appPath", paramValue: appPath);
            Initialize();
            printFileName = Global.processTest.CaptureWholeScreen();
            string welcomeWindowName = excelReader.ReadCellValueToString(worksheet, "welcomeWindowName", rowNumber);
            WindowsElement welcomeWindow = elementHandler.FindElementByName(welcomeWindowName);
            if (welcomeWindow != null)
            {
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: "app aberto");
            }
            else
            {
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "erro na abertura do app");
            }

            string loja = excelReader.ReadCellValueToString(worksheet, "loja", rowNumber);
            string matricula = excelReader.ReadCellValueToString(worksheet, "matricula", rowNumber);
            lgsID = Global.processTest.StartStep("Login do analista", logMsg: "Tentando login", paramName: "matricula", paramValue: matricula);
            if (loja != null)
            {
                Authenticate(matricula, loja);
            }
            else
            {

                Authenticate(matricula);
            }
            SetAppSession();
            printFileName = Global.processTest.CaptureWholeScreen();
            string databaseWarningName = excelReader.ReadCellValueToString(worksheet, "databaseWarningName", rowNumber);
            WindowsElement warning = elementHandler.FindElementByXPathPartialName(databaseWarningName);
            if (warning != null)
            {
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: "login efetuado");
            }
            else
            {
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "erro no login");
            }

            lgsID = Global.processTest.StartStep("Tela final", logMsg: "Tentando acessar tela principal", paramName: "", paramValue: "");
            ConfirmDatabaseWarning(warning);
            string mainWindowClassName = excelReader.ReadCellValueToString(worksheet, "mainWindowClassName", rowNumber);
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
        }

        protected void LoginWithInputData(InputData inputData)
        {
            int lgsID;
            string printFileName;

            lgsID = Global.processTest.StartStep("Abrir app", logMsg: "Tentando abrir app", paramName: "appPath", paramValue: appPath);
            Initialize();
            printFileName = Global.processTest.CaptureWholeScreen();
            string welcomeWindowName = inputData.GetValue("WELCOMEWINDOWNAME");
            WindowsElement welcomeWindow = elementHandler.FindElementByName(welcomeWindowName);
            if (welcomeWindow != null)
            {
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: "app aberto");
            }
            else
            {
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "erro na abertura do app");
            }

            string loja = inputData.GetValue("LOJA");
            string matricula = inputData.GetValue("MATRICULA");
            lgsID = Global.processTest.StartStep("Login do analista", logMsg: "Tentando login", paramName: "matricula", paramValue: matricula);
            if (loja != null)
            {
                Authenticate(matricula, loja);
            }
            else
            {
                Authenticate(matricula);
            }
            SetAppSession();
            printFileName = Global.processTest.CaptureWholeScreen();
            string databaseWarningName = inputData.GetValue("DATABASEWARNINGNAME");
            WindowsElement warning = elementHandler.FindElementByXPathPartialName(databaseWarningName);
            if (warning != null)
            {
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: "login efetuado");
            }
            else
            {
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "erro no login");
            }

            lgsID = Global.processTest.StartStep("Tela final", logMsg: "Tentando acessar tela principal", paramName: "", paramValue: "");
            ConfirmDatabaseWarning(warning);
            string mainWindowClassName = inputData.GetValue("MAINWINDOWCLASSNAME");
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
        }

        [TestMethod]
        public void RealizarLogin()
        {
            // Global Variables
            int rowNumber = 3;
            string worksheetName = "MaxComprasInit";
            ExcelWorksheet worksheet = excelReader.OpenWorksheet(excelFilePath, worksheetName);

            // Test Variables
            int lgsID;
            string printFileName;
            int reportID = int.Parse(excelReader.ReadCellValueToString(worksheet, "reportID", rowNumber));
            string scenarioName = excelReader.ReadCellValueToString(worksheet, "scenarioName", rowNumber);
            string testName = excelReader.ReadCellValueToString(worksheet, "testName", rowNumber);
            string testType = excelReader.ReadCellValueToString(worksheet, "testType", rowNumber);
            string analystName = excelReader.ReadCellValueToString(worksheet, "analystName", rowNumber);
            string testDesc = excelReader.ReadCellValueToString(worksheet, "testDesc", rowNumber);

            Global.processTest.StartTest(Global.customerName, suiteName, scenarioName, testName, testType, analystName, testDesc, reportID);

            // Test Details
            string preCondition = excelReader.ReadCellValueToString(worksheet, "preCondition", rowNumber);
            string postCondition = excelReader.ReadCellValueToString(worksheet, "postCondition", rowNumber);
            string inputData = excelReader.ReadCellValueToString(worksheet, "inputData", rowNumber);
            Global.processTest.DoTest(preCondition, postCondition, inputData);

            // Steps Definition
            Global.processTest.DoStep("Abrir app", "Abertura do app com sucesso");
            Global.processTest.DoStep("Login do analista", "Login com sucesso");
            Global.processTest.DoStep("Tela final", "Tela principal exibida com sucesso");

            Login(worksheet, rowNumber);

            // Teardown function");
            Global.processTest.EndTest(reportID);
        }

        [TestMethod]
        public void RealizarLoginComSelectExcel()
        {
            // Global Variables
            int reportID = 2;
            InputData inputExcel = new InputData(ConnType: "Excel", ConnXLS: excelFilePath);
            int rows = inputExcel.NewQuery(QueryText: $"SELECT * FROM [MaxComprasInit$] WHERE reportID = {reportID}");

            // Examples
            //inputExcel.RunDDL("insert into [Planilha1$] values (@v_txt_x)", false, "v_txt_x: z");
            //inputExcel.RunDDL("update [Planilha1$] set TESTE = 'oi' where TESTE = @v_txt_x", false, "v_txt_x: valor");

            // Test Variables
            int lgsID;
            string printFileName;
            string scenarioName = inputExcel.GetValue("SCENARIONAME");
            string testName = inputExcel.GetValue("TESTNAME");
            string testType = inputExcel.GetValue("TESTTYPE");
            string analystName = inputExcel.GetValue("ANALYSTNAME");
            string testDesc = inputExcel.GetValue("TESTDESC");

            Global.processTest.StartTest(Global.customerName, suiteName, scenarioName, testName, testType, analystName, testDesc, reportID);

            // Test Details
            string preCondition = inputExcel.GetValue("PRECONDITION");
            string postCondition = inputExcel.GetValue("POSTCONDITION");
            string inputData = inputExcel.GetValue("INPUTDATA");
            Global.processTest.DoTest(preCondition, postCondition, inputData);

            // Steps Definition
            Global.processTest.DoStep("Abrir app", "Abertura do app com sucesso");
            Global.processTest.DoStep("Login do analista", "Login com sucesso");
            Global.processTest.DoStep("Tela final", "Tela principal exibida com sucesso");

            LoginWithInputData(inputExcel);

            // Teardown function
            Global.processTest.EndTest(reportID);
        }
    }
}
