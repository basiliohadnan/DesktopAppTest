using Consinco.Helpers;
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

        private void DefineSteps(string testName)
        {
            switch (testName)
            {
                case "RealizarLogin":
                    Global.processTest.DoStep("Abrir app", "App aberto.");
                    Global.processTest.DoStep("Realizar login do analista", "Login efetuado.");
                    Global.processTest.DoStep("Validar tela principal exibida", "Tela principal exibida.");
                    break;
                default:
                    throw new Exception($"{testName} has no steps definition.");
            }
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

        protected void Login(InputData inputData, string queryName)
        {
            string stepDescription = "Abrir app";
            int lgsID;
            string printFileName;
            string paramName = "appPath";
            string paramValue = appPath;
            string expectedResult = "App aberto.";

            lgsID = Global.processTest.StartStep(stepDescription, logMsg:
                $"Tentando {stepDescription}", paramName: paramName, paramValue: paramValue);
            try
            {
                Initialize();
                printFileName = Global.processTest.CaptureWholeScreen();
                string welcomeWindowName = "Conexão de Sistemas Consinco";
                WindowsElement welcomeWindow = elementHandler.FindElementByName(welcomeWindowName);
                Assert.IsNotNull(welcomeWindow);
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: expectedResult);
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg:
                    $"Erro ao tentar {stepDescription}.");
                throw new Exception($"Erro ao tentar {stepDescription}.");
            };

            stepDescription = "Realizar login do analista";
            paramName = "matricula";
            string matricula = inputData.GetValue("MATRICULA", queryName);
            string loja = inputData.GetValue("LOJA", queryName);
            paramValue = matricula;
            expectedResult = "Login efetuado.";

            lgsID = Global.processTest.StartStep(stepDescription, logMsg: $"Tentando {stepDescription}", paramName: paramName, paramValue: paramValue);
            try
            {
                if (loja != null)
                {
                    Authenticate(matricula, loja);
                }
                else
                {
                    Authenticate(matricula, queryName);
                }
                SetAppSession();
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: expectedResult);
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg:
                    $"Erro ao tentar {stepDescription}.");
                throw new Exception($"Erro ao tentar {stepDescription}.");
            }

            stepDescription = "Validar tela principal exibida";
            paramName = "";
            string databaseWarningName = "Não foi definido a versão do módulo no BANCO DE DADOS, para o sistema de Segurança";
            WindowsElement warning = elementHandler.FindElementByXPathPartialName(databaseWarningName);
            paramValue = "";
            expectedResult = "Tela principal exibida";
            lgsID = Global.processTest.StartStep(stepDescription, logMsg: $"Tentando {stepDescription}", paramName: paramName, paramValue: paramValue);
            Assert.IsNotNull(warning);
            try
            {
                ConfirmDatabaseWarning(warning);
                string mainWindowClassName = "Centura:MDIFrame";
                WindowsElement mainWindow = elementHandler.FindElementByClassName(mainWindowClassName);
                Assert.IsNotNull(mainWindow);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: expectedResult);
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg:
                    $"Erro ao tentar {stepDescription}.");
                throw new Exception($"Erro ao tentar {stepDescription}.");
            }
        }

        [TestMethod]
        public void RealizarLogin()
        {
            // Global Variables
            int testId = 1;
            string queryName = "RealizarLoginComSelectExcel";
            InputData inputExcel = new InputData(ConnType: "Excel", ConnXLS: excelFilePath);
            inputExcel.NewQuery(
                QueryName: queryName,
                QueryText: $"SELECT * FROM [MaxComprasInit$] WHERE testId = {testId}"
                );

            // Examples
            //inputExcel.RunDDL("insert into [Planilha1$] values (@v_txt_x)", false, "v_txt_x: z");
            //inputExcel.RunDDL("update [Planilha1$] set TESTE = 'oi' where TESTE = @v_txt_x", false, "v_txt_x: valor");

            // Test Variables
            int lgsID;
            string printFileName;
            string scenarioName = inputExcel.GetValue("SCENARIONAME", queryName);
            string testName = inputExcel.GetValue("TESTNAME", queryName);
            string testType = inputExcel.GetValue("TESTTYPE", queryName);
            string analystName = inputExcel.GetValue("ANALYSTNAME", queryName);
            string testDesc = inputExcel.GetValue("TESTDESC", queryName);
            int reportID = int.Parse(inputExcel.GetValue("REPORTID", queryName));

            Global.processTest.StartTest(Global.customerName, suiteName, scenarioName, testName, testType, analystName, testDesc, reportID);

            // Test Details
            string preCondition = inputExcel.GetValue("PRECONDITION", queryName);
            string postCondition = inputExcel.GetValue("POSTCONDITION", queryName);
            string inputData = inputExcel.GetValue("INPUTDATA", queryName);
            Global.processTest.DoTest(preCondition, postCondition, inputData);

            // Steps Definition
            DefineSteps("RealizarLogin");

            // Steps Execution
            Login(inputExcel, queryName);

            // Teardown function
            Global.processTest.EndTest(reportID, queryName);
        }
    }
}