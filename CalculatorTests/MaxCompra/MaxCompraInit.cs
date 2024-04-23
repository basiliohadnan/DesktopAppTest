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

            ////
            //PressEnter();

            //SetAppSession("Centura:MDIFrame");
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
            string scenarioName = "Relatorio modelo login";
            int reportID = 1;
            int lgsID;
            string printFileName;
            // name
            string expectedResult = "Não foi definido a versão do módulo no BANCO DE DADOS, para o sistema de Segurança";

            Console.WriteLine("Test Variables");
            string testName = "Login";
            string testType = "Funcional";
            string analystName = "Hadnan Basilio";
            string testDesc = "Realizar login no MaxCompra com matricula do analista";
            Global.processTest.StartTest(Global.customerName, suiteName, scenarioName, testName, testType, analystName, testDesc, reportID);

            Console.WriteLine("Test Details");
            string preCondition = "Caixa aberto";
            string postCondition = "Nenhuma";
            string inputData = "Código de produto válido";
            Global.processTest.DoTest(preCondition, postCondition, inputData);

            Console.WriteLine("Steps Definition");
            Global.processTest.DoStep("Abrir app", "Abertura do app com sucesso");
            Global.processTest.DoStep("Login do analista", "Login com sucesso");
            Global.processTest.DoStep("Tela final", expectedResult);

            Console.WriteLine("Steps Execution");

            lgsID = Global.processTest.StartStep("Login do analista", logMsg: "Tentando login", paramName: "matricula", paramValue: user);
            // Step actions
            Login();

            printFileName = Global.processTest.PrintScreen();
            if (true)
            {
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: "login efetuado");
            }
            else
            {
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "erro no login");
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
