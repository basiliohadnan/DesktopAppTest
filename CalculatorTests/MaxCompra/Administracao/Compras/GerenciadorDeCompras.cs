using Consinco.Helpers;
using Consinco.MaxCompra;
using DesktopAppTests.MaxCompra.PageObjects.Administracao.Compras;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium;
using Starline;
using System.Windows;
using OpenQA.Selenium.Appium;
using System.Diagnostics;
using OpenCvSharp;
using System.Security.Cryptography;

namespace DesktopAppTests.MaxCompra.Administracao.Compras
{
    [TestClass]
    public class GerenciadorDeComprasTests : MaxCompraInit
    {
        private GerenciadorDeComprasPO gerenciadorDeComprasPO;
        private OCRScanner scan = new OCRScanner();

        [TestMethod]
        public void CreateLoteLojaALoja()
        {
            //Console.WriteLine("Global Variables");
            string scenarioName = "Criar lote - loja a loja";
            int reportID = 3;
            string printFileName = "";
            bool doAction = true;
            string printPath;

            LoginWithReport();
            OpenMenu("Administração", "02-OpenMenuAdm");
            OpenMenu("Compras", "03-OpenSubMenuCompras");
            OpenMenu("Gerenciador de Compras", "04-OpenSubMenuGerenciadorDeCompras");

            //Console.WriteLine("Test Variables");
            int fornecedor = 478;
            int qtdLojas = 11;
            string categoria = "LIQ2";
            int diasAbastecimento = 60;
            string checkBoxesClass = "Centura:GPCheck";
            string comprador = "LIQ1";
            int qtdeCompra = 12;
            int qtdProdutos = 4;

            gerenciadorDeComprasPO = new GerenciadorDeComprasPO(new ElementHandler());
            gerenciadorDeComprasPO.FillFornecedor(fornecedor);
            gerenciadorDeComprasPO.AddLojas(qtdLojas);
            gerenciadorDeComprasPO.SelectCategoria(categoria);
            gerenciadorDeComprasPO.FillAbastecimentoDias(diasAbastecimento);
            gerenciadorDeComprasPO.EnableCheckBoxesSugestaoDeCompras(checkBoxesClass);
            gerenciadorDeComprasPO.IncluirLote();
            gerenciadorDeComprasPO.AddCompradores(comprador);
            gerenciadorDeComprasPO.ConfirmProdutosInativosWindow();
            gerenciadorDeComprasPO.ConfirmTributacaoWindow();
            gerenciadorDeComprasPO.FillQtdeCompra(qtdLojas, qtdProdutos, qtdeCompra);
            gerenciadorDeComprasPO.ClickGerarPedidos();
        }
        //public void restante()
        //{

        //    string testName = "Pedido de compra - loja a loja";
        //    string testType = "Funcional";
        //    string analystName = "Hadnan Basilio";
        //    string testDesc = "Criar lote de pedido de compras - loja a loja";

        //    Global.processTest.StartTest(customerName, suiteName, scenarioName, testName, testType, analystName, testDesc, reportID);

        //    //Console.WriteLine("Test Details");
        //    string preCondition = "Matricula valida, app abrindo normalmente";
        //    string postCondition = "Pedido gerado com sucesso";
        //    string inputData = "Dados fornecedor, comprador e parâmetros do pedido";
        //    Global.processTest.DoTest(preCondition, postCondition, inputData);

        //    //Console.WriteLine("Steps Definition");
        //    Global.processTest.DoStep("Abertura da aplicação", "Tela de login exibida");
        //    Global.processTest.DoStep("Login da aplicação", "Aplicação aberta");
        //    Global.processTest.DoStep("Alerta de versão de aplicação", "Alerta confirmado");
        //    Global.processTest.DoStep("Aplicação disponível", "Aplicação aberta e com acesso aos menus");
        //    Global.processTest.DoStep("Abertura do menu de telas de recebimento", "Itens de menu exibidos");
        //    Global.processTest.DoStep("Abertura da tela de consulta", "Tela de relatório exibida");
        //    Global.processTest.DoStep("Definição dos filtros do relatório", "Filtros configurados");
        //    Global.processTest.DoStep("Consulta do relatório", "Relatório consultado");
        //    Global.processTest.DoStep("Iteração nos itens do relatório", "Relatório contabilizado");

        //    //Console.WriteLine("Steps Execution");
        //    int appLeftOffset = 2;
        //    int appTopOffset = 0;
        //    int appLeft = 0;
        //    int appTop = 0;
        //    int appWidth = (int)SystemParameters.FullPrimaryScreenWidth;
        //    int appHeight = (int)SystemParameters.FullPrimaryScreenHeight;
        //    string errorLogMsg = "WinAppDriver";
        //    string exePath = appPath;

        //    if (doAction)
        //    {
        //        try
        //        {
        //            // Create session at root level
        //            WindowsDriver<WindowsElement> winSession;
        //            AppiumOptions rootCapabilities = new AppiumOptions();
        //            //rootCapabilities.AddAdditionalCapability("app", "C:\\Starline\\Desktop\\AppGrid\\ExportTable.exe");
        //            rootCapabilities.AddAdditionalCapability("app", "Root");
        //            winSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), rootCapabilities);

        //            lgsID = Global.processTest.StartStep("Abertura da aplicação", logMsg: "Abrindo executável", paramName: "exePath", paramValue: exePath);
        //            errorLogMsg = "Erro na abertura da aplicação";
        //            printFileName = "01-login_open.png";
        //            var process = Process.Start(exePath);
        //            Global.processTest.Wait(2000);
        //            var RootWindow = winSession.FindElementByName("Conexão de Sistemas Consinco");
        //            appLeft = RootWindow.Rect.Left + appLeftOffset;
        //            appTop = RootWindow.Rect.Top + appTopOffset;
        //            appWidth = RootWindow.Rect.Width;
        //            appHeight = RootWindow.Rect.Height;
        //            printPath = Global.processTest.PrintScreen(false, 0, appLeft, appTop, appWidth, appHeight);
        //            Global.processTest.EndStep(lgsID, printPath: printPath + printFileName, logMsg: "Login exibido");

        //            lgsID = Global.processTest.StartStep("Login da aplicação", logMsg: "Efetuando login", paramName: "matricula", paramValue: matricula);
        //            errorLogMsg = "Erro preenchendo credenciais de login";
        //            printFileName = "02-login_credentials.png";
        //            // Create session by attaching to App top level window
        //            WindowsDriver<WindowsElement> appSession;
        //            AppiumOptions appCapabilities = new AppiumOptions();
        //            var RootTopLevelWindowHandle = RootWindow.GetAttribute("NativeWindowHandle");
        //            RootTopLevelWindowHandle = (int.Parse(RootTopLevelWindowHandle)).ToString("x"); // Convert to Hex
        //            appCapabilities.AddAdditionalCapability("appTopLevelWindow", RootTopLevelWindowHandle);
        //            appSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);
        //            var loginEdits = appSession.FindElementsByClassName("Edit");
        //            loginEdits[1].SendKeys(matricula);
        //            Global.processTest.Wait(500);
        //            printPath = Global.processTest.PrintScreen(false, 0, appLeft, appTop, appWidth, appHeight);
        //            Global.processTest.EndStep(lgsID, printPath: printPath + printFileName, logMsg: "Credenciais inseridas");

        //            lgsID = Global.processTest.StartStep("Alerta de versão de aplicação", logMsg: "Validando a versão", paramName: "", paramValue: "");
        //            errorLogMsg = "Erro no login da aplicação";
        //            printFileName = "03-app_open.png";
        //            var loginButtons = appSession.FindElementsByClassName("Button");
        //            loginButtons[0].Click();
        //            Global.processTest.Wait(500);
        //            var centuraWindow = winSession.FindElementByClassName("Centura:MDIFrame");
        //            appLeftOffset = 8;
        //            appTopOffset = 8;
        //            appLeft = centuraWindow.Rect.Left + appLeftOffset;
        //            appTop = centuraWindow.Rect.Top + appTopOffset;
        //            appWidth = centuraWindow.Rect.Width;
        //            appHeight = centuraWindow.Rect.Height;
        //            printPath = Global.processTest.PrintScreen(false, 0, appLeft, appTop, appWidth, appHeight);
        //            Global.processTest.EndStep(lgsID, printPath: printPath + printFileName, logMsg: "Alerta de versão exibido");

        //            lgsID = Global.processTest.StartStep("Aplicação disponível", logMsg: "Fechando alerta de validação de versão", paramName: "", paramValue: "");
        //            errorLogMsg = "Erro confirmando alerta de versão";
        //            printFileName = "04-app_alert_closed.png";
        //            var centuraTopLevelWindowHandle = centuraWindow.GetAttribute("NativeWindowHandle");
        //            centuraTopLevelWindowHandle = (int.Parse(centuraTopLevelWindowHandle)).ToString("x"); // Convert to Hex
        //            WindowsDriver<WindowsElement> centuraSession;
        //            AppiumOptions centuraCapabilities = new AppiumOptions();
        //            centuraCapabilities.AddAdditionalCapability("appTopLevelWindow", centuraTopLevelWindowHandle);
        //            centuraSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), centuraCapabilities);
        //            var alertButton = centuraSession.FindElementsByName("OK");
        //            alertButton[0].Click();
        //            Global.processTest.Wait(500);
        //            printPath = Global.processTest.PrintScreen(false, 0, appLeft, appTop, appWidth, appHeight);
        //            Global.processTest.EndStep(lgsID, printPath: printPath + printFileName, logMsg: "Aplicação aberta e com acesso aos menus");

        //            //lgsID = Global.processTest.StartStep("Abertura do menu de telas de recebimento", logMsg: "Acessando menu de telas de recebimento", paramName: "", paramValue: "");
        //            //errorLogMsg = "Erro acessando menu";
        //            //printFileName = "05-menu_open.png";
        //            //var menuRecebimentoButton = centuraSession.FindElementsByName("Recebimento");
        //            //menuRecebimentoButton[0].Click();
        //            //Global.processTest.Wait(500);
        //            //printPath = Global.processTest.PrintScreen(false, 0, appLeft, appTop, appWidth, appHeight);
        //            //Global.processTest.EndStep(lgsID, printPath: printPath + printFileName, logMsg: "Itens de menu exibidos");

        //            //lgsID = Global.processTest.StartStep("Abertura da tela de consulta", logMsg: "Abrir tela de consulta", paramName: "", paramValue: "");
        //            //errorLogMsg = "Erro abrindo tela de relatório";
        //            //printFileName = "06-report_open.png";
        //            //var menuRecebimentoNFeButton = centuraSession.FindElementsByName("Recebimento NFe");
        //            //menuRecebimentoNFeButton[0].Click();
        //            //Global.processTest.Wait(500);
        //            //printPath = Global.processTest.PrintScreen(false, 0, appLeft, appTop, appWidth, appHeight);
        //            //Global.processTest.EndStep(lgsID, printPath: printPath + printFileName, logMsg: "Tela de relatório exibida");

        //            //lgsID = Global.processTest.StartStep("Definição dos filtros do relatório", logMsg: "Definindo filtros de consulta do relatório", paramName: "referenceMonth", paramValue: "01/2022");
        //            //errorLogMsg = "Erro no preenchimento dos filtros do relatório";
        //            //printFileName = "07-report_fill_fields.png";

        //            //var centuraEmpresa = centuraSession.FindElementsByClassName("ComboBox");
        //            //centuraEmpresa[0].SendKeys("144");

        //            //var centuraEdits = centuraSession.FindElementsByClassName("Edit");
        //            //centuraEdits[2].Clear();
        //            //centuraEdits[2].SendKeys("01/01/2022");
        //            //centuraEdits[3].Clear();
        //            //centuraEdits[3].SendKeys("10/01/2022");
        //            //Global.processTest.Wait(500);
        //            //printPath = Global.processTest.PrintScreen(false, 0, appLeft, appTop, appWidth, appHeight);
        //            //Global.processTest.EndStep(lgsID, printPath: printPath + printFileName, logMsg: "Filtros configurados");

        //            //lgsID = Global.processTest.StartStep("Consulta do relatório", logMsg: "Consultando filtros", paramName: "", paramValue: "");
        //            //errorLogMsg = "Erro executando relatório";
        //            //printFileName = "08-report_results.png";
        //            //var centuraButtons = centuraSession.FindElementsByClassName("Button");
        //            //centuraButtons[2].Click();
        //            //Global.processTest.Wait(500);
        //            //printPath = Global.processTest.PrintScreen(false, 0, appLeft, appTop, appWidth, appHeight);

        //            bool errorDetected = false;
        //            try
        //            {
        //                var errorAlert = winSession.FindElementByName("Erro na execução da instrução SQL");
        //                WindowsDriver<WindowsElement> alertSession;
        //                AppiumOptions alertCapabilities = new AppiumOptions();
        //                var AlertTopLevelWindowHandle = errorAlert.GetAttribute("NativeWindowHandle");
        //                AlertTopLevelWindowHandle = (int.Parse(AlertTopLevelWindowHandle)).ToString("x"); // Convert to Hex
        //                alertCapabilities.AddAdditionalCapability("appTopLevelWindow", AlertTopLevelWindowHandle);
        //                alertSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), alertCapabilities);
        //                errorDetected = true;
        //                Global.processTest.EndStep(lgsID, status: "erro", printPath: printPath + printFileName, logMsg: "Erro ao consultar relatório");
        //                Global.processTest.Wait(500);
        //                alertSession.Quit();
        //                Global.processTest.Wait(500);
        //            }
        //            catch
        //            {
        //                errorDetected = false;
        //            }

        //            if (!errorDetected)
        //            {
        //                Global.processTest.EndStep(lgsID, printPath: printPath + printFileName, logMsg: "Relatório consultado");

        //                lgsID = Global.processTest.StartStep("Iteração nos itens do relatório", logMsg: "Contabilizando registros do relatório", paramName: "", paramValue: "");
        //                errorLogMsg = "Erro na contabilização dos registros do relatório";
        //                printFileName = "09-report_check.png";

        //                // Find Grid element and enable edit on first row
        //                WindowsElement gridTable = centuraSession.FindElementByClassName("Centura:ChildTable");

        //                //string teste = gridTable.GetAttribute("Grid.ColumnCount").ToString();
        //                centuraSession.Mouse.MouseMove(gridTable.Coordinates, 140, 45);
        //                centuraSession.Mouse.Click(null);

        //                // campo CGO -> 32 x 17
        //                // 511 x 255
        //                // 511 x 273

        //                int gridLeftOffset = 0;
        //                int gridTopOffset = 0;

        //                int gridLeft = gridTable.Rect.Left + gridLeftOffset;
        //                int gridTop = gridTable.Rect.Top + gridTopOffset;
        //                int gridWidth = gridTable.Rect.Width;
        //                int gridHeight = gridTable.Rect.Height;

        //                ScreenPrinter.CustomPrintScreen(printPath + printFileName, false, 0, gridLeft, gridTop, gridWidth, gridHeight);

        //                // Edit element
        //                //WindowsElement editField;
        //                int fieldWidth = 17;
        //                int fieldHeight = 32;
        //                int fieldLine = 1;
        //                int fieldLeft = appLeft + 511;
        //                int fieldTop = appTop + 255 - fieldHeight;
        //                int fieldMovement = 17;
        //                string fieldValue = "";
        //                string txt = "";


        //                // Scroll element
        //                WindowsElement scrollBar = appSession.FindElementByName("Vertical");
        //                WindowsElement downButton = appSession.FindElementByName("Uma linha abaixo");
        //                int pos = -1;
        //                int max = int.Parse(scrollBar.GetAttribute("RangeValue.Maximum").ToString());
        //                int field = 1;
        //                int counter = 0;

        //                string fullFilename = "";
        //                string fullFilenameOCR = "";

        //                // Iterate rows
        //                while (pos < max)
        //                {
        //                    printPath = Global.processTest.GetAppPath() + "/Prints" + "/" + Global.processTest.CustomerName + "/" + Global.processTest.RptID.ToString() + "/" + Global.processTest.SuiteName + "/" + Global.processTest.ScenarioName + "/" + Global.processTest.TestName;
        //                    fullFilename = printPath + "/" + "field-" + field.ToString().PadLeft(4, '0') + ".png";
        //                    fullFilenameOCR = printPath + "/" + "field-ocr-" + field.ToString().PadLeft(4, '0') + ".png";
        //                    //fullFilename = printPath + "/" + "field.png";
        //                    //fullFilenameOCR = printPath + "/" + "field-ocr.png";

        //                    if (field < fieldMovement)
        //                    {
        //                        ScreenPrinter.CustomPrintScreen(fullFilename, false, 0, fieldLeft, fieldTop + ((fieldHeight + fieldLine) * field), fieldWidth, fieldHeight);
        //                    }
        //                    else
        //                    {
        //                        ScreenPrinter.CustomPrintScreen(fullFilename, false, 0, fieldLeft, fieldTop + ((fieldHeight + fieldLine) * fieldMovement), fieldWidth, fieldHeight);
        //                    }

        //                    fieldValue = scan.FieldValidator(fullFilename, fullFilenameOCR, 50, 4);
        //                    using (System.IO.StreamWriter outputFile = new System.IO.StreamWriter(@"C:\Starline\Desktop\WinAppDriver\UnitTestProject1\Prints\table.txt", true))
        //                    {
        //                        outputFile.WriteLine(fieldValue);
        //                    }

        //                    //editField = appSession.FindElementByClassName("Edit");
        //                    //fieldValue = editField.GetAttribute("Value.Value").ToString();

        //                    txt = fieldValue.Substring(0, 7);
        //                    if (txt == "String5")
        //                    {
        //                        counter++;
        //                    }

        //                    pos = int.Parse(scrollBar.GetAttribute("RangeValue.Value").ToString());
        //                    appSession.Keyboard.SendKeys(Keys.ArrowDown);
        //                    //editField.SendKeys(Keys.ArrowDown);
        //                    //downButton.Click();
        //                    field++;
        //                }
        //                using (StreamWriter outputFile = new System.IO.StreamWriter(@"C:\Starline\Desktop\WinAppDriver\UnitTestProject1\Prints\table.txt", true))
        //                {
        //                    outputFile.WriteLine("Resultado: " + counter);
        //                }
        //                Global.processTest.EndStep(lgsID, printPath: printPath + printFileName, logMsg: "Relatório contabilizado");
        //            }


        //            try
        //            {
        //                centuraSession.Quit();
        //                process.Kill();
        //            }
        //            catch
        //            {

        //            }

        //        }
        //        catch (Exception ex)
        //        {
        //            printPath = Global.processTest.PrintScreen(false, 0, appLeft, appTop, appWidth, appHeight);
        //            Global.processTest.EndStep(lgsID, status: "erro", printPath: printPath + printFileName, logMsg: "Erro");
        //            Console.WriteLine("errorLogMsg: " + errorLogMsg);
        //            Console.WriteLine(ex.Message);
        //        }
        //    }
        //    else
        //    {
        //        try
        //        {
        //            lgsID = Global.processTest.StartStep("Abertura da aplicação", logMsg: "Abrindo executável", paramName: "exePath", paramValue: exePath);
        //            printFileName = "01-login_open.png";
        //            printPath = Global.processTest.PrintScreen(false, 0, appLeft, appTop, appWidth, appHeight);
        //            Global.processTest.EndStep(lgsID, printPath: printPath + printFileName, logMsg: "Login exibido");

        //            lgsID = Global.processTest.StartStep("Login da aplicação", logMsg: "Efetuando login", paramName: "matricula", paramValue: matricula);
        //            printFileName = "02-login_credentials.png";
        //            printPath = Global.processTest.PrintScreen(false, 0, appLeft, appTop, appWidth, appHeight);
        //            Global.processTest.EndStep(lgsID, printPath: printPath + printFileName, logMsg: "Credenciais inseridas");

        //            lgsID = Global.processTest.StartStep("Alerta de versão de aplicação", logMsg: "Validando a versão", paramName: "", paramValue: "");
        //            printFileName = "03-app_open.png";
        //            printPath = Global.processTest.PrintScreen(false, 0, appLeft, appTop, appWidth, appHeight);
        //            Global.processTest.EndStep(lgsID, printPath: printPath + printFileName, logMsg: "Alerta de versão exibido");

        //            lgsID = Global.processTest.StartStep("Aplicação disponível", logMsg: "Fechando alerta de validação de versão", paramName: "", paramValue: "");
        //            printFileName = "04-app_alert_closed.png";
        //            printPath = Global.processTest.PrintScreen(false, 0, appLeft, appTop, appWidth, appHeight);
        //            Global.processTest.EndStep(lgsID, printPath: printPath + printFileName, logMsg: "Aplicação aberta e com acesso aos menus");

        //            lgsID = Global.processTest.StartStep("Abertura do menu de telas de recebimento", logMsg: "Acessando menu de telas de recebimento", paramName: "", paramValue: "");
        //            printFileName = "05-menu_open.png";
        //            printPath = Global.processTest.PrintScreen(false, 0, appLeft, appTop, appWidth, appHeight);
        //            Global.processTest.EndStep(lgsID, printPath: printPath + printFileName, logMsg: "Itens de menu exibidos");

        //            lgsID = Global.processTest.StartStep("Abertura da tela de consulta", logMsg: "Abrir tela de consulta", paramName: "", paramValue: "");
        //            printFileName = "06-report_open.png";
        //            printPath = Global.processTest.PrintScreen(false, 0, appLeft, appTop, appWidth, appHeight);
        //            Global.processTest.EndStep(lgsID, printPath: printPath + printFileName, logMsg: "Tela de relatório exibida");

        //            lgsID = Global.processTest.StartStep("Definição dos filtros do relatório", logMsg: "Definindo filtros de consulta do relatório", paramName: "referenceMonth", paramValue: "01/2022");
        //            printFileName = "07-report_fill_fields.png";
        //            printPath = Global.processTest.PrintScreen(false, 0, appLeft, appTop, appWidth, appHeight);
        //            Global.processTest.EndStep(lgsID, printPath: printPath + printFileName, logMsg: "Filtros configurados");

        //            lgsID = Global.processTest.StartStep("Consulta do relatório", logMsg: "Consultando filtros", paramName: "", paramValue: "");
        //            printFileName = "08-report_results.png";
        //            printPath = Global.processTest.PrintScreen(false, 0, appLeft, appTop, appWidth, appHeight);
        //            Global.processTest.EndStep(lgsID, status: "erro", printPath: printPath + printFileName, logMsg: "Erro ao consultar relatório");
        //            //processTest.EndStep(lgsID, printPath: printPath + printFileName, logMsg: "Relatório consultado");

        //            //lgsID = Global.processTest.StartStep("Iteração nos itens do relatório", logMsg: "Contabilizando registros do relatório", paramName: "", paramValue: "");
        //            //printFileName = "09-report_check.png";
        //            //processTest.CustomPrintScreen(printPath + printFileName, false, 0, appLeft, appTop, appWidth, appHeight);
        //            //processTest.EndStep(lgsID, printPath: printPath + printFileName, logMsg: "Relatório contabilizado");
        //        }
        //        catch (Exception ex)
        //        {
        //            printPath = Global.processTest.PrintScreen(false, 0, appLeft, appTop, appWidth, appHeight);
        //            Global.processTest.EndStep(lgsID, status: "erro", printPath: printPath + printFileName, logMsg: "Erro");
        //            Console.WriteLine("errorLogMsg: " + errorLogMsg);
        //            Console.WriteLine(ex.Message);
        //        }
        //    }
        //    //Console.WriteLine("Teardown function");
        //    Global.processTest.EndTest(reportID);
        //    //Console.WriteLine("Test finished");
        //    Console.ReadKey();

        //}
    }
}