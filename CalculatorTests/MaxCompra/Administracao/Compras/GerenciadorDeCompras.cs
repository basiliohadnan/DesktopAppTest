using Consinco.Helpers;
using Consinco.MaxCompra.PageObjects.Administracao.Compras;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OpenQA.Selenium.Appium.Windows;

namespace Consinco.MaxCompra.Administracao.Compras
{
    [TestClass]
    public class GerenciadorDeComprasTests : MaxCompraInit
    {
        private GerenciadorDeComprasPO gerenciadorDeComprasPO;
        private OCRScanner scan = new OCRScanner();

        [TestMethod]
        public void CriarLoteDeCompraLojaALojaUmaLoja()
        {
            // Global Variables
            int rowNumber = 2;
            string worksheetName = "GerenciadorDeCompras";
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

            // Steps Definition");
            Global.processTest.DoStep("Abrir app", "Abertura do app com sucesso");
            Global.processTest.DoStep("Login do analista", "Login com sucesso");
            Global.processTest.DoStep("Tela final", "Tela principal exibida com sucesso");
            Global.processTest.DoStep("Abrir menu Administração", "Menu Administracao aberto com sucesso");
            Global.processTest.DoStep("Abrir menu Compras", "Menu Compras aberto com sucesso");
            Global.processTest.DoStep("Abrir menu Gerenciador de Compras", "Gerenciador de Compras aberto com sucesso");
            Global.processTest.DoStep("Preencher fornecedor", "Fornecedor preenchido com sucesso");
            //Global.processTest.DoStep("Adicionar lojas", "Adição de lojas com sucesso");
            //Global.processTest.DoStep("Confirmar janela Seleção de Empresas do Lote", "Confirmação janela Seleção de Empresas do Lote com sucesso");
            Global.processTest.DoStep("Selecionar categoria", "Seleção da categoria com sucesso");
            Global.processTest.DoStep("Preencher dias abastecimento", "Preenchimento dias abastecimento com sucesso");
            Global.processTest.DoStep("Habilitar checkboxes Sugestão de compra", "Habilitação checkboxes sugestão de compra com sucesso");
            Global.processTest.DoStep("Incluir lote de compra", "Inclusão do lote de compra com sucesso");
            Global.processTest.DoStep("Adicionar comprador", "Inclusão de comprador com sucesso");
            Global.processTest.DoStep("Confirmar janela Seleção de Produtos", "Confirmação janela Seleção de Produtos com sucesso");
            //Global.processTest.DoStep("Confirmar janela Produtos Inativos", "Confirmação janela Produtos Inativos com sucesso");
            //Global.processTest.DoStep("Confirmar janela Tributação", "Confirmação janela Tributação com sucesso");
            Global.processTest.DoStep("Preencher quantidade de compra por loja e produto",
                "Preenchimento quantidade de compra por loja e produto com sucesso");
            Global.processTest.DoStep("Gerar Pedidos", "Pedidos gerados com sucesso");
            Global.processTest.DoStep("Confirmar Consulta do lote", "Lote gerado com sucesso");

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
            Login(matricula);
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

            string menuName = "Administração";
            lgsID = Global.processTest.StartStep($"Abrir menu {menuName}", logMsg: $"Abrir menu {menuName}",
                paramName: "menuName", paramValue: menuName);
            try
            {
                OpenMenu(menuName);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"menu {menuName} aberto com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"erro na abertura do menu {menuName}");
            }

            menuName = "Compras";
            lgsID = Global.processTest.StartStep($"Abrir menu {menuName}", logMsg: $"Abrir menu {menuName}", paramName: "menuName",
                paramValue: menuName);
            try
            {
                OpenMenu(menuName);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"menu {menuName} aberto com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"erro na abertura do menu {menuName}");
            }

            menuName = "Gerenciador de Compras";
            lgsID = Global.processTest.StartStep($"Abrir menu {menuName}", logMsg: $"Abrir menu {menuName}", paramName: "menuName",
                paramValue: menuName);
            try
            {
                OpenMenu(menuName);
                WaitSeconds(1);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"menu {menuName} aberto com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"erro na abertura do menu {menuName}");
            }

            string codFornecedor = excelReader.ReadCellValue(worksheet, "codFornecedor", rowNumber);
            lgsID = Global.processTest.StartStep("Preencher fornecedor", logMsg: $"Preencher fornecedor {codFornecedor}",
                paramName: "fornecedorId", paramValue: codFornecedor);
            GerenciadorDeComprasPO gerenciadorDeComprasPO = new GerenciadorDeComprasPO(new ElementHandler());
            try
            {
                gerenciadorDeComprasPO.FillFornecedor(codFornecedor);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: "Fornecedor preenchido com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "erro no preenchimento do fornecedor");
            }

            //lgsID = Global.processTest.StartStep("Adicionar lojas", logMsg: $"Adicionar {qtdLojas} lojas", paramName: "qtdLojas",
            //    paramValue: qtdLojas.ToString());
            //try
            //{
            //    gerenciadorDeComprasPO.AddLojas(qtdLojas);
            //    int startX = 50;
            //    int startY = 60;
            //    int endX = 470;
            //    int endY = 591;
            //    printFileName = Global.processTest.PrintScreen(fullScreen: false, startX: startX, startY: startY, endX: endX, endY: endY);
            //    Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: "Adição de lojas com sucesso");
            //}
            //catch
            //{
            //                printFileName = Global.processTest.PrintScreen();
            //    Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "erro na adição de lojas");
            //}

            //string windowName = "Seleção de Empresas do Lote";
            //lgsID = Global.processTest.StartStep($"Confirmar janela {windowName}", logMsg: $"Tentando confirmar janela {windowName}",
            //    paramName: "windowName", paramValue: windowName);
            //try
            //{
            //    gerenciadorDeComprasPO.ConfirmSelecaoDeLojasWindow();
            //    // Verificar motivo estar salvando screenshot após seleção da categoria.
            //    //int startX = 50;
            //    //int startY = 60;
            //    //int endX = 470;
            //    //int endY = 591;
            //    //printFileName = Global.processTest.PrintScreen(fullScreen: false, startX: startX, startY: startY, endX: endX, endY: endY);
            //    printFileName = Global.processTest.PrintScreen();
            //    Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Confirmação janela {windowName} com sucesso");
            //}
            //catch
            //{
            //                printFileName = Global.processTest.PrintScreen();
            //    Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"erro na confirmação janela {windowName}");
            //}

            string categoria = excelReader.ReadCellValue(worksheet, "categoria", rowNumber);
            lgsID = Global.processTest.StartStep("Selecionar categoria", logMsg: $"Selecionar categoria {categoria}",
                paramName: "categoria", paramValue: categoria);
            try
            {
                gerenciadorDeComprasPO.SelectCategoria(categoria);
                WaitSeconds(1);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: "Seleção da categoria com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "erro na seleção da categoria");
            }

            string diasAbastecimento = excelReader.ReadCellValue(worksheet, "diasAbastecimento", rowNumber);
            lgsID = Global.processTest.StartStep("Preencher dias abastecimento", logMsg: $"Preencher dias abastecimento com {diasAbastecimento}",
                paramName: "diasAbastecimento", paramValue: diasAbastecimento);
            try
            {
                gerenciadorDeComprasPO.FillAbastecimentoDias(diasAbastecimento);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: "Preenchimento dias abastecimento com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "erro no preenchimento de dias abastecimento");
            }

            lgsID = Global.processTest.StartStep("Habilitar checkboxes Sugestão de compra", logMsg: $"Habilitar checkboxes sugestão de compras",
                paramName: "", paramValue:"");
            try
            {
                gerenciadorDeComprasPO.EnableCheckBoxesSugestaoDeCompras();
                // Verificar motivo estar salvando screenshot após incluir lote.
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: "Habilitação checkboxes sugestão de compra com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName,
                    logMsg: "erro na habilitação dos checkboxes de Sugestão de compra");
            }

            lgsID = Global.processTest.StartStep("Incluir lote de compra", logMsg: $"Incluir lote de compra", paramName: "", paramValue: "");
            try
            {
                gerenciadorDeComprasPO.IncluirLote();
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: "Inclusão do lote de compra com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "erro na inclusão do lote de compra");
            }

            string comprador = excelReader.ReadCellValue(worksheet, "comprador", rowNumber);
            lgsID = Global.processTest.StartStep("Adicionar comprador", logMsg: $"Adicionar comprador {comprador}", paramName: "comprador",
                paramValue: comprador);
            try
            {
                gerenciadorDeComprasPO.AddCompradores(comprador);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: "Inclusão de comprador com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "erro na inclusão de comprador");
            }

            string windowName = excelReader.ReadCellValue(worksheet, "windowName", rowNumber);
            lgsID = Global.processTest.StartStep($"Confirmar janela {windowName}", logMsg: $"Tentando confirmar janela {windowName}",
                paramName: "windowName", paramValue: windowName);
            try
            {
                gerenciadorDeComprasPO.ConfirmSelecaoDeProdutosWindow();
                //string produtosInativosWindowName = "Produtos Inativos";
                //WindowsElement foundWindow = elementHandler.FindElementByName(produtosInativosWindowName, 120 * 1000 / 10);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Confirmação janela {windowName} com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"erro na confirmação janela {windowName}");
            }

            //windowName = "Produtos Inativos";
            //lgsID = Global.processTest.StartStep($"Confirmar janela {windowName}", logMsg: $"Tentando confirmar janela {windowName}",
            //    paramName: "windowName", paramValue: windowName);
            //try
            //{
            //    gerenciadorDeComprasPO.ConfirmProdutosInativosWindow();
            //    //printFileName = Global.processTest.PrintScreen();
            //    int startX = 53;
            //    int startY = 63;
            //    int endX = 677;
            //    int endY = 86;
            //    printFileName = Global.processTest.PrintScreen(fullScreen: false, startX: startX, startY: startY, endX: endX, endY: endY);
            //    Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Confirmação janela {windowName} com sucesso");
            //}
            //catch
            //{
            //  printFileName = Global.processTest.PrintScreen();
            //    Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"erro na confirmação janela {windowName}");
            //}

            //windowName = "Tributação";
            //lgsID = Global.processTest.StartStep($"Confirmar janela {windowName}", logMsg: $"Tentando confirmar janela {windowName}",
            //    paramName: "windowName", paramValue: windowName);
            //try
            //{
            //    gerenciadorDeComprasPO.ConfirmTributacaoWindow();
            //    int startX = 2;
            //    int startY = 45;
            //    int endX = 1023;
            //    int endY = 785;
            //    printFileName = Global.processTest.PrintScreen(fullScreen: false, startX: startX, startY: startY, endX: endX, endY: endY);
            //    Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Confirmação janela {windowName} com sucesso");
            //}
            //catch
            //{
            //    printFileName = Global.processTest.PrintScreen();
            //    Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"erro na confirmação janela {windowName}");
            //}

            int qtdLojas = int.Parse(excelReader.ReadCellValue(worksheet, "qtdLojas", rowNumber));
            int qtdProdutos = int.Parse(excelReader.ReadCellValue(worksheet, "qtdProdutos", rowNumber));
            int qtdeCompra = int.Parse(excelReader.ReadCellValue(worksheet, "qtdeCompra", rowNumber));
            lgsID = Global.processTest.StartStep($"Preencher quantidade de compra por loja e produto",
                logMsg: $"Preencher quantidade de compra por loja e produto",
                paramName: "qtdLojas, qtdProdutos, qtdeCompra", paramValue: $"{qtdLojas}, {qtdProdutos}, {qtdeCompra}");
            try
            {
                gerenciadorDeComprasPO.FillQtdeCompraPorLoja(qtdProdutos, qtdeCompra);
                
                //Validação total de compra
                string qtdeComprValue = gerenciadorDeComprasPO.GetQtdeComprValue();
                int qtdeComprasValue = int.Parse(qtdeComprValue);
                if (qtdeComprasValue != qtdLojas * qtdProdutos * qtdeCompra)
                {
                    Console.Write($"Erro no preenchimento: qtdeComprasValue atual: {qtdeComprasValue}, Total esperado: {qtdLojas * qtdProdutos * qtdeCompra}");
                }
                //Maximize();
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Preenchimento quantidade de compra por loja e produto com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName,
                    logMsg: $"erro no preenchimento quantidade de compra por loja e produto");
            }

            lgsID = Global.processTest.StartStep($"Gerar Pedidos", logMsg: $"Tentando Gerar Pedidos", paramName: "", paramValue: "");
            try
            {
                gerenciadorDeComprasPO.ClickGerarPedidos();
                gerenciadorDeComprasPO.ConfirmPedidosWindow();
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Pedidos gerados com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName,
                    logMsg: $"erro ao clicar no botão Gera Pedidos");
            }

            lgsID = Global.processTest.StartStep($"Confirmar Consulta do lote", logMsg: $"Tentando confirmar Consulta do Lote", paramName: "", paramValue: "");
            try
            {
                gerenciadorDeComprasPO.ConfirmConsultaLoteCompraWindow();
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Lote gerado com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName,
                    logMsg: $"erro ao clicar no confirmar Lote");
            }
            //Teardown function
            PressEnter();
            Global.processTest.EndTest(reportID);
        }

        [TestMethod]
        public void CriarLoteDeCompraIncorporaCD()
        {
            // Global Variables
            int rowNumber = 4; // criar
            //string matricula = excelReader.ReadCellValue("Matricula", rowNumber);
            string scenarioName = "Gerenciador de Compras";
            int reportID = 6;
            int lgsID;
            string printFileName;
            string divisao = "SAO PAULO";
            List<string> lojas = [
                "001 Manilha",
                "002 Cs Verde",
                "003 S Miguel",
                "005 St Amaro",
                "006 Tatuape",
                "007 Santos",
                "008 V Sonia",
                "010 SBCampo",
                "011 Gu Dutra",
                "012 Jundiai",
                "013 Sorocaba",
                "915 CAJAMAR"
                ];
            int qtdeCompra = 100;
            int qtdProdutos = 5;
            string codFornecedor = "273"; // Quimica Amparo Ltda
            string categoria = "DPH1";
            string diasAbastecimento = "90";
            string cdNome = "915 CAJAMAR";

            // Test Variables
            string welcomeWindowName = "Conexão de Sistemas Consinco";
            string testName = "Criar lote de compras - Incorpora CD";
            string testType = "Funcional";
            string analystName = "Hadnan Basilio";
            string testDesc = $"{testName} com {lojas.Count} lojas, {qtdProdutos} produtos, qtdeCompra {qtdeCompra}, fornecedor {codFornecedor} e CD {cdNome}";
            Global.processTest.StartTest(Global.customerName, suiteName, scenarioName, testName, testType, analystName, testDesc, reportID);

            // Test Details
            string preCondition = "App iniciando";
            string postCondition = "Lote criado com sucesso";
            string inputData = "Nenhum";
            Global.processTest.DoTest(preCondition, postCondition, inputData);

            // Steps Definition
            Global.processTest.DoStep("Abrir app", "Abertura do app com sucesso");
            Global.processTest.DoStep("Login do analista", "Login com sucesso");
            Global.processTest.DoStep("Tela final", "Tela principal exibida com sucesso");
            Global.processTest.DoStep("Abrir menu Administração", "Menu Administracao aberto com sucesso");
            Global.processTest.DoStep("Abrir menu Compras", "Menu Compras aberto com sucesso");
            Global.processTest.DoStep("Abrir menu Gerenciador de Compras", "Gerenciador de Compras aberto com sucesso");
            Global.processTest.DoStep("Preencher fornecedor", "Fornecedor preenchido com sucesso");
            Global.processTest.DoStep("Adicionar lojas", "Adição de lojas com sucesso");
            Global.processTest.DoStep("Confirmar janela Seleção de Empresas do Lote", "Confirmação janela Seleção de Empresas do Lote com sucesso");
            Global.processTest.DoStep("Selecionar categoria", "Seleção da categoria com sucesso");
            Global.processTest.DoStep("Habilitar Incorporar Sugestão", "Habilitação do Incorporar Sugestão com sucesso");
            Global.processTest.DoStep("Preencher dias abastecimento", "Preenchimento dias abastecimento com sucesso");
            Global.processTest.DoStep("Habilitar checkboxes Sugestão de compra", "Habilitação checkboxes sugestão de compra com sucesso");
            Global.processTest.DoStep("Incluir lote de compra", "Inclusão do lote de compra com sucesso");
            Global.processTest.DoStep("Confirmar janela Seleção de Produtos", "Confirmação janela Seleção de Produtos com sucesso");
            //Global.processTest.DoStep("Confirmar janela Produtos Inativos", "Confirmação janela Produtos Inativos com sucesso");
            //Global.processTest.DoStep("Preencher quantidade de compra por produto e quantidade",
            //    "Preenchimento da quantidade de compra por produto e quantidade com sucesso");
            //Global.processTest.DoStep("Gerar Pedidos", "Pedidos gerados com sucesso");
            //Global.processTest.DoStep("Confirmar Consulta do lote", "Lote gerado com sucesso");

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

            string menuName = "Administração";
            lgsID = Global.processTest.StartStep($"Abrir menu {menuName}", logMsg: $"Abrir menu {menuName}",
                paramName: "menuName", paramValue: menuName);
            try
            {
                OpenMenu(menuName);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"menu {menuName} aberto com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"erro na abertura do menu {menuName}");
            }

            menuName = "Compras";
            lgsID = Global.processTest.StartStep($"Abrir menu {menuName}", logMsg: $"Abrir menu {menuName}", paramName: "menuName",
                paramValue: menuName);
            try
            {
                OpenMenu(menuName);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"menu {menuName} aberto com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"erro na abertura do menu {menuName}");
            }

            menuName = "Gerenciador de Compras";
            lgsID = Global.processTest.StartStep($"Abrir menu {menuName}", logMsg: $"Abrir menu {menuName}", paramName: "menuName",
                paramValue: menuName);
            try
            {
                OpenMenu(menuName);
                WaitSeconds(1);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"menu {menuName} aberto com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"erro na abertura do menu {menuName}");
            }

            gerenciadorDeComprasPO = new GerenciadorDeComprasPO(new ElementHandler());
            lgsID = Global.processTest.StartStep("Preencher fornecedor", logMsg: $"Preencher fornecedor {codFornecedor}",
                paramName: "fornecedorId", paramValue: codFornecedor);
            try
            {
                gerenciadorDeComprasPO.FillFornecedor(codFornecedor);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: "Fornecedor preenchido com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "erro no preenchimento do fornecedor");
            }

            lgsID = Global.processTest.StartStep("Adicionar lojas", logMsg: $"Adicionar {lojas.Count} lojas", paramName: "lojas",
                paramValue: "");
            try
            {
                gerenciadorDeComprasPO.OpenSelecaoDeLojas();
                gerenciadorDeComprasPO.RemoveDivisoes();
                gerenciadorDeComprasPO.AddDivisao(divisao);
                gerenciadorDeComprasPO.RemoveLojas();
                gerenciadorDeComprasPO.AddLojasPorNome(lojas);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: "Adição de lojas com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "erro na adição de lojas");
            }

            string windowName = "Seleção de Empresas do Lote";
            lgsID = Global.processTest.StartStep($"Confirmar janela {windowName}", logMsg: $"Tentando confirmar janela {windowName}",
                paramName: "windowName", paramValue: windowName);
            try
            {
                gerenciadorDeComprasPO.ConfirmSelecaoDeLojasWindow();
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Confirmação janela {windowName} com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"erro na confirmação janela {windowName}");
            }

            lgsID = Global.processTest.StartStep("Selecionar categoria", logMsg: $"Selecionar categoria {categoria}",
                paramName: "categoria", paramValue: categoria);
            try
            {
                gerenciadorDeComprasPO.SelectCategoria(categoria);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: "Seleção da categoria com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "erro na seleção da categoria");
            }

            lgsID = Global.processTest.StartStep("Habilitar Incorporar Sugestão", logMsg: $"Habilitar Incorporar Sugestão para o CD {cdNome}",
                paramName: "cdNome", paramValue: cdNome);
            try
            {
                gerenciadorDeComprasPO.EnableCheckBoxIncorporarSugestao();
                gerenciadorDeComprasPO.SetCD(cdNome);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: "Habilitação do Incorporar Sugestão com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "erro na habilitação do Incorporar Sugestão");
            }


            lgsID = Global.processTest.StartStep("Preencher dias abastecimento", logMsg: $"Preencher dias abastecimento com {diasAbastecimento}",
                paramName: "diasAbastecimento", paramValue: diasAbastecimento);
            try
            {
                gerenciadorDeComprasPO.FillAbastecimentoDias(diasAbastecimento);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: "Preenchimento dias abastecimento com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "erro no preenchimento de dias abastecimento");
            }

            lgsID = Global.processTest.StartStep("Habilitar checkboxes Sugestão de compra", logMsg: $"Habilitar checkboxes sugestão de compras",
                paramName: "", paramValue: "");
            try
            {
                gerenciadorDeComprasPO.EnableCheckBoxesSugestaoDeCompras();
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: "Habilitação checkboxes sugestão de compra com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName,
                    logMsg: "erro na habilitação dos checkboxes de Sugestão de compra");
            }

            lgsID = Global.processTest.StartStep("Incluir lote de compra", logMsg: $"Incluir lote de compra", paramName: "", paramValue: "");
            try
            {
                gerenciadorDeComprasPO.IncluirLote();
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: "Inclusão do lote de compra com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "erro na inclusão do lote de compra");
            }

            windowName = "Seleção de Produtos";
            lgsID = Global.processTest.StartStep($"Confirmar janela {windowName}", logMsg: $"Tentando confirmar janela {windowName}",
                paramName: "windowName", paramValue: windowName);
            try
            {
                gerenciadorDeComprasPO.ConfirmSelecaoDeProdutosWindow();
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Confirmação janela {windowName} com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"erro na confirmação janela {windowName}");
            }

            //windowName = "Produtos Inativos";
            //lgsID = Global.processTest.StartStep($"Confirmar janela {windowName}", logMsg: $"Tentando confirmar janela {windowName}",
            //    paramName: "windowName", paramValue: windowName);
            //try
            //{
            //    gerenciadorDeComprasPO.ConfirmProdutosInativosWindow();
            //    printFileName = Global.processTest.CaptureWholeScreen();
            //    Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Confirmação janela {windowName} com sucesso");
            //}
            //catch
            //{
            //    printFileName = Global.processTest.CaptureWholeScreen();
            //    Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"erro na confirmação janela {windowName}");
            //}

            //Global.processTest.DoStep("Preencher quantidade de compra por produto e quantidade",
            //    "Preenchimento da quantidade de compra por produto e quantidade com sucesso");

            //lgsID = Global.processTest.StartStep($"Preencher quantidade de compra por produto e quantidade",
            //    logMsg: $"Tentando preencher quantidade de compra por produto e quantidade",
            //    paramName: "qtdProdutos e qtdeCompra", paramValue: $"{qtdProdutos} e {qtdeCompra}");
            //try
            //{
            //    gerenciadorDeComprasPO.FillQtdeCompraPorProduto(qtdProdutos, qtdeCompra);
            //    //Validação total de compra
            //    string qtdeComprValue = gerenciadorDeComprasPO.GetQtdeComprValue();
            //    int qtdeComprasValue = int.Parse(qtdeComprValue);
            //    if (qtdeComprasValue != qtdProdutos * qtdeCompra)
            //    {
            //        Console.Write($"Erro no preenchimento: qtdeComprasValue atual: {qtdeComprasValue}, Total esperado: {qtdProdutos * qtdeCompra}");
            //    }
            //    printFileName = Global.processTest.CaptureWholeScreen();
            //    Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Preenchimento da quantidade de compra por produto e quantidade com sucesso");
            //}
            //catch
            //{
            //    printFileName = Global.processTest.CaptureWholeScreen();
            //    Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName,
            //        logMsg: $"Erro no preenchimento da quantidade de compra por produto e quantidade");
            //}

            //lgsID = Global.processTest.StartStep($"Gerar Pedidos", logMsg: $"Tentando Gerar Pedidos", paramName: "", paramValue: "");
            //try
            //{
            //    gerenciadorDeComprasPO.ClickGerarPedidos();
            //    gerenciadorDeComprasPO.ConfirmPedidosWindow();
            //    printFileName = Global.processTest.CaptureWholeScreen();
            //    Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Pedidos gerados com sucesso");
            //}
            //catch
            //{
            //    printFileName = Global.processTest.CaptureWholeScreen();
            //    Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName,
            //        logMsg: $"erro ao clicar no botão Gera Pedidos");
            //}

            //lgsID = Global.processTest.StartStep($"Confirmar Consulta do lote", logMsg: $"Tentando confirmar Consulta do Lote", paramName: "", paramValue: "");
            //try
            //{
            //    gerenciadorDeComprasPO.ConfirmConsultaLoteCompraWindow();
            //    printFileName = Global.processTest.CaptureWholeScreen();
            //    Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Lote gerado com sucesso");
            //}
            //catch
            //{
            //    printFileName = Global.processTest.CaptureWholeScreen();
            //    Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName,
            //        logMsg: $"erro ao clicar no confirmar Lote");
            //}
            // Teardown function
            PressEnter();
            Global.processTest.EndTest(reportID);
        }
    }
}