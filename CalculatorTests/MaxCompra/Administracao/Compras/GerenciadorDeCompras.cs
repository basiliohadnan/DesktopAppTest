using Consinco.Helpers;
using Consinco.MaxCompra.PageObjects.Administracao.Compras;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace Consinco.MaxCompra.Administracao.Compras
{
    [TestClass]
    public class GerenciadorDeComprasTests : MaxCompraInit
    {
        private static GerenciadorDeComprasPO gerenciadorDeComprasPO = new GerenciadorDeComprasPO(new ElementHandler());


        private void OpenGerenciadorDeCompras()
        {
            string printFileName;

            string menuName = "Administração";
            int lgsID = Global.processTest.StartStep($"Abrir menu {menuName}", logMsg: $"Abrir menu {menuName}",
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
        }

        private void FillFornecedor(string codFornecedor)
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep("Preencher fornecedor", logMsg: $"Preencher fornecedor {codFornecedor}",
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
        }

        private void AddLojas(List<string> lojas, string divisao)
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep("Adicionar lojas", logMsg: $"Adicionar {lojas.Count} lojas", paramName: "lojas",
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
        }

        private void GenerateInvoices()
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep($"Gerar Pedidos", logMsg: $"Tentando Gerar Pedidos", paramName: "", paramValue: "");
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
        }

        private void FillProdutos(int qtdProdutos, int qtdeCompra, int qtdLojas, string tipoLote)
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep($"Preencher quantidade de compra dos produtos",
                logMsg: $"Tentando preencher quantidade de compra dos produtos",
                paramName: "tipoLote, qtdProdutos, qtdeCompra, qtdLojas", paramValue: $"{tipoLote}, {qtdProdutos}, {qtdeCompra}, {qtdLojas}");
            try
            {
                gerenciadorDeComprasPO.FillQtdeCompra(qtdLojas: qtdLojas, qtdProdutos: qtdProdutos, qtdeCompra: qtdeCompra, tipoLote: tipoLote);
                gerenciadorDeComprasPO.ValidateQtdeComprasValue(qtdProdutos, qtdeCompra);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Preenchimento da quantidade de compra por produto e quantidade com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName,
                    logMsg: $"Erro no preenchimento da quantidade de compra por produto e quantidade");
            }
        }

        private void IncludeLote()
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep("Incluir lote de compra", logMsg: $"Incluir lote de compra", paramName: "", paramValue: "");
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
        }

        private void EnableCheckboxesSugestaoCompras()
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep("Habilitar checkboxes Sugestão de compra", logMsg: $"Habilitar checkboxes sugestão de compras",
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
        }

        private void EnableIncorporarSugestaoCompras(string cdNome)
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep("Habilitar Incorporar Sugestão", logMsg: $"Habilitar Incorporar Sugestão para o CD {cdNome}",
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
        }

        private void FillAbastecimentoDias(string diasAbastecimento)
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep("Preencher dias abastecimento", logMsg: $"Preencher dias abastecimento com {diasAbastecimento}",
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
        }

        private void SelectCategoria(string categoria)
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep("Selecionar categoria", logMsg: $"Selecionar categoria {categoria}",
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
        }

        private void ConfirmWindow(string windowName)
        {
            string printFileName;
            int lgsID;

            switch (windowName)
            {
                case "Seleção de Produtos":
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
                    break;

                case "Produtos Inativos":
                    lgsID = Global.processTest.StartStep($"Confirmar janela {windowName}", logMsg: $"Tentando confirmar janela {windowName}",
                        paramName: "windowName", paramValue: windowName);
                    try
                    {
                        gerenciadorDeComprasPO.ConfirmWindow(windowName);
                        printFileName = Global.processTest.CaptureWholeScreen();
                        Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Confirmação janela {windowName} com sucesso");
                    }
                    catch
                    {
                        printFileName = Global.processTest.CaptureWholeScreen();
                        Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"erro na confirmação janela {windowName}");
                    }
                    break;

                case "Seleção de Empresas do Lote":
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
                    break;
                case "Consulta Lote de Compra":
                    lgsID = Global.processTest.StartStep($"Confirmar janela {windowName}", logMsg: $"Tentando confirmar janela {windowName}",
                        paramName: "windowName", paramValue: windowName);
                    try
                    {
                        gerenciadorDeComprasPO.ConfirmConsultaLoteCompraWindow();
                        printFileName = Global.processTest.CaptureWholeScreen();
                        Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Confirmação janela {windowName} com sucesso");
                    }
                    catch
                    {
                        printFileName = Global.processTest.CaptureWholeScreen();
                        Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"erro na confirmação janela {windowName}");
                    }
                    break;
                case "Tributação":
                    lgsID = Global.processTest.StartStep($"Confirmar janela {windowName}", logMsg: $"Tentando confirmar janela {windowName}",
                        paramName: "windowName", paramValue: windowName);
                    try
                    {
                        gerenciadorDeComprasPO.ConfirmTributacaoWindow();
                        printFileName = Global.processTest.CaptureWholeScreen();
                        Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Confirmação janela {windowName} com sucesso");
                    }
                    catch
                    {
                        printFileName = Global.processTest.CaptureWholeScreen();
                        Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"erro na confirmação janela {windowName}");
                    }
                    break;
                default:
                    // Handle any other cases here
                    break;
            }
        }

        private void AddComprador(string comprador)
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep("Adicionar comprador", logMsg: $"Adicionar comprador {comprador}", paramName: "comprador",
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
        }

        [TestMethod]
        public void CriarLoteDeCompraLojaALojaUmaLoja()
        {
            // Global Variables
            int rowNumber = 2;
            string worksheetName = "GerenciadorDeCompras";
            ExcelWorksheet worksheet = excelReader.OpenWorksheet(excelFilePath, worksheetName);

            // Test Variables
            string codFornecedor = excelReader.ReadCellValueToString(worksheet, "codFornecedor", rowNumber);
            string categoria = excelReader.ReadCellValueToString(worksheet, "categoria", rowNumber);
            string diasAbastecimento = excelReader.ReadCellValueToString(worksheet, "diasAbastecimento", rowNumber);
            string comprador = excelReader.ReadCellValueToString(worksheet, "comprador", rowNumber);
            int qtdLojas = int.Parse(excelReader.ReadCellValueToString(worksheet, "qtdLojas", rowNumber));
            int qtdProdutos = int.Parse(excelReader.ReadCellValueToString(worksheet, "qtdProdutos", rowNumber));
            int qtdeCompra = int.Parse(excelReader.ReadCellValueToString(worksheet, "qtdeCompra", rowNumber));
            string tipoLote = excelReader.ReadCellValueToString(worksheet, "tipoLote", rowNumber);

            int reportID = int.Parse(excelReader.ReadCellValueToString(worksheet, "reportID", rowNumber));
            string scenarioName = excelReader.ReadCellValueToString(worksheet, "scenarioName", rowNumber);
            string database = excelReader.ReadCellValueToString(worksheet, "database", rowNumber);
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
            Global.processTest.DoStep("Abrir menu Administração", "Menu Administracao aberto com sucesso");
            Global.processTest.DoStep("Abrir menu Compras", "Menu Compras aberto com sucesso");
            Global.processTest.DoStep("Abrir menu Gerenciador de Compras", "Gerenciador de Compras aberto com sucesso");
            Global.processTest.DoStep("Preencher fornecedor", "Fornecedor preenchido com sucesso");
            Global.processTest.DoStep("Selecionar categoria", "Seleção da categoria com sucesso");
            Global.processTest.DoStep("Preencher dias abastecimento", "Preenchimento dias abastecimento com sucesso");
            Global.processTest.DoStep("Habilitar checkboxes Sugestão de compra", "Habilitação checkboxes sugestão de compra com sucesso");
            Global.processTest.DoStep("Incluir lote de compra", "Inclusão do lote de compra com sucesso");
            Global.processTest.DoStep("Adicionar comprador", "Inclusão de comprador com sucesso");
            Global.processTest.DoStep("Confirmar janela Seleção de Produtos", "Confirmação janela Seleção de Produtos com sucesso");
            Global.processTest.DoStep("Preencher quantidade de compra dos produtos", "Preenchimento quantidade de compra dos produto com sucesso");
            Global.processTest.DoStep("Gerar Pedidos", "Pedidos gerados com sucesso");
            Global.processTest.DoStep("Confirmar janela Consulta Lote de Compra", "Confirmação janela Consulta Lote de Compra com sucesso");

            Login(worksheet, rowNumber);
            OpenGerenciadorDeCompras();
            FillFornecedor(codFornecedor);
            SelectCategoria(categoria);
            FillAbastecimentoDias(diasAbastecimento);
            EnableCheckboxesSugestaoCompras();
            IncludeLote();
            AddComprador(comprador);
            ConfirmWindow("Seleção de Produtos");
            FillProdutos(qtdProdutos, qtdeCompra, qtdLojas, tipoLote);
            GenerateInvoices();
            ConfirmWindow("Consulta Lote de Compra");

            //Teardown function
            Global.processTest.EndTest(reportID);
        }

        [TestMethod]
        public void CriarLoteDeCompraIncorporaCD()
        {
            // Global Variables
            int rowNumber = 3;
            string worksheetName = "GerenciadorDeCompras";
            ExcelWorksheet worksheet = excelReader.OpenWorksheet(excelFilePath, worksheetName);

            // Test Variables
            string codFornecedor = excelReader.ReadCellValueToString(worksheet, "codFornecedor", rowNumber);
            List<string> lojas = excelReader.ReadCellValueToList(worksheet, "lojas", rowNumber);
            string divisao = excelReader.ReadCellValueToString(worksheet, "divisao", rowNumber);
            string categoria = excelReader.ReadCellValueToString(worksheet, "categoria", rowNumber);
            string diasAbastecimento = excelReader.ReadCellValueToString(worksheet, "diasAbastecimento", rowNumber);
            string cdNome = excelReader.ReadCellValueToString(worksheet, "cdNome", rowNumber);
            int qtdProdutos = int.Parse(excelReader.ReadCellValueToString(worksheet, "qtdProdutos", rowNumber));
            int qtdeCompra = int.Parse(excelReader.ReadCellValueToString(worksheet, "qtdeCompra", rowNumber));
            int qtdLojas = lojas.Count;
            string tipoLote = excelReader.ReadCellValueToString(worksheet, "tipoLote", rowNumber);

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
            Global.processTest.DoStep("Abrir menu Administração", "Menu Administracao aberto com sucesso");
            Global.processTest.DoStep("Abrir menu Compras", "Menu Compras aberto com sucesso");
            Global.processTest.DoStep("Abrir menu Gerenciador de Compras", "Gerenciador de Compras aberto com sucesso");
            Global.processTest.DoStep("Preencher fornecedor", "Fornecedor preenchido com sucesso");
            Global.processTest.DoStep("Adicionar lojas", "Adição de lojas com sucesso");
            Global.processTest.DoStep("Confirmar janela Seleção de Empresas do Lote", "Confirmação janela Seleção de Empresas do Lote com sucesso");
            Global.processTest.DoStep("Selecionar categoria", "Seleção da categoria com sucesso");
            Global.processTest.DoStep("Preencher dias abastecimento", "Preenchimento dias abastecimento com sucesso");
            Global.processTest.DoStep("Habilitar Incorporar Sugestão", "Habilitação do Incorporar Sugestão com sucesso");
            Global.processTest.DoStep("Habilitar checkboxes Sugestão de compra", "Habilitação checkboxes sugestão de compra com sucesso");
            Global.processTest.DoStep("Incluir lote de compra", "Inclusão do lote de compra com sucesso");
            Global.processTest.DoStep("Confirmar janela Seleção de Produtos", "Confirmação janela Seleção de Produtos com sucesso");
            Global.processTest.DoStep("Confirmar janela Produtos Inativos", "Confirmação janela Produtos Inativos com sucesso");
            Global.processTest.DoStep("Preencher quantidade de compra dos produtos", "Preenchimento quantidade de compra dos produto com sucesso");
            Global.processTest.DoStep("Gerar Pedidos", "Pedidos gerados com sucesso");
            Global.processTest.DoStep("Confirmar janela Consulta Lote de Compra", "Confirmação janela Consulta Lote de Compra com sucesso");

            Login(worksheet, rowNumber);
            OpenGerenciadorDeCompras();
            FillFornecedor(codFornecedor);
            AddLojas(lojas, divisao);
            ConfirmWindow("Seleção de Empresas do Lote");
            SelectCategoria(categoria);
            FillAbastecimentoDias(diasAbastecimento);
            EnableIncorporarSugestaoCompras(cdNome);
            EnableCheckboxesSugestaoCompras();
            IncludeLote();
            ConfirmWindow("Seleção de Produtos");
            ConfirmWindow("Produtos Inativos");
            FillProdutos(qtdProdutos, qtdeCompra, qtdLojas, tipoLote);
            GenerateInvoices();
            ConfirmWindow("Consulta Lote de Compra");

            // Teardown function
            PressEnter();
            Global.processTest.EndTest(reportID);
        }

        [TestMethod]
        public void CriarLoteDeCompraFLV()
        {
            // Global Variables
            int rowNumber = 2;
            string worksheetName = "GerenciadorDeCompras";
            ExcelWorksheet worksheet = excelReader.OpenWorksheet(excelFilePath, worksheetName);

            // Test Variables
            int lgsID;
            string printFileName;
            int reportID = int.Parse(excelReader.ReadCellValueToString(worksheet, "reportID", rowNumber));
            string scenarioName = excelReader.ReadCellValueToString(worksheet, "scenarioName", rowNumber);
            string database = excelReader.ReadCellValueToString(worksheet, "database", rowNumber);
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

            if (database == "ASSAIHOMOL")
            {
                Global.processTest.DoStep("Confirmar janela Tributação", "Confirmação janela Tributação com sucesso");
            }

            Global.processTest.DoStep("Preencher quantidade de compra por loja e produto",
                "Preenchimento quantidade de compra por loja e produto com sucesso");
            Global.processTest.DoStep("Gerar Pedidos", "Pedidos gerados com sucesso");
            Global.processTest.DoStep("Confirmar Consulta do lote", "Lote gerado com sucesso");

            Login(worksheet, rowNumber);

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

            string codFornecedor = excelReader.ReadCellValueToString(worksheet, "codFornecedor", rowNumber);
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

            string categoria = excelReader.ReadCellValueToString(worksheet, "categoria", rowNumber);
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

            string diasAbastecimento = excelReader.ReadCellValueToString(worksheet, "diasAbastecimento", rowNumber);
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

            string comprador = excelReader.ReadCellValueToString(worksheet, "comprador", rowNumber);
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

            string windowName = excelReader.ReadCellValueToString(worksheet, "windowName", rowNumber);
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

            if (database == "ASSAIHOMOL")
            {
                windowName = "Tributação";
                Global.processTest.DoStep("Confirmar janela Tributação", "Confirmação janela Tributação com sucesso");
                lgsID = Global.processTest.StartStep($"Confirmar janela {windowName}", logMsg: $"Tentando confirmar janela {windowName}",
                    paramName: "windowName", paramValue: windowName);
                try
                {
                    gerenciadorDeComprasPO.ConfirmTributacaoWindow();
                    printFileName = Global.processTest.CaptureWholeScreen();
                    Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Confirmação janela {windowName} com sucesso");
                }
                catch
                {
                    printFileName = Global.processTest.CaptureWholeScreen();
                    Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"erro na confirmação janela {windowName}");
                }
            }


            int qtdLojas = int.Parse(excelReader.ReadCellValueToString(worksheet, "qtdLojas", rowNumber));
            int qtdProdutos = int.Parse(excelReader.ReadCellValueToString(worksheet, "qtdProdutos", rowNumber));
            int qtdeCompra = int.Parse(excelReader.ReadCellValueToString(worksheet, "qtdeCompra", rowNumber));
            string tipoLote = "loja-a-loja";
            lgsID = Global.processTest.StartStep($"Preencher quantidade de compra por loja e produto",
                logMsg: $"Preencher quantidade de compra por loja e produto",
                paramName: "qtdLojas, qtdProdutos, qtdeCompra, tipoLote", paramValue: $"{qtdLojas}, {qtdProdutos}, {qtdeCompra}, {tipoLote}");
            try
            {
                gerenciadorDeComprasPO.FillQtdeCompra(qtdLojas: qtdLojas, qtdProdutos: qtdProdutos, qtdeCompra: qtdeCompra, tipoLote: tipoLote);
                gerenciadorDeComprasPO.ValidateQtdeComprasValue(qtdProdutos, qtdeCompra);
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
            Global.processTest.EndTest(reportID);
        }
    }
}