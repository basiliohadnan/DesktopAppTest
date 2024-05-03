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
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"Erro na abertura do menu {menuName}");
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
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"Erro na abertura do menu {menuName}");
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
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"Erro na abertura do menu {menuName}");
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
                WinAppDriver.WaitSeconds(2);
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

        private void GenerateLoteDeCompra()
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
                    logMsg: $"Erro ao clicar no botão Gera Pedidos");
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
                WaitSeconds(2);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: "Inclusão do lote de compra com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "erro na inclusão do lote de compra");
            }
        }

        private void EnableCheckbox(string feature, string paramName = "", string paramValue = "")
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep($"Habilitar checkbox {feature}", logMsg: $"Habilitar {feature}",
                            paramName: paramName, paramValue: paramValue);
            try
            {
                switch (feature)
                {


                    case "Sugestão de compra":
                        gerenciadorDeComprasPO.EnableCheckbox("Sugestão Compras");
                        break;
                    case "Incorporar Sugestão CD":
                        gerenciadorDeComprasPO.EnableCheckbox(feature);
                        gerenciadorDeComprasPO.SetCD(paramValue);
                        break;
                    case "Restringe Empresa Loja":
                        gerenciadorDeComprasPO.EnableCheckbox(feature);
                        break;
                    default:
                        throw new ArgumentException($"Unsupported feature: {feature}");
                }

                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Habilitação {feature} com sucesso");
            }
            catch (Exception ex)
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName,
                     logMsg: $"erro na habilitação do {feature}: {ex.Message}");
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
                        Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"Erro na confirmação janela {windowName}");
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
                        Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"Erro na confirmação janela {windowName}");
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
                        Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"Erro na confirmação janela {windowName}");
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
                        Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"Erro na confirmação janela {windowName}");
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
                        Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"Erro na confirmação janela {windowName}");
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

        private void OpenLote(string idLote)
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep($"Abrir lote de compras", logMsg: $"Abrir lote de compras",
                paramName: "idLote", paramValue: idLote);
            try
            {
                gerenciadorDeComprasPO.OpenLote(idLote);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Lote de compras aberto com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"Erro ao tentar abrir lote de compras: {idLote}");
            }
        }

        private string GetIdLote()
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep($"Resgatar id do lote criado", logMsg: $"Resgatar id do lote criado",
                paramName: "", paramValue: "");
            try
            {
                string idLote = gerenciadorDeComprasPO.GetIdLoteDeCompra();
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Resgate do id do lote com sucesso");
                return idLote;
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"Erro ao tentar resgatar id do lote");
                return null;
            }
        }

        private void UpdateLoteDeCompra()
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep($"Atualizar Pedidos", logMsg: $"Atualização dos Pedidos com sucesso",
                paramName: "", paramValue: "");
            try
            {
                gerenciadorDeComprasPO.UpdateLoteDeCompra();
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Atualização dos Pedidos com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"Erro ao tentar atualizar pedidos");
            }
        }

        private void DefineSteps(string testName)
        {
            switch (testName)
            {
                case "Login":
                    Global.processTest.DoStep("Abrir app", "Abertura do app com sucesso");
                    Global.processTest.DoStep("Login do analista", "Login com sucesso");
                    Global.processTest.DoStep("Tela final", "Tela principal exibida com sucesso");
                    break;
                case "CriarLoteDeCompraLojaALojaUmaLoja":
                    DefineSteps("Login");
                    Global.processTest.DoStep("Abrir menu Administração", "Menu Administracao aberto com sucesso");
                    Global.processTest.DoStep("Abrir menu Compras", "Menu Compras aberto com sucesso");
                    Global.processTest.DoStep("Abrir menu Gerenciador de Compras", "Gerenciador de Compras aberto com sucesso");
                    Global.processTest.DoStep("Preencher fornecedor", "Fornecedor preenchido com sucesso");
                    Global.processTest.DoStep("Selecionar categoria", "Seleção da categoria com sucesso");
                    Global.processTest.DoStep("Preencher dias abastecimento", "Preenchimento dias abastecimento com sucesso");
                    Global.processTest.DoStep("Habilitar checkbox Sugestão de compra", "Habilitação checkboxes sugestão de compra com sucesso");
                    Global.processTest.DoStep("Incluir lote de compra", "Inclusão do lote de compra com sucesso");
                    Global.processTest.DoStep("Adicionar comprador", "Inclusão de comprador com sucesso");
                    Global.processTest.DoStep("Confirmar janela Seleção de Produtos", "Confirmação janela Seleção de Produtos com sucesso");
                    Global.processTest.DoStep("Confirmar janela Tributação", "Confirmação janela Tributação com sucesso");
                    Global.processTest.DoStep("Preencher quantidade de compra dos produtos", "Preenchimento quantidade de compra dos produto com sucesso");
                    Global.processTest.DoStep("Gerar Pedidos", "Pedidos gerados com sucesso");
                    Global.processTest.DoStep("Confirmar janela Consulta Lote de Compra", "Confirmação janela Consulta Lote de Compra com sucesso");
                    break;
                default:
                    throw new Exception($"{testName}'s steps has not been defined.");
            }
        }


        [TestMethod] // Quebrado
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
            DefineSteps("CriarLoteDeCompraLojaALojaUmaLoja");

            Login(worksheet, rowNumber);
            OpenGerenciadorDeCompras();
            FillFornecedor(codFornecedor);
            SelectCategoria(categoria);
            FillAbastecimentoDias(diasAbastecimento);
            EnableCheckbox("Sugestão de compra");
            IncludeLote();
            AddComprador(comprador);
            ConfirmWindow("Seleção de Produtos");
            ConfirmWindow("Tributação");
            FillProdutos(qtdProdutos, qtdeCompra, qtdLojas, tipoLote); //verificar metodo
            GenerateLoteDeCompra();
            ConfirmWindow("Consulta Lote de Compra");

            //Teardown function
            Global.processTest.EndTest(reportID);
        }



        [TestMethod] // Impedimento
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
            Global.processTest.DoStep("Habilitar Incorporar Sugestão CD", "Habilitação do Incorporar Sugestão com sucesso");
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
            EnableCheckbox("Incorporar Sugestão CD", "cdNome", cdNome);
            EnableCheckbox("Sugestão de compra");
            IncludeLote();
            ConfirmWindow("Seleção de Produtos");
            ConfirmWindow("Produtos Inativos");
            FillProdutos(qtdProdutos, qtdeCompra, qtdLojas, tipoLote);
            GenerateLoteDeCompra();
            ConfirmWindow("Consulta Lote de Compra");

            // Teardown function
            Global.processTest.EndTest(reportID);
        }

        [TestMethod]
        public void CriarLoteDeCompraFLVComprador()
        {
            // Global Variables
            int rowNumber = 4;
            string worksheetName = "GerenciadorDeCompras";
            ExcelWorksheet worksheet = excelReader.OpenWorksheet(excelFilePath, worksheetName);

            // Test Variables
            string codFornecedor = excelReader.ReadCellValueToString(worksheet, "codFornecedor", rowNumber);
            List<string> lojas = excelReader.ReadCellValueToList(worksheet, "lojas", rowNumber);
            string divisao = excelReader.ReadCellValueToString(worksheet, "divisao", rowNumber);
            string categoria = excelReader.ReadCellValueToString(worksheet, "categoria", rowNumber);
            string diasAbastecimento = excelReader.ReadCellValueToString(worksheet, "diasAbastecimento", rowNumber);

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
            Global.processTest.DoStep("Habilitar checkboxes Sugestão de compra", "Habilitação checkboxes sugestão de compra com sucesso");
            Global.processTest.DoStep("Habilitar checkbox Restringe Empresa Loja", "Habilitação checkbox Restringe Empresa Loja com sucesso");
            Global.processTest.DoStep("Incluir lote de compra", "Inclusão do lote de compra com sucesso");
            Global.processTest.DoStep("Confirmar janela Seleção de Produtos", "Confirmação janela Seleção de Produtos com sucesso");
            Global.processTest.DoStep("Confirmar janela Produtos Inativos", "Confirmação janela Produtos Inativos com sucesso");
            Global.processTest.DoStep("Confirmar janela Tributação", "Confirmação janela Tributação com sucesso");
            Global.processTest.DoStep("Resgatar id do lote criado", "Resgate do id do lote com sucesso");
            Global.processTest.DoStep("Atualizar Pedidos", "Atualização dos Pedidos com sucesso");
            Global.processTest.DoStep("Gerar Pedidos", "Pedidos gerados com sucesso");
            Global.processTest.DoStep("Confirmar janela Consulta Lote de Compra", "Confirmação janela Consulta Lote de Compra com sucesso");

            Login(worksheet, rowNumber);
            OpenGerenciadorDeCompras();
            FillFornecedor(codFornecedor);
            AddLojas(lojas, divisao);
            ConfirmWindow("Seleção de Empresas do Lote");
            SelectCategoria(categoria);
            FillAbastecimentoDias(diasAbastecimento);
            EnableCheckbox("Habilitar checkboxes Sugestão de compra");
            EnableCheckbox("Habilitar checkbox Restringe Empresa Loja");
            IncludeLote();
            ConfirmWindow("Seleção de Produtos");
            ConfirmWindow("Produtos Inativos");
            ConfirmWindow("Tributação");
            string idLote = GetIdLote();

            // validar winappdriver pegar sessao atual
            CriarLoteDeCompraFLVChefeSessao(5, idLote);
            CriarLoteDeCompraFLVChefeSessao(6, idLote);

            UpdateLoteDeCompra();
            GenerateLoteDeCompra();
            ConfirmWindow("Consulta Lote de Compra");

            // Teardown function
            Global.processTest.EndTest(reportID);
        }

        [TestMethod]
        public void CriarLoteDeCompraFLVChefeSessao(int rowNumber, string idLote)
        {
            // Global Variables
            //int rowNumber = 5;
            string worksheetName = "GerenciadorDeCompras";
            ExcelWorksheet worksheet = excelReader.OpenWorksheet(excelFilePath, worksheetName);

            // Test Variables
            // abstrair
            List<string> lojas = excelReader.ReadCellValueToList(worksheet, "lojas", rowNumber);
            int qtdProdutos = int.Parse(excelReader.ReadCellValueToString(worksheet, "qtdProdutos", rowNumber));
            int qtdeCompra = int.Parse(excelReader.ReadCellValueToString(worksheet, "qtdeCompra", rowNumber));
            int qtdLojas = lojas.Count;
            string tipoLote = excelReader.ReadCellValueToString(worksheet, "tipoLote", rowNumber);
            // tem que vir do lote gerado pelo Comprador (teste anterior)
            //string idLote = excelReader.ReadCellValueToString(worksheet, "idLote", rowNumber);

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
            Global.processTest.DoStep("Abrir lote de compras", "Lote de compras aberto com sucesso");
            Global.processTest.DoStep("Preencher quantidade de compra dos produtos", "Preenchimento quantidade de compra dos produto com sucesso");

            Login(worksheet, rowNumber);
            OpenGerenciadorDeCompras();
            OpenLote(idLote);
            FillProdutos(qtdProdutos, qtdeCompra, qtdLojas, tipoLote);

            // Teardown function
            Global.processTest.EndTest(reportID);
        }
    }
}