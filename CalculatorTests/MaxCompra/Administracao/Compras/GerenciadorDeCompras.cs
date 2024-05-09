using Consinco.Helpers;
using Consinco.MaxCompra.PageObjects.Administracao.Compras;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using static Google.Protobuf.Compiler.CodeGeneratorResponse.Types;

namespace Consinco.MaxCompra.Administracao.Compras
{
    [TestClass]
    public class GerenciadorDeComprasTests : MaxCompraInit
    {
        private static GerenciadorDeComprasPO gerenciadorDeComprasPO = new GerenciadorDeComprasPO(new ElementHandler());
        private string idLote;

        private void DefineSteps(string testName, int qtdLojas = 1)
        {
            switch (testName)
            {
                case "Login":
                    Global.processTest.DoStep("Abrir app", "Abertura do app com sucesso");
                    Global.processTest.DoStep("Login do analista", "Login com sucesso");
                    Global.processTest.DoStep("Tela final", "Tela principal exibida com sucesso");
                    break;
                case "Abrir Gerenciador de Compras":
                    Global.processTest.DoStep("Abrir menu Administração", "Menu Administracao aberto com sucesso");
                    Global.processTest.DoStep("Abrir menu Compras", "Menu Compras aberto com sucesso");
                    Global.processTest.DoStep("Abrir menu Gerenciador de Compras", "Gerenciador de Compras aberto com sucesso");
                    break;
                case "Preencher fornecedor, categoria, abastecimento e checkboxes":
                    Global.processTest.DoStep("Preencher fornecedor", "Fornecedor preenchido com sucesso");
                    Global.processTest.DoStep("Selecionar categoria", "Seleção da categoria com sucesso");
                    Global.processTest.DoStep("Preencher dias abastecimento", "Preenchimento dias abastecimento com sucesso");
                    Global.processTest.DoStep("Habilitar checkbox Sugestão de compra", "Habilitação checkboxes sugestão de compra com sucesso");
                    break;
                case "Adicionar lojas":
                    Global.processTest.DoStep("Adicionar lojas", "Adição de lojas com sucesso");
                    Global.processTest.DoStep("Confirmar janela Seleção de Empresas do Lote", "Confirmação janela Seleção de Empresas do Lote com sucesso");
                    break;
                case "Incluir lote com produtos inativos":
                    Global.processTest.DoStep("Incluir lote de compra", "Inclusão do lote de compra com sucesso");
                    Global.processTest.DoStep("Confirmar janela Filtros para Seleção de Produtos", "Confirmação janela Seleção de Produtos com sucesso");
                    Global.processTest.DoStep("Confirmar janela Produtos Inativos", "Confirmação janela Produtos Inativos com sucesso");
                    break;
                case "Gerar pedidos":
                    Global.processTest.DoStep("Gerar Pedidos", "Pedidos gerados com sucesso");
                    Global.processTest.DoStep("Confirmar janela Consulta Lote de Compra", "Confirmação janela Consulta Lote de Compra com sucesso");
                    break;
                case "CriarLoteDeCompraLojaALoja":
                    DefineSteps("Login");
                    DefineSteps("Abrir Gerenciador de Compras");
                    DefineSteps("Preencher fornecedor, categoria, abastecimento e checkboxes");
                    Global.processTest.DoStep("Incluir lote de compra", "Inclusão do lote de compra com sucesso");
                    Global.processTest.DoStep("Adicionar comprador", "Inclusão de comprador com sucesso");
                    Global.processTest.DoStep("Confirmar janela Filtros para Seleção de Produtos", "Confirmação janela Seleção de Produtos com sucesso");
                    Global.processTest.DoStep("Confirmar janela Tributação", "Confirmação janela Tributação com sucesso");
                    Global.processTest.DoStep("Preencher quantidade de compra dos produtos", "Preenchimento quantidade de compra dos produto com sucesso");
                    DefineSteps("Gerar pedidos");
                    break;
                case "CriarLoteDeCompraIncorporaCD":
                    DefineSteps("Login");
                    DefineSteps("Abrir Gerenciador de Compras");
                    DefineSteps("Preencher fornecedor, categoria, abastecimento e checkboxes");
                    DefineSteps("Adicionar lojas");
                    Global.processTest.DoStep("Habilitar checkbox Incorporar Sugestão CD", "Habilitação do checkbox Incorporar Sugestão com sucesso");
                    DefineSteps("Incluir lote com produtos inativos");
                    Global.processTest.DoStep("Validar duplo click no campo QTD sugerida", "Duplo click no campo QTD sugerida exibe aviso");
                    Global.processTest.DoStep("Edição do campo QtdeCompra no grid de produtos para lotes com Incorpora CD", "Campo QtdeCompra no grid de produtos é editável");
                    Global.processTest.DoStep("Preencher quantidade de compra dos produtos", "Preenchimento quantidade de compra dos produto com sucesso");
                    DefineSteps("Gerar pedidos");
                    break;
                case "CriarLoteDeCompraFLVComprador":
                    DefineSteps("Login");
                    DefineSteps("Abrir Gerenciador de Compras");
                    DefineSteps("Preencher fornecedor, categoria, abastecimento e checkboxes");
                    DefineSteps("Adicionar lojas");
                    Global.processTest.DoStep("Habilitar checkbox Restringe Empresa Loja", "Habilitação checkbox Restringe Empresa Loja com sucesso");
                    DefineSteps("Incluir lote com produtos inativos");
                    Global.processTest.DoStep("Confirmar janela Tributação", "Confirmação janela Tributação com sucesso");
                    Global.processTest.DoStep("Resgatar id do lote", "Resgate do id do lote com sucesso");
                    Global.processTest.DoStep($"Fechar app", "App fechado com sucesso");
                    break;
                case "FinalizarLoteDeCompraFLVComprador":
                    DefineSteps("Login");
                    DefineSteps("Abrir Gerenciador de Compras");
                    Global.processTest.DoStep($"Abrir lote de compras", "Lote aberto com sucesso");
                    Global.processTest.DoStep($"Validar quantidade de compra dos produtos", "Validação da quantidade de compra por produto e quantidade com sucesso");
                    DefineSteps("Gerar pedidos");
                    break;
                case "PreencherLoteDeCompraFLVChefeSessao":
                    DefineSteps("Login");
                    DefineSteps("Abrir Gerenciador de Compras");
                    Global.processTest.DoStep("Abrir lote de compras", "Lote de compras aberto com sucesso");
                    Global.processTest.DoStep("Preencher quantidade de compra dos produtos", "Preenchimento quantidade de compra dos produto com sucesso");
                    Global.processTest.DoStep($"Fechar app", "App fechado com sucesso");
                    break;
                case "CriarLoteDeCompraFLVCompleto":
                    //Comprador | Criar lote
                    DefineSteps("CriarLoteDeCompraFLVComprador");

                    //Chefes de sessão | Preencher lote
                    for (int i = 0; i < qtdLojas; i++)
                    {
                        DefineSteps("PreencherLoteDeCompraFLVChefeSessao");
                    }

                    //Comprador | Finalizar lote
                    DefineSteps("FinalizarLoteDeCompraFLVComprador");
                    break;
                default:
                    throw new Exception($"{testName}'s steps has not been defined.");
            }
        }

        private void GetLoteId()
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep("Resgatar id do lote", logMsg: $"Resgatar id do lote",
                paramName: "Global.app", paramValue: Global.app);
            try
            {
                idLote = gerenciadorDeComprasPO.GetIdLoteDeCompra();
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Resgate do id {idLote} com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "Erro ao tentar resgatar id do lote");
            }

        }

        private void CloseApp()
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep("Fechar app", logMsg: $"Fechar app {Global.app}",
                paramName: "Global.app", paramValue: Global.app);
            try
            {
                CloseApp(Global.app);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: "App fechado com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "Erro ao tentar fechar app");
            }
        }

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

        private void GeneratePedidos(string tipoLote)
        {
            string printFileName;
            int lgsID;

            switch (tipoLote)
            {
                case "loja-a-loja":
                case "cd":
                case "flv":
                    lgsID = Global.processTest.StartStep($"Gerar Pedidos", logMsg: $"Tentando Gerar Pedidos", paramName: "", paramValue: "");
                    try
                    {
                        gerenciadorDeComprasPO.ClickGerarPedidos();
                        if (tipoLote == "cd")
                        {
                            gerenciadorDeComprasPO.ExitWindow("Atenção");
                        }
                        if (tipoLote == "flv")
                        {
                            gerenciadorDeComprasPO.ExitWindow("Atenção");
                            gerenciadorDeComprasPO.ExitWindow("Atenção");
                        }
                        gerenciadorDeComprasPO.ExitWindow("Opções de geração do(s) pedido(s)");
                        elementHandler.FindElementByXPathPartialName("O Pedido de Compra foi gerado com sucesso.");
                        printFileName = Global.processTest.CaptureWholeScreen();
                        Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Pedidos gerados com sucesso");
                    }
                    catch
                    {
                        printFileName = Global.processTest.CaptureWholeScreen();
                        Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName,
                            logMsg: $"Erro ao tentar gerar Pedidos");
                    }
                    break;
                default:
                    throw new Exception("Erro ao tentar gerar pedidos!");
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
                gerenciadorDeComprasPO.FillQtdeCompra(qtdProdutos: qtdProdutos, qtdeCompra: qtdeCompra, tipoLote: tipoLote);
                gerenciadorDeComprasPO.ValidateQtdeComprasValue(qtdProdutos, qtdeCompra, qtdLojas);
                WaitSeconds(3);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg:
                    $"Preenchido QtdeCompra: {qtdeCompra} com {qtdProdutos} produto(s), total validado: {qtdProdutos * qtdeCompra}");
            }
            catch
            {
                string totalComErro = gerenciadorDeComprasPO.GetQtdeComprValue();
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName,
                    logMsg: $"Erro, total esperado: {qtdProdutos * qtdeCompra * qtdLojas}, total atual: {totalComErro}");
            }
        }

        private void ValidateQtdeComprasValue(int qtdProdutos, int qtdeCompra, int qtdLojas, string tipoLote)
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep($"Validar quantidade de compra dos produtos",
                logMsg: $"Tentando validar quantidade de compra dos produtos",
                paramName: "tipoLote, qtdProdutos, qtdeCompra, qtdLojas", paramValue: $"{tipoLote}, {qtdProdutos}, {qtdeCompra}, {qtdLojas}");
            try
            {
                gerenciadorDeComprasPO.ValidateQtdeComprasValue(qtdProdutos, qtdeCompra, qtdLojas);
                WaitSeconds(3);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Validação da quantidade de compra por produto e quantidade com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName,
                    logMsg: $"Erro na validação da quantidade de compra por produto e quantidade");
            }
        }

        private void IncludeLote()
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep("Incluir lote de compra", logMsg: $"Incluir lote de compra", paramName: "", paramValue: "");
            try
            {
                gerenciadorDeComprasPO.IncludeLote();
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
                WaitSeconds(3);
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
                case "Seleção de Empresas do Lote":
                case "Filtros para Seleção de Produtos":
                case "Produtos Inativos":
                case "Tributação":
                case "Atenção":
                case "Opções de geração do(s) pedido(s)":
                case "Consulta Lote de Compra":
                    lgsID = Global.processTest.StartStep($"Confirmar janela {windowName}", logMsg: $"Tentando confirmar janela {windowName}",
                        paramName: "windowName", paramValue: windowName);
                    try
                    {
                        gerenciadorDeComprasPO.ExitWindow(windowName);
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
                    throw new Exception($"Window {windowName} not found.");
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
                WaitSeconds(3);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Lote de compras aberto com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"Erro ao tentar abrir lote de compras: {idLote}");
            }
        }

        private void ValidateDoubleClickOnQtdSugerida()
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep($"Validar duplo click no campo QTD sugerida",
                logMsg: $"Duplo click no campo QTD sugerida para lotes com incorpora cd",
                paramName: "", paramValue: "");
            try
            {
                gerenciadorDeComprasPO.DoubleClickOnQtdSugerida();
                printFileName = Global.processTest.CaptureWholeScreen();
                PressEnter();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Duplo click no campo QTD sugerida exibe aviso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg:
                    $"Erro ao tentar duplo click no campo QTD sugerida para lotes com incorpora cd");
            }
        }

        private void ValidateProductsGridEdit(int qtdeCompra)
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep($"Edição do campo QtdeCompra no grid de produtos para lotes com Incorpora CD",
                logMsg: $"Campo QtdeCompra no grid de produtos é editável",
                paramName: "qtdeCompra", paramValue: qtdeCompra.ToString());
            try
            {
                MaximizeWindow();
                gerenciadorDeComprasPO.FillProductsGrid(qtdeCompra);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Edição do campo QtdeCompra permitido no grid de produtos");
                RestoreWindow();
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg:
                    $"Erro ao tentar editar o campo QtdeCompra, no grid de produtos, para lotes com incorpora cd");
            }
        }

        [TestMethod]
        public void CriarLoteDeCompraLojaALoja()
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
            DefineSteps("CriarLoteDeCompraLojaALoja");

            Login(worksheet, rowNumber);
            OpenGerenciadorDeCompras();
            FillFornecedor(codFornecedor);
            SelectCategoria(categoria);
            FillAbastecimentoDias(diasAbastecimento);
            EnableCheckbox("Sugestão de compra");
            IncludeLote();
            AddComprador(comprador);
            ConfirmWindow("Filtros para Seleção de Produtos");
            ConfirmWindow("Tributação");
            FillProdutos(qtdProdutos, qtdeCompra, qtdLojas, tipoLote);
            GeneratePedidos(tipoLote);
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
            int qtdLojas = int.Parse(excelReader.ReadCellValueToString(worksheet, "qtdLojas", rowNumber));
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
            DefineSteps("CriarLoteDeCompraIncorporaCD");

            Login(worksheet, rowNumber);
            OpenGerenciadorDeCompras();
            FillFornecedor(codFornecedor);
            SelectCategoria(categoria);
            FillAbastecimentoDias(diasAbastecimento);
            EnableCheckbox("Sugestão de compra");
            AddLojas(lojas, divisao);
            ConfirmWindow("Seleção de Empresas do Lote");
            EnableCheckbox(feature: "Incorporar Sugestão CD", paramName: "cdNome", paramValue: cdNome);
            IncludeLote();
            ConfirmWindow("Filtros para Seleção de Produtos");
            ConfirmWindow("Produtos Inativos");
            ValidateDoubleClickOnQtdSugerida();
            ValidateProductsGridEdit(qtdeCompra);
            FillProdutos(qtdProdutos, qtdeCompra, qtdLojas, tipoLote);
            GeneratePedidos(tipoLote);
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
            DefineSteps("CriarLoteDeCompraFLVComprador");

            Login(worksheet, rowNumber);
            OpenGerenciadorDeCompras();
            FillFornecedor(codFornecedor);
            SelectCategoria(categoria);
            FillAbastecimentoDias(diasAbastecimento);
            EnableCheckbox("Sugestão de compra");
            AddLojas(lojas, divisao);
            ConfirmWindow("Seleção de Empresas do Lote");
            EnableCheckbox("Restringe Empresa Loja");
            IncludeLote();
            ConfirmWindow("Filtros para Seleção de Produtos");
            ConfirmWindow("Produtos Inativos");
            ConfirmWindow("Tributação");
            GetLoteId();
            CloseApp();

            // Teardown function
            Global.processTest.EndTest(reportID);
        }

        [TestMethod]
        public void FinalizarLoteDeCompraFLVComprador()
        {
            // Global Variables
            int rowNumber = 8;
            string worksheetName = "GerenciadorDeCompras";
            ExcelWorksheet worksheet = excelReader.OpenWorksheet(excelFilePath, worksheetName);

            // Test Variables
            string tipoLote = excelReader.ReadCellValueToString(worksheet, "tipoLote", rowNumber);
            //string idLote = excelReader.ReadCellValueToString(worksheet, "idLote", rowNumber);
            int qtdProdutos = int.Parse(excelReader.ReadCellValueToString(worksheet, "qtdProdutos", rowNumber));
            int qtdeCompra = int.Parse(excelReader.ReadCellValueToString(worksheet, "qtdeCompra", rowNumber));
            int qtdLojas = int.Parse(excelReader.ReadCellValueToString(worksheet, "qtdLojas", rowNumber));

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
            DefineSteps("FinalizarLoteDeCompraFLVComprador");

            Login(worksheet, rowNumber);
            OpenGerenciadorDeCompras();
            OpenLote(idLote);
            ValidateQtdeComprasValue(qtdProdutos: qtdProdutos, qtdeCompra: qtdeCompra, tipoLote: tipoLote, qtdLojas: qtdLojas);
            GeneratePedidos(tipoLote);
            ConfirmWindow("Consulta Lote de Compra");

            // Teardown function
            Global.processTest.EndTest(reportID);
        }

        [TestMethod]
        public void PreencherLoteDeCompraFLVChefeSessao()
        {
            // Global Variables
            List<int> rowNumbers = [5, 6, 7];
            for (int i = 0; i < rowNumbers.Count; i++)
            {
                int rowNumber = rowNumbers[i];
                string worksheetName = "GerenciadorDeCompras";
                ExcelWorksheet worksheet = excelReader.OpenWorksheet(excelFilePath, worksheetName);

                // Test Variables
                List<string> lojas = excelReader.ReadCellValueToList(worksheet, "lojas", rowNumber);
                int qtdProdutos = int.Parse(excelReader.ReadCellValueToString(worksheet, "qtdProdutos", rowNumber));
                int qtdeCompra = int.Parse(excelReader.ReadCellValueToString(worksheet, "qtdeCompra", rowNumber));
                int qtdLojas = lojas.Count;
                string tipoLote = excelReader.ReadCellValueToString(worksheet, "tipoLote", rowNumber);
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
                DefineSteps("PreencherLoteDeCompraFLVChefeSessao", rowNumbers.Count);

                Login(worksheet, rowNumber);
                OpenGerenciadorDeCompras();
                OpenLote(idLote);
                FillProdutos(qtdProdutos, qtdeCompra, qtdLojas, tipoLote);
                CloseApp();

                // Teardown function
                Global.processTest.EndTest(reportID);
            }
        }

        [TestMethod]
        public void CriarLoteDeCompraFLVCompleto()
        {
            CriarLoteDeCompraFLVComprador();
            PreencherLoteDeCompraFLVChefeSessao();
            FinalizarLoteDeCompraFLVComprador();
        }
    }

}