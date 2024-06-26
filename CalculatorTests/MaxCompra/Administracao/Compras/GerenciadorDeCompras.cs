﻿using Consinco.Helpers;
using Consinco.MaxCompra.PageObjects.Administracao.Compras;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Starline;

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
                case "Criar capa lote":
                    base.DefineSteps("RealizarLogin");
                    DefineSteps("Abrir Gerenciador de Compras");
                    DefineSteps("Preencher fornecedor, categoria, abastecimento e checkboxes");
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
                    DefineSteps("Criar capa lote");
                    Global.processTest.DoStep("Incluir lote de compra", "Inclusão do lote de compra com sucesso");
                    Global.processTest.DoStep("Adicionar comprador", "Inclusão de comprador com sucesso");
                    Global.processTest.DoStep("Confirmar janela Filtros para Seleção de Produtos", "Confirmação janela Seleção de Produtos com sucesso");
                    Global.processTest.DoStep("Confirmar janela Tributação", "Confirmação janela Tributação com sucesso");
                    Global.processTest.DoStep("Preencher quantidade de compra dos produtos", "Preenchimento quantidade de compra dos produto com sucesso");
                    Global.processTest.DoStep($"Validar quantidade de compra dos produtos", "Validação da quantidade de compra por produto e quantidade com sucesso");
                    DefineSteps("Gerar pedidos");
                    break;
                case "CriarLoteDeCompraIncorporaCD":
                    DefineSteps("Criar capa lote");
                    DefineSteps("Adicionar lojas");
                    Global.processTest.DoStep("Habilitar checkbox Incorporar Sugestão CD", "Habilitação do checkbox Incorporar Sugestão com sucesso");
                    DefineSteps("Incluir lote com produtos inativos");
                    Global.processTest.DoStep("Validar duplo click no campo QTD sugerida", "Duplo click no campo QTD sugerida exibe aviso");
                    Global.processTest.DoStep("Edição do campo QtdeCompra no grid de produtos para lotes com Incorpora CD", "Campo QtdeCompra no grid de produtos é editável");
                    Global.processTest.DoStep("Preencher quantidade de compra dos produtos", "Preenchimento quantidade de compra dos produto com sucesso");
                    Global.processTest.DoStep($"Validar quantidade de compra dos produtos", "Validação da quantidade de compra por produto e quantidade com sucesso");
                    DefineSteps("Gerar pedidos");
                    break;
                case "CriarLoteDeCompraLojaALojaBonificacao":
                    DefineSteps("Criar capa lote");
                    Global.processTest.DoStep("Preencher Limite Recebimento", "Campo Limite Recebimento preenchido");
                    Global.processTest.DoStep("Alterar tipo do pedido", "Alteração do tipo do pedido com sucesso");
                    Global.processTest.DoStep("Alterar tipo do acordo", "Alteração do tipo do acordo com sucesso");
                    Global.processTest.DoStep("Incluir lote de compra", "Inclusão do lote de compra com sucesso");
                    Global.processTest.DoStep("Confirmar janela Filtros para Seleção de Produtos", "Confirmação janela Seleção de Produtos com sucesso");
                    Global.processTest.DoStep("Confirmar janela Tributação", "Confirmação janela Tributação com sucesso");
                    Global.processTest.DoStep("Preencher quantidade de compra dos produtos", "Preenchimento quantidade de compra dos produto com sucesso");
                    Global.processTest.DoStep($"Validar quantidade de compra dos produtos", "Validação da quantidade de compra por produto e quantidade com sucesso");
                    Global.processTest.DoStep("Gerar Pedidos", "Pedidos gerados com sucesso");
                    Global.processTest.DoStep("Confirmar janela Manutenção de Acordos Promocionais", "Confirmação janela Manutenção de Acordos Promocionais com sucesso");
                    break;
                case "CriarLoteDeCompraIncorporaCDBonificacao":
                    DefineSteps("Criar capa lote");
                    DefineSteps("Adicionar lojas");
                    Global.processTest.DoStep("Habilitar checkbox Incorporar Sugestão CD", "Habilitação do checkbox Incorporar Sugestão com sucesso");
                    Global.processTest.DoStep("Preencher Limite Recebimento", "Campo Limite Recebimento preenchido");
                    Global.processTest.DoStep("Alterar tipo do pedido", "Alteração do tipo do pedido com sucesso");
                    Global.processTest.DoStep("Alterar tipo do acordo", "Alteração do tipo do acordo com sucesso");
                    DefineSteps("Incluir lote com produtos inativos");
                    Global.processTest.DoStep("Validar duplo click no campo QTD sugerida", "Duplo click no campo QTD sugerida exibe aviso");
                    Global.processTest.DoStep("Edição do campo QtdeCompra no grid de produtos para lotes com Incorpora CD", "Campo QtdeCompra no grid de produtos é editável");
                    Global.processTest.DoStep("Preencher quantidade de compra dos produtos", "Preenchimento quantidade de compra dos produto com sucesso");
                    Global.processTest.DoStep($"Validar quantidade de compra dos produtos", "Validação da quantidade de compra por produto e quantidade com sucesso");
                    Global.processTest.DoStep("Gerar Pedidos", "Pedidos gerados com sucesso");
                    Global.processTest.DoStep("Confirmar janela Manutenção de Acordos Promocionais", "Confirmação janela Manutenção de Acordos Promocionais com sucesso"); break;
                case "CriarLoteDeCompraFLVComprador":
                    DefineSteps("Criar capa lote");
                    DefineSteps("Adicionar lojas");
                    Global.processTest.DoStep("Habilitar checkbox Restringe Empresa Loja", "Habilitação checkbox Restringe Empresa Loja com sucesso");
                    DefineSteps("Incluir lote com produtos inativos");
                    Global.processTest.DoStep("Confirmar janela Tributação", "Confirmação janela Tributação com sucesso");
                    Global.processTest.DoStep("Resgatar id do lote", "Resgate do id do lote com sucesso");
                    Global.processTest.DoStep($"Fechar app", "App fechado com sucesso");
                    break;
                case "FinalizarLoteDeCompraFLVComprador":
                    base.DefineSteps("RealizarLogin");
                    DefineSteps("Abrir Gerenciador de Compras");
                    Global.processTest.DoStep($"Abrir lote de compras", "Lote aberto com sucesso");
                    Global.processTest.DoStep($"Validar quantidade de compra dos produtos", "Validação da quantidade de compra por produto e quantidade com sucesso");
                    DefineSteps("Gerar pedidos");
                    break;
                case "PreencherLoteDeCompraFLVChefeSessao":
                    base.DefineSteps("RealizarLogin");
                    DefineSteps("Abrir Gerenciador de Compras");
                    Global.processTest.DoStep("Abrir lote de compras", "Lote de compras aberto com sucesso");
                    Global.processTest.DoStep("Preencher quantidade de compra dos produtos", "Preenchimento quantidade de compra dos produto com sucesso");
                    Global.processTest.DoStep($"Validar quantidade de compra dos produtos", "Validação da quantidade de compra por produto e quantidade com sucesso");
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
                case "CriarCapaLoteLojaALoja":
                    DefineSteps("Criar capa lote");
                    Global.processTest.DoStep("Incluir lote de compra", "Inclusão do lote de compra com sucesso");
                    Global.processTest.DoStep("Confirmar janela Filtros para Seleção de Produtos", "Confirmação janela Seleção de Produtos com sucesso");
                    Global.processTest.DoStep("Confirmar janela Tributação", "Confirmação janela Tributação com sucesso");
                    Global.processTest.DoStep("Resgatar id do lote", "Resgate do id do lote com sucesso");
                    Global.processTest.DoStep($"Fechar app", "App fechado com sucesso");
                    break;

                case "CriarCapaLoteIncorporaCD":
                    DefineSteps("Criar capa lote");
                    DefineSteps("Adicionar lojas");
                    Global.processTest.DoStep("Habilitar checkbox Incorporar Sugestão CD", "Habilitação do checkbox Incorporar Sugestão com sucesso");
                    DefineSteps("Incluir lote com produtos inativos");
                    Global.processTest.DoStep("Resgatar id do lote", "Resgate do id do lote com sucesso");
                    Global.processTest.DoStep($"Fechar app", "App fechado com sucesso");
                    break;
                case "ValidarAlteracaoPrazoPagamento":
                    base.DefineSteps("RealizarLogin");
                    DefineSteps("Abrir Gerenciador de Compras");
                    Global.processTest.DoStep($"Abrir lote de compras", "Lote aberto com sucesso");
                    Global.processTest.DoStep($"Validar Prazo de Pagamento", "Validação do Prazo de Pagamento realizada com sucesso");
                    Global.processTest.DoStep($"Validar Prazo de Pagamento", "Validação do Prazo de Pagamento realizada com sucesso");
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

        private void AddLojas(List<string> lojas, string divisao, int qtdLojas)
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep("Adicionar lojas", logMsg: $"Adicionar {lojas.Count} lojas", paramName: "lojas",
                paramValue: "");
            try
            {
                gerenciadorDeComprasPO.OpenSelecaoDeLojas();
                WaitSeconds(2);
                gerenciadorDeComprasPO.RemoveDivisoes();
                gerenciadorDeComprasPO.AddDivisao(divisao);
                gerenciadorDeComprasPO.RemoveLojas();
                gerenciadorDeComprasPO.AddLojasPorNome(lojas);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Adição de {qtdLojas} lojas com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"Erro na adição de {qtdLojas} lojas");
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
       logMsg: $"Tentando preencher quantidade de compra do(s) produto(s)",
       paramName: "tipoLote, qtdProdutos, qtdeCompra, qtdLojas", paramValue: $"{tipoLote}, {qtdProdutos}, {qtdeCompra}, {qtdLojas}");
            try
            {
                gerenciadorDeComprasPO.FillQtdeCompra(qtdProdutos: qtdProdutos, qtdeCompra: qtdeCompra, tipoLote: tipoLote);
                printFileName = Global.processTest.CaptureWholeScreen();
                if (tipoLote == "cd")
                {
                    RestoreWindow();
                }
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg:
                    $"Preenchido: {qtdProdutos} produto(s) com qtdeCompra: {qtdeCompra}");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName,
                    logMsg: $"Erro ao tentar preencher quantidade de compra do(s) produto(s)");
            }
        }

        private void ValidateQtdeComprasValue(int qtdProdutos, int qtdeCompra, int qtdLojas, string tipoLote)
        {
            string printFileName;
            if (tipoLote == "cd")
            {
                int lgsID = Global.processTest.StartStep($"Validar quantidade de compra dos produtos",
             logMsg: $"Tentando validar quantidade de compra dos produtos",
             paramName: "qtdProdutos, qtdeCompra", paramValue: $"{qtdProdutos}, {qtdeCompra}");

                int total = gerenciadorDeComprasPO.GetQtdeComprValue();
                bool totalIsValid = gerenciadorDeComprasPO.ValidateQtdeComprasValue(total, qtdProdutos, qtdeCompra, qtdLojas, tipoLote);

                if (!totalIsValid)
                {
                    printFileName = Global.processTest.CaptureWholeScreen();
                    Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName,
                        logMsg: $"Erro, esperado: {qtdProdutos * qtdeCompra}, atual: {total}");
                    throw new Exception($"Erro, esperado: {qtdProdutos * qtdeCompra}, atual: {total}");
                }

                WaitSeconds(3);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Sucesso, esperado: {qtdProdutos * qtdeCompra}, atual: {total}");
            }
            else
            {
                int lgsID = Global.processTest.StartStep($"Validar quantidade de compra dos produtos",
             logMsg: $"Tentando validar quantidade de compra dos produtos",
             paramName: "qtdProdutos, qtdeCompra, qtdLojas", paramValue: $"{qtdProdutos}, {qtdeCompra}, {qtdLojas}");

                int total = gerenciadorDeComprasPO.GetQtdeComprValue();
                bool totalIsValid = gerenciadorDeComprasPO.ValidateQtdeComprasValue(total, qtdProdutos, qtdeCompra, qtdLojas, tipoLote);

                if (!totalIsValid)
                {
                    printFileName = Global.processTest.CaptureWholeScreen();
                    Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName,
                        logMsg: $"Erro, esperado: {qtdProdutos * qtdeCompra * qtdLojas}, atual: {total}");
                    throw new Exception($"Erro, esperado: {qtdProdutos * qtdeCompra * qtdLojas}, atual: {total}");
                }

                WaitSeconds(3);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Sucesso, esperado: {qtdProdutos * qtdeCompra * qtdLojas}, atual: {total}");
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
                case "Manutenção de Acordos Promocionais":
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

        private void UpdateTipoPedido(string tipoPedido)
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep($"Alterar tipo do pedido",
                logMsg: $"Tentando alterar o tipo de pedido para {tipoPedido}",
                paramName: "tipoLote", paramValue: tipoPedido);
            try
            {
                gerenciadorDeComprasPO.UpdateTipoPedido(tipoPedido);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Alteração do tipo do pedido com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg:
                    $"Erro ao tentar alterar tipo do pedido");
            }
        }

        private void FillLimiteRecebimento(string dataAtual)
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep($"Preencher Limite Recebimento",
                logMsg: $"Preencher Limite Recebimento com {dataAtual}",
                paramName: "dataAtual", paramValue: dataAtual);
            try
            {
                gerenciadorDeComprasPO.FillLimiteRecebimento(dataAtual);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Campo Limite Recebimento preenchido");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg:
                    $"Erro ao tentar preencher campo Limite Recebimento");
            }
        }

        private void UpdateTipoAcordo(string tipoAcordo)
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep($"Alterar tipo do acordo",
                logMsg: $"Alterar tipo acordo para {tipoAcordo}",
                paramName: "tipoAcordo", paramValue: tipoAcordo);
            try
            {
                gerenciadorDeComprasPO.UpdateTipoAcordo(tipoAcordo);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Tipo acordo alterado.");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg:
                    $"Erro ao tentar alterar tipo acordo.");
            }
        }

        private void ValidatePrazoPagamento(string prazoPagamento)
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep($"Validar Prazo de Pagamento",
                logMsg: $"Alterar Prazo de Pagamento para {prazoPagamento}",
                paramName: "prazoPagamento", paramValue: prazoPagamento);
            try
            {
                gerenciadorDeComprasPO.ValidatePrazoPagamento(prazoPagamento);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Prazo de Pagamento alterado.");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg:
                    $"Erro ao tentar alterar Prazo de Pagamento.");
            }
        }

        [TestMethod]
        public void CriarLoteDeCompraLojaALoja()
        {
            int testId = 2;
            string queryName = TestFactory.GetCurrentMethodName();
            InputData inputExcel = new InputData(ConnType: "Excel", ConnXLS: excelFilePath);
            inputExcel.NewQuery(
                QueryName: queryName,
                QueryText: $"SELECT * FROM [GerenciadorDeCompras$] WHERE testId = {testId}"
                );

            // Test Variables
            string codFornecedor = inputExcel.GetValue("CODFORNECEDOR", queryName);
            string categoria = inputExcel.GetValue("CATEGORIA", queryName);
            string diasAbastecimento = inputExcel.GetValue("DIASABASTECIMENTO", queryName);
            string comprador = inputExcel.GetValue("COMPRADOR", queryName);
            int qtdLojas = int.Parse(inputExcel.GetValue("QTDLOJAS", queryName));
            int qtdProdutos = int.Parse(inputExcel.GetValue("QTDPRODUTOS", queryName));
            int qtdeCompra = int.Parse(inputExcel.GetValue("QTDECOMPRA", queryName));
            string tipoLote = inputExcel.GetValue("TIPOLOTE", queryName);

            // Start Test
            StartTest(inputExcel, queryName);

            // Test Details
            DoTest(inputExcel, queryName);

            // Steps Definition
            DefineSteps(TestFactory.GetCurrentMethodName());

            Login(inputExcel, queryName);
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
            ValidateQtdeComprasValue(qtdProdutos: qtdProdutos, qtdeCompra: qtdeCompra, tipoLote: tipoLote, qtdLojas: qtdLojas);
            GeneratePedidos(tipoLote);
            ConfirmWindow("Consulta Lote de Compra");

            //Teardown function
            EndTest(inputExcel, queryName);
        }

        [TestMethod]
        public void CriarLoteDeCompraIncorporaCD()
        {
            int testId = 3;
            string queryName = TestFactory.GetCurrentMethodName();
            InputData inputExcel = new InputData(ConnType: "Excel", ConnXLS: excelFilePath);
            inputExcel.NewQuery(
                QueryName: queryName,
                QueryText: $"SELECT * FROM [GerenciadorDeCompras$] WHERE testId = {testId}"
                );

            // Test Variables
            List<string> lojas = StringHandler.ParseStringToList(inputExcel.GetValue("LOJAS", queryName));
            string divisao = inputExcel.GetValue("DIVISAO", queryName);
            string cdNome = inputExcel.GetValue("CDNOME", queryName);
            string codFornecedor = inputExcel.GetValue("CODFORNECEDOR", queryName);
            string categoria = inputExcel.GetValue("CATEGORIA", queryName);
            string diasAbastecimento = inputExcel.GetValue("DIASABASTECIMENTO", queryName);
            int qtdProdutos = int.Parse(inputExcel.GetValue("QTDPRODUTOS", queryName));
            int qtdeCompra = int.Parse(inputExcel.GetValue("QTDECOMPRA", queryName));
            int qtdLojas = int.Parse(inputExcel.GetValue("QTDLOJAS", queryName));
            string tipoLote = inputExcel.GetValue("TIPOLOTE", queryName);

            // Start Test
            StartTest(inputExcel, queryName);

            // Test Details
            DoTest(inputExcel, queryName);

            // Steps Definition
            DefineSteps(TestFactory.GetCurrentMethodName());

            Login(inputExcel, queryName);
            OpenGerenciadorDeCompras();
            FillFornecedor(codFornecedor);
            SelectCategoria(categoria);
            FillAbastecimentoDias(diasAbastecimento);
            EnableCheckbox("Sugestão de compra");
            AddLojas(lojas, divisao, qtdLojas);
            ConfirmWindow("Seleção de Empresas do Lote");
            EnableCheckbox(feature: "Incorporar Sugestão CD", paramName: "cdNome", paramValue: cdNome);
            IncludeLote();
            ConfirmWindow("Filtros para Seleção de Produtos");
            ConfirmWindow("Produtos Inativos");
            ValidateDoubleClickOnQtdSugerida();
            ValidateProductsGridEdit(qtdeCompra);
            FillProdutos(qtdProdutos, qtdeCompra, qtdLojas, tipoLote);
            ValidateQtdeComprasValue(qtdProdutos: qtdProdutos, qtdeCompra: qtdeCompra, tipoLote: tipoLote, qtdLojas: qtdLojas);
            GeneratePedidos(tipoLote);
            ConfirmWindow("Consulta Lote de Compra");

            // Teardown function
            EndTest(inputExcel, queryName);
        }

        public void CriarLoteDeCompraFLVComprador()
        {
            int testId = 4;
            string queryName = TestFactory.GetCurrentMethodName();
            InputData inputExcel = new InputData(ConnType: "Excel", ConnXLS: excelFilePath);
            inputExcel.NewQuery(
                QueryName: queryName,
                QueryText: $"SELECT * FROM [GerenciadorDeCompras$] WHERE testId = {testId}"
                );

            // Test Variables
            string codFornecedor = inputExcel.GetValue("CODFORNECEDOR", queryName);
            List<string> lojas = StringHandler.ParseStringToList(inputExcel.GetValue("LOJAS", queryName));
            int qtdLojas = int.Parse(inputExcel.GetValue("QTDLOJAS", queryName));
            string divisao = inputExcel.GetValue("DIVISAO", queryName);
            string categoria = inputExcel.GetValue("CATEGORIA", queryName);
            string diasAbastecimento = inputExcel.GetValue("DIASABASTECIMENTO", queryName);

            // Start Test
            StartTest(inputExcel, queryName);

            // Test Details
            DoTest(inputExcel, queryName);

            // Steps Definition
            DefineSteps(TestFactory.GetCurrentMethodName());

            Login(inputExcel, queryName);
            OpenGerenciadorDeCompras();
            FillFornecedor(codFornecedor);
            SelectCategoria(categoria);
            FillAbastecimentoDias(diasAbastecimento);
            EnableCheckbox("Sugestão de compra");
            AddLojas(lojas, divisao, qtdLojas);
            ConfirmWindow("Seleção de Empresas do Lote");
            EnableCheckbox("Restringe Empresa Loja");
            IncludeLote();
            ConfirmWindow("Filtros para Seleção de Produtos");
            ConfirmWindow("Produtos Inativos");
            ConfirmWindow("Tributação");
            GetLoteId();
            CloseApp();

            // Teardown function
            EndTest(inputExcel, queryName);
        }

        public void FinalizarLoteDeCompraFLVComprador()
        {
            int testId = 8;
            string queryName = TestFactory.GetCurrentMethodName();
            InputData inputExcel = new InputData(ConnType: "Excel", ConnXLS: excelFilePath);
            inputExcel.NewQuery(
                QueryName: queryName,
                QueryText: $"SELECT * FROM [GerenciadorDeCompras$] WHERE testId = {testId}"
                );

            // Test Variables
            string tipoLote = inputExcel.GetValue("TIPOLOTE", queryName);
            int qtdProdutos = int.Parse(inputExcel.GetValue("QTDPRODUTOS", queryName));
            int qtdeCompra = int.Parse(inputExcel.GetValue("QTDECOMPRA", queryName));
            int qtdLojas = int.Parse(inputExcel.GetValue("QTDLOJAS", queryName));

            // Start Test
            StartTest(inputExcel, queryName);

            // Test Details
            DoTest(inputExcel, queryName);

            // Steps Definition
            DefineSteps(TestFactory.GetCurrentMethodName());

            Login(inputExcel, queryName);
            OpenGerenciadorDeCompras();
            OpenLote(idLote);
            ValidateQtdeComprasValue(qtdProdutos: qtdProdutos, qtdeCompra: qtdeCompra, tipoLote: tipoLote, qtdLojas: qtdLojas);
            GeneratePedidos(tipoLote);
            ConfirmWindow("Consulta Lote de Compra");

            // Teardown function
            EndTest(inputExcel, queryName);
        }

        public void PreencherLoteDeCompraFLVChefeSessao()
        {
            // Global Variables
            List<int> testIds = [5, 6, 7];

            for (int i = 0; i < testIds.Count; i++)
            {
                int testId = testIds[i];
                string queryName = TestFactory.GetCurrentMethodName();
                InputData inputExcel = new InputData(ConnType: "Excel", ConnXLS: excelFilePath);
                inputExcel.NewQuery(
                    QueryName: queryName,
                    QueryText: $"SELECT * FROM [GerenciadorDeCompras$] WHERE testId = {testId}"
                    );

                // Test Variables
                List<string> lojas = StringHandler.ParseStringToList(inputExcel.GetValue("LOJAS", queryName));
                int qtdProdutos = int.Parse(inputExcel.GetValue("QTDPRODUTOS", queryName));
                int qtdeCompra = int.Parse(inputExcel.GetValue("QTDECOMPRA", queryName));
                int qtdLojas = lojas.Count;
                string tipoLote = inputExcel.GetValue("TIPOLOTE", queryName);

                // Start Test
                StartTest(inputExcel, queryName);

                // Test Details
                DoTest(inputExcel, queryName);

                // Steps Definition
                DefineSteps(TestFactory.GetCurrentMethodName());

                Login(inputExcel, queryName);
                OpenGerenciadorDeCompras();
                OpenLote(idLote);
                FillProdutos(qtdProdutos, qtdeCompra, qtdLojas, tipoLote);
                ValidateQtdeComprasValue(qtdProdutos: qtdProdutos, qtdeCompra: qtdeCompra, tipoLote: tipoLote, qtdLojas: qtdLojas);
                CloseApp();

                // Teardown function
                EndTest(inputExcel, queryName);
            }
        }

        [TestMethod]
        public void CriarLoteDeCompraFLVCompleto()
        {
            CriarLoteDeCompraFLVComprador();
            PreencherLoteDeCompraFLVChefeSessao();
            FinalizarLoteDeCompraFLVComprador();
        }

        [TestMethod]
        public void CriarLoteDeCompraLojaALojaBonificacao()
        {
            int testId = 9;
            string queryName = TestFactory.GetCurrentMethodName();
            InputData inputExcel = new InputData(ConnType: "Excel", ConnXLS: excelFilePath);
            inputExcel.NewQuery(
                QueryName: queryName,
                QueryText: $"SELECT * FROM [GerenciadorDeCompras$] WHERE testId = {testId}"
                );

            // Test Variables
            string dataAtual = DateHandler.GetTodaysDate().ToString("ddMMyyyy");
            string codFornecedor = inputExcel.GetValue("CODFORNECEDOR", queryName);
            string categoria = inputExcel.GetValue("CATEGORIA", queryName);
            string diasAbastecimento = inputExcel.GetValue("DIASABASTECIMENTO", queryName);
            int qtdLojas = int.Parse(inputExcel.GetValue("QTDLOJAS", queryName));
            int qtdProdutos = int.Parse(inputExcel.GetValue("QTDPRODUTOS", queryName));
            int qtdeCompra = int.Parse(inputExcel.GetValue("QTDECOMPRA", queryName));
            string tipoLote = inputExcel.GetValue("TIPOLOTE", queryName);
            string tipoPedido = inputExcel.GetValue("TIPOPEDIDO", queryName);
            string tipoAcordo = inputExcel.GetValue("TIPOACORDO", queryName);

            // Start Test
            StartTest(inputExcel, queryName);

            // Test Details
            DoTest(inputExcel, queryName);

            // Steps Definition
            DefineSteps(TestFactory.GetCurrentMethodName());

            Login(inputExcel, queryName);
            OpenGerenciadorDeCompras();
            FillFornecedor(codFornecedor);
            SelectCategoria(categoria);
            FillAbastecimentoDias(diasAbastecimento);
            EnableCheckbox("Sugestão de compra");
            FillLimiteRecebimento(dataAtual);
            UpdateTipoPedido(tipoPedido);
            UpdateTipoAcordo(tipoAcordo);
            IncludeLote();
            ConfirmWindow("Filtros para Seleção de Produtos");
            ConfirmWindow("Tributação");
            FillProdutos(qtdProdutos, qtdeCompra, qtdLojas, tipoLote);
            ValidateQtdeComprasValue(qtdProdutos: qtdProdutos, qtdeCompra: qtdeCompra, tipoLote: tipoLote, qtdLojas: qtdLojas);
            GeneratePedidos(tipoLote);
            ConfirmWindow("Manutenção de Acordos Promocionais");

            //Teardown function
            EndTest(inputExcel, queryName);
        }

        [TestMethod]
        public void CriarLoteDeCompraIncorporaCDBonificacao()
        {
            int testId = 10;
            string queryName = TestFactory.GetCurrentMethodName();
            InputData inputExcel = new InputData(ConnType: "Excel", ConnXLS: excelFilePath);
            inputExcel.NewQuery(
                QueryName: queryName,
                QueryText: $"SELECT * FROM [GerenciadorDeCompras$] WHERE testId = {testId}"
                );

            // Test Variables
            string dataAtual = DateHandler.GetTodaysDate().ToString("ddMMyyyy");
            List<string> lojas = StringHandler.ParseStringToList(inputExcel.GetValue("LOJAS", queryName));
            string divisao = inputExcel.GetValue("DIVISAO", queryName);
            string cdNome = inputExcel.GetValue("CDNOME", queryName);
            string codFornecedor = inputExcel.GetValue("CODFORNECEDOR", queryName);
            string categoria = inputExcel.GetValue("CATEGORIA", queryName);
            string diasAbastecimento = inputExcel.GetValue("DIASABASTECIMENTO", queryName);
            int qtdLojas = int.Parse(inputExcel.GetValue("QTDLOJAS", queryName));
            int qtdProdutos = int.Parse(inputExcel.GetValue("QTDPRODUTOS", queryName));
            int qtdeCompra = int.Parse(inputExcel.GetValue("QTDECOMPRA", queryName));
            string tipoLote = inputExcel.GetValue("TIPOLOTE", queryName);
            string tipoPedido = inputExcel.GetValue("TIPOPEDIDO", queryName);
            string tipoAcordo = inputExcel.GetValue("TIPOACORDO", queryName);

            // Start Test
            StartTest(inputExcel, queryName);

            // Test Details
            DoTest(inputExcel, queryName);

            // Steps Definition
            DefineSteps(TestFactory.GetCurrentMethodName());

            Login(inputExcel, queryName);
            OpenGerenciadorDeCompras();
            FillFornecedor(codFornecedor);
            SelectCategoria(categoria);
            FillAbastecimentoDias(diasAbastecimento);
            EnableCheckbox("Sugestão de compra");
            AddLojas(lojas, divisao, qtdLojas);
            ConfirmWindow("Seleção de Empresas do Lote");
            EnableCheckbox(feature: "Incorporar Sugestão CD", paramName: "cdNome", paramValue: cdNome);
            FillLimiteRecebimento(dataAtual);
            UpdateTipoPedido(tipoPedido);
            UpdateTipoAcordo(tipoAcordo);
            IncludeLote();
            ConfirmWindow("Filtros para Seleção de Produtos");
            ConfirmWindow("Produtos Inativos");
            ValidateDoubleClickOnQtdSugerida();
            ValidateProductsGridEdit(qtdeCompra);
            FillProdutos(qtdProdutos, qtdeCompra, qtdLojas, tipoLote);
            ValidateQtdeComprasValue(qtdProdutos: qtdProdutos, qtdeCompra: qtdeCompra, tipoLote: tipoLote, qtdLojas: qtdLojas);

            GeneratePedidos(tipoLote);
            ConfirmWindow("Manutenção de Acordos Promocionais");

            //Teardown function
            EndTest(inputExcel, queryName);
        }

        public void CriarCapaLoteLojaALoja()
        {
            int testId = 11;
            string queryName = TestFactory.GetCurrentMethodName();
            InputData inputExcel = new InputData(ConnType: "Excel", ConnXLS: excelFilePath);
            inputExcel.NewQuery(
                QueryName: queryName,
                QueryText: $"SELECT * FROM [GerenciadorDeCompras$] WHERE testId = {testId}"
                );

            // Test Variables
            string codFornecedor = inputExcel.GetValue("CODFORNECEDOR", queryName);
            string categoria = inputExcel.GetValue("CATEGORIA", queryName);
            string diasAbastecimento = inputExcel.GetValue("DIASABASTECIMENTO", queryName);

            // Start Test
            StartTest(inputExcel, queryName);

            // Test Details
            DoTest(inputExcel, queryName);

            // Steps Definition
            DefineSteps(TestFactory.GetCurrentMethodName());

            Login(inputExcel, queryName);
            OpenGerenciadorDeCompras();
            FillFornecedor(codFornecedor);
            SelectCategoria(categoria);
            FillAbastecimentoDias(diasAbastecimento);
            EnableCheckbox("Sugestão de compra");
            IncludeLote();
            ConfirmWindow("Filtros para Seleção de Produtos");
            ConfirmWindow("Tributação");
            GetLoteId();
            CloseApp();

            //Teardown function
            EndTest(inputExcel, queryName);
        }

        public void ValidarAlteracaoPrazoPagamento(string tipoLote)
        {
            int testId = 12;
            string queryName = TestFactory.GetCurrentMethodName();
            InputData inputExcel = new InputData(ConnType: "Excel", ConnXLS: excelFilePath);
            inputExcel.NewQuery(
                QueryName: queryName,
                QueryText: $"SELECT * FROM [GerenciadorDeCompras$] WHERE testId = {testId}"
                );

            // Test Variables
            List<string> prazoPagamento = StringHandler.ParseStringToList(inputExcel.GetValue("PRAZOPAGAMENTO", queryName));

            // Start Test
            StartTest(inputExcel, queryName);

            // Test Details
            DoTest(inputExcel, queryName);

            // Steps Definition
            DefineSteps(TestFactory.GetCurrentMethodName());

            Login(inputExcel, queryName);
            OpenGerenciadorDeCompras();
            OpenLote(idLote);
            WaitSeconds(20);
            ValidatePrazoPagamento(prazoPagamento[0]);
            ValidatePrazoPagamento(prazoPagamento[1]);

            //Teardown function
            EndTest(inputExcel, queryName);
        }

        public void CriarCapaLoteIncorporaCD()
        {
            int testId = 13;
            string queryName = TestFactory.GetCurrentMethodName();
            InputData inputExcel = new InputData(ConnType: "Excel", ConnXLS: excelFilePath);
            inputExcel.NewQuery(
                QueryName: queryName,
                QueryText: $"SELECT * FROM [GerenciadorDeCompras$] WHERE testId = {testId}"
                );

            // Test Variables
            List<string> lojas = StringHandler.ParseStringToList(inputExcel.GetValue("LOJAS", queryName));
            string divisao = inputExcel.GetValue("DIVISAO", queryName);
            string cdNome = inputExcel.GetValue("CDNOME", queryName);
            string codFornecedor = inputExcel.GetValue("CODFORNECEDOR", queryName);
            string categoria = inputExcel.GetValue("CATEGORIA", queryName);
            string diasAbastecimento = inputExcel.GetValue("DIASABASTECIMENTO", queryName);
            int qtdLojas = int.Parse(inputExcel.GetValue("QTDLOJAS", queryName));

            // Start Test
            StartTest(inputExcel, queryName);

            // Test Details
            DoTest(inputExcel, queryName);

            // Steps Definition
            DefineSteps(TestFactory.GetCurrentMethodName());

            Login(inputExcel, queryName);
            OpenGerenciadorDeCompras();
            FillFornecedor(codFornecedor);
            SelectCategoria(categoria);
            FillAbastecimentoDias(diasAbastecimento);
            EnableCheckbox("Sugestão de compra");
            AddLojas(lojas, divisao, qtdLojas);
            ConfirmWindow("Seleção de Empresas do Lote");
            EnableCheckbox(feature: "Incorporar Sugestão CD", paramName: "cdNome", paramValue: cdNome);
            IncludeLote();
            ConfirmWindow("Filtros para Seleção de Produtos");
            ConfirmWindow("Produtos Inativos");
            GetLoteId();
            CloseApp();

            // Teardown function
            EndTest(inputExcel, queryName);
        }

        [TestMethod]
        public void ValidarAlteracaoPrazoPagamentoLojaALojaCompleto()
        {
            CriarCapaLoteLojaALoja();
            ValidarAlteracaoPrazoPagamento("Loja a Loja");
        }

        [TestMethod]
        public void ValidarAlteracaoPrazoPagamentoIncorporaCDCompleto()
        {
            CriarCapaLoteIncorporaCD();
            ValidarAlteracaoPrazoPagamento("Incorpora CD");
        }
    }
}