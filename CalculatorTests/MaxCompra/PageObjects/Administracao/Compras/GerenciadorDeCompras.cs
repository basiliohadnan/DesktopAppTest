using Consinco.Helpers;
using OpenCvSharp;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using System.Collections.ObjectModel;
using static Consinco.Helpers.ElementHandler;

namespace Consinco.MaxCompra.PageObjects.Administracao.Compras
{
    public class GerenciadorDeComprasPO
    {
        private ElementHandler elementHandler;
        private WindowsElement window;
        private AppiumWebElement pane;

        public GerenciadorDeComprasPO(ElementHandler elementHandler)
        {
            this.elementHandler = elementHandler;
            window = elementHandler.FindElementByName("Gerenciador de Compras");
            pane = window.FindElementByClassName("Centura:Form");
        }

        public void ExitWindow(string windowName)
        {
            switch (windowName)
            {
                case "Seleção de Empresas do Lote":
                    elementHandler.ConfirmWindow(windowName, 1);
                    break;
                case "Filtros para Seleção de Produtos":
                    elementHandler.ConfirmWindow(windowName, 6);
                    break;
                case "Pesquisa de Lotes de Compra":
                    elementHandler.ConfirmWindow(windowName, 2);
                    break;
                case "Produtos Inativos":
                case "Tributação":
                    elementHandler.ConfirmWindow(windowName, 0, 2000);
                    break;
                case "Manutenção de Acordos Promocionais":
                    elementHandler.ConfirmWindow("OK");
                    elementHandler.ConfirmWindow(windowName, 0, 3000);
                    break;
                case "Atenção":
                    elementHandler.ConfirmWindow(windowName);
                    break;
                case "Opções de geração do(s) pedido(s)":
                    elementHandler.ConfirmWindow(windowName, 1); // Atualiza pedido
                    elementHandler.ConfirmWindow(windowName, 0); // Confirma validação
                    ExitWindow("Atenção");
                    break;
                case "Consulta Lote de Compra":
                    elementHandler.ConfirmWindow("OK");
                    elementHandler.ConfirmWindow(windowName, 2, 3000);
                    break;
                default:
                    throw new Exception($"Window {windowName} not found.");
            }
        }

        public void FillFornecedor(string codFornecedor)
        {
            ReadOnlyCollection<AppiumWebElement> paneEditFields = pane.FindElementsByClassName("Edit");
            AppiumWebElement fornecedorField = paneEditFields[2];
            fornecedorField.SendKeys(codFornecedor);
            WinAppDriver.PressEnter();
        }

        public void OpenSelecaoDeLojas()
        {
            AppiumWebElement empresasButton = pane.FindElementByName("Empresas");
            empresasButton.Click();
        }

        public void RemoveDivisoes()
        {
            BoundingRectangle removeDivisoesButton = new BoundingRectangle(247, 270, 271, 295);
            WinAppDriver.ClickOn(removeDivisoesButton);
        }

        public void AddDivisao(string divisao)
        {
            var divisaoListItem = elementHandler.FindElementByName(divisao);
            divisaoListItem.Click();

            BoundingRectangle addDivisaoButton = new BoundingRectangle(247, 195, 271, 220);
            WinAppDriver.ClickOn(addDivisaoButton);
        }

        public void RemoveLojas()
        {
            BoundingRectangle removeLojasButton = new BoundingRectangle(247, 497, 271, 522);
            WinAppDriver.ClickOn(removeLojasButton);
        }

        public void AddLojasPorQuantidade(int qtdLojas)
        {
            AppiumWebElement empresasButton = pane.FindElementByName("Empresas");
            empresasButton.Click();

            string windowName = "Seleção de Empresas do Lote";
            try
            {
                elementHandler.FindElementByName(windowName);

                // Adds first X items from the list of Lojas
                BoundingRectangle empresasFirstItem = new BoundingRectangle(80, 391, 207, 404);
                for (int i = 0; i < qtdLojas; i++)
                {
                    WinAppDriver.DoubleClickOn(empresasFirstItem);
                }
            }
            catch
            {
                throw new Exception($"Erro ao tentar adicionar ${qtdLojas} lojas na janela {windowName}");
            }
        }

        public void AddLojasPorNome(List<string> lojas)
        {
            BoundingRectangle addLojasButton = new BoundingRectangle(247, 422, 271, 447);
            string windowName = "Seleção de Empresas do Lote";
            try
            {
                lojas.ForEach(loja =>
                {
                    WindowsElement lojaListItem = elementHandler.FindElementByName(loja);
                    lojaListItem.Click();
                });
                WinAppDriver.ClickOn(addLojasButton);
            }
            catch
            {
                throw new Exception($"Erro ao tentar adicionar {lojas.Count} lojas na janela {windowName}");
            }
        }

        public void SelectCategoria(string categoria)
        {
            var comboxboxes = pane.FindElementsByClassName("ComboBox");
            AppiumWebElement compradorComboBox = comboxboxes[1];
            compradorComboBox.SendKeys(categoria);
        }

        public void FillAbastecimentoDias(string dias)
        {
            BoundingRectangle abastecimentoDiasField = new BoundingRectangle(431, 338, 460, 357);
            WinAppDriver.DoubleClickOn(abastecimentoDiasField);
            WinAppDriver.FillField(dias);
        }

        public void EnableCheckbox(string feature)
        {
            switch (feature)
            {
                case "Sugestão Compras":
                    string checkBoxesClass = "Centura:GPCheck";
                    ReadOnlyCollection<WindowsElement> checkboxes = elementHandler.FindElementsByClassName(checkBoxesClass);

                    for (int i = 14; i <= 18; i++)
                    {
                        var checkbox = checkboxes[i];
                        if (elementHandler.VerifyCheckBoxIsOn(checkbox))
                        {
                            continue;
                        }
                        else
                        {
                            checkbox.Click();
                        }
                    }
                    break;

                case "Restringe Empresa Loja":
                    var checkboxRestringe = elementHandler.FindElementByName(feature);
                    if (!elementHandler.VerifyCheckBoxIsOn(checkboxRestringe))
                    {
                        checkboxRestringe.Click();
                    }
                    break;
                case "Incorporar Sugestão CD":
                    string checkBoxName = "Incorporar Sugestão";
                    var checkboxIncorporar = elementHandler.FindElementByName(checkBoxName);
                    if (!elementHandler.VerifyCheckBoxIsOn(checkboxIncorporar))
                    {
                        checkboxIncorporar.Click();
                    }
                    break;

                default:
                    throw new ArgumentException($"Unsupported feature: {feature}");
            }
        }

        public void SetCD(string cdNome)
        {
            BoundingRectangle cdPrincipalCombobox = new BoundingRectangle(532, 633, 549, 650);
            WinAppDriver.ClickOn(cdPrincipalCombobox);
            WindowsElement chosenCd = elementHandler.FindElementByName(cdNome);
            chosenCd.Click();
        }

        public void IncludeLote()
        {
            BoundingRectangle incluirButton = new BoundingRectangle(67, 78, 95, 106);
            WinAppDriver.ClickOn(incluirButton);
        }

        public void AddCompradores(string comprador)
        {
            elementHandler.FindElementByClassName("Centura:AccFrame");
            BoundingRectangle filtrosButton = new BoundingRectangle(263, 78, 291, 106);
            WinAppDriver.ClickOn(filtrosButton);

            BoundingRectangle compradoresButton = new BoundingRectangle(42, 388, 106, 409);
            WinAppDriver.ClickOn(compradoresButton);

            WindowsElement compradorListItem = elementHandler.FindElementByXPathPartialName(comprador);
            compradorListItem.Click();

            WindowsElement selecaoDeCompradoresWindow = elementHandler.FindElementByName("Seleção de Compradores");
            ReadOnlyCollection<AppiumWebElement> buttons = selecaoDeCompradoresWindow.FindElementsByClassName("Button");
            AppiumWebElement addButton = buttons[0];
            addButton.Click();

            AppiumWebElement exitButton = buttons[4];
            exitButton.Click();
        }

        public void FillQtdeCompra(int qtdProdutos, int qtdeCompra, string tipoLote)
        {
            string gridClassName = "Centura:ChildTable";
            WindowsElement grid = elementHandler.FindElementByClassName(gridClassName);
            if (grid == null)
                throw new Exception("Grid not found");

            WinAppDriver.WaitSeconds(2);

            switch (tipoLote)
            {
                case "loja-a-loja":
                    {
                        BoundingRectangle qtdeCompraFirstLoja = new BoundingRectangle(469, 453, 521, 466);
                        WinAppDriver.ClickOn(qtdeCompraFirstLoja);

                        for (int i = 0; i < qtdProdutos; i++)
                        {
                            WinAppDriver.FillField(qtdeCompra.ToString());
                            WinAppDriver.PressEnter();
                        }
                    }
                    break;
                case "cd":
                    WinAppDriver.MaximizeWindow();

                    // Grid de produtos
                    BoundingRectangle qtdeCompraSegundoProduto = new BoundingRectangle(1047, 156, 1099, 169);
                    WinAppDriver.ClickOn(qtdeCompraSegundoProduto);
                    WinAppDriver.WaitSeconds(3);

                    for (int i = 0; i < qtdProdutos; i++)
                    {
                        WinAppDriver.FillField(qtdeCompra.ToString());
                        WinAppDriver.PressEnter();
                    }
                    break;
                case "flv":
                    {
                        BoundingRectangle qtdeCompraFirstLoja = new BoundingRectangle(469, 453, 521, 466);
                        WinAppDriver.ClickOn(qtdeCompraFirstLoja);
                        for (int i = 0; i < qtdProdutos; i++)
                        {
                            WinAppDriver.FillField(qtdeCompra.ToString());
                            WinAppDriver.PressEnter();
                        }
                    }
                    break;
                default:
                    // Handle any other cases here
                    break;
            }
        }

        public bool ValidateQtdeComprasValue(int total, int qtdProdutos, int qtdeCompra, int qtdLojas, string tipoLote)
        {
            if (tipoLote == "cd")
            {
                return total == qtdProdutos * qtdeCompra;
            }
            else
            {
                return total == qtdProdutos * qtdeCompra * qtdLojas;
            }
        }

        public void ClickGerarPedidos()
        {
            AppiumWebElement geraPedidosButton = elementHandler.FindElementByName("Gera Pedidos");
            geraPedidosButton.Click();
        }

        public void UpdateLote()
        {
            BoundingRectangle refreshButton = new BoundingRectangle(179, 78, 207, 106);
            WinAppDriver.ClickOn(refreshButton);

            elementHandler.ConfirmWindow("Atenção");
        }

        public void OpenLote(string idLote)
        {
            BoundingRectangle pesquisarButton = new BoundingRectangle(151, 78, 179, 106);
            WinAppDriver.ClickOn(pesquisarButton);

            BoundingRectangle sequenciaqLoteField = new BoundingRectangle(114, 137, 165, 157);
            WinAppDriver.ClickOn(sequenciaqLoteField);
            WinAppDriver.FillField(idLote);

            string windowName = "Pesquisa de Lotes de Compra";
            ExitWindow(windowName);
        }

        public int GetQtdeComprValue()
        {
            WinAppDriver.WaitSeconds(3);
            WinAppDriver.MaximizeWindow();
            BoundingRectangle qtdeCompraPos = new BoundingRectangle(1047, 321, 1099, 334);
            WinAppDriver.ClickOn(qtdeCompraPos);
            string className = "Edit";
            ReadOnlyCollection<WindowsElement> editElements = elementHandler.FindElementsByClassName(className);
            WindowsElement qtdeCompraEdit = editElements[9];
            string qtdeComprValue = qtdeCompraEdit.GetAttribute("Value.Value");
            return int.Parse(qtdeComprValue);
        }

        public string GetIdLoteDeCompra()
        {
            ReadOnlyCollection<WindowsElement> editList = elementHandler.FindElementsByClassName("Edit");
            return editList[5].GetAttribute("Value.Value");
        }

        public void DoubleClickOnQtdSugerida()
        {
            WinAppDriver.WaitSeconds(7);
            WinAppDriver.DoubleClickOn(916, 443);
        }

        public void FillProductsGrid(int qtdeCompra)
        {
            // Grid de produtos
            BoundingRectangle qtdeCompraSegundoProduto = new BoundingRectangle(1047, 156, 1099, 169);
            WinAppDriver.ClickOn(qtdeCompraSegundoProduto);
            WinAppDriver.WaitSeconds(3);
            WinAppDriver.FillField(qtdeCompra.ToString());
        }

        public void UpdateTipoPedido(string tipoPedido)
        {
            BoundingRectangle tipoPedidoComboBox = new BoundingRectangle(299, 203, 316, 220);
            WinAppDriver.ClickOn(tipoPedidoComboBox);
            WinAppDriver.FillField(tipoPedido);
            WinAppDriver.PressEnter();
        }

        public void UpdateTipoAcordo(string tipoAcordo)
        {
            switch (tipoAcordo)
            {
                case "DIFERENCA DE PRECO":
                    WindowsElement tipoAcordoButton = elementHandler.FindElementByName("Tipo Acordo");
                    tipoAcordoButton.Click();
                    WinAppDriver.DoubleClickOn(125, 295);
                    break;
                default:
                    throw new Exception($"{tipoAcordo} não encontrado.");
            }
        }

        public void FillLimiteRecebimento(string dataAtual)
        {
            BoundingRectangle limiteRecebimentoEdit = new BoundingRectangle(125, 506, 195, 527);
            WinAppDriver.ClickOn(limiteRecebimentoEdit);
            WinAppDriver.FillField(dataAtual);
        }

        public void ValidatePrazoPagamento(string prazoPagamento)
        {
            switch (prazoPagamento)
            {
                case "Prazo Único":
                    WinAppDriver.ClickOn(30, 120);
                    WindowsElement radioButtonPrazoPagamento = elementHandler.FindElementByName(prazoPagamento);
                    radioButtonPrazoPagamento.Click();
                    WinAppDriver.PressKey("TAB");
                    WinAppDriver.FillField("90");
                    break;
                case "Prazo Fixo":
                    WinAppDriver.ClickOn(30, 120);
                    radioButtonPrazoPagamento = elementHandler.FindElementByName(prazoPagamento);
                    radioButtonPrazoPagamento.Click();
                    WinAppDriver.PressKey("TAB");
                    WinAppDriver.FillField("01012025");
                    break;
                default:
                    throw new Exception($"{prazoPagamento} não encontrado.");
            }
        }
    }
}