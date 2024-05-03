using Consinco.Helpers;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows.Documents;
using System.Xml.Linq;
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

        public void ConfirmWindow(string windowName, int buttonIndex = 0)
        {
            WindowsElement foundWindow = elementHandler.FindElementByName(windowName);
            ReadOnlyCollection<AppiumWebElement> buttons = foundWindow.FindElementsByClassName("Button");
            AppiumWebElement button = buttons[buttonIndex];
            button.Click();
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

        public void ConfirmSelecaoDeLojasWindow()
        {
            string windowName = "Seleção de Empresas do Lote";
            ConfirmWindow(windowName, 1);
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

        public void IncluirLote()
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

        public void ConfirmSelecaoDeProdutosWindow()
        {
            //WindowsElement confirmButton = elementHandler.FindElementByName("Estoque em Dias");
            //confirmButton.Click();
            ConfirmWindow("Filtros para Seleção de Produtos", 6);
        }

        public void ConfirmTributacaoWindow()
        {
            ConfirmWindow("Tributação");
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
                    // Grid de cima
                    BoundingRectangle segundoProduto = new BoundingRectangle(893, 187, 945, 200);
                    WinAppDriver.ClickOn(segundoProduto);
                    WinAppDriver.WaitSeconds(2);

                    //Grid de baixo
                    BoundingRectangle qtdeCompraFirstProduct = new BoundingRectangle(548, 435, 600, 448);
                    WinAppDriver.ClickOn(qtdeCompraFirstProduct);

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

        public void ValidateQtdeComprasValue(int qtdProdutos, int qtdeCompra)
        {
            string qtdeComprValue = GetQtdeComprValue();
            int qtdeComprasValue = int.Parse(qtdeComprValue);
            if (qtdeComprasValue != qtdProdutos * qtdeCompra)
            {
                throw new Exception($"Erro no preenchimento: qtdeComprasValue atual: {qtdeComprasValue}, Total esperado: {qtdProdutos * qtdeCompra}");
            }
        }

        public void ClickGerarPedidos()
        {
            AppiumWebElement geraPedidosButton = elementHandler.FindElementByName("Gera Pedidos");
            geraPedidosButton.Click();
        }

        public string GetQtdeComprValue()
        {
            WinAppDriver.WaitSeconds(3);
            WinAppDriver.Maximize();
            BoundingRectangle qtdeCompraPos = new BoundingRectangle(1047, 321, 1099, 334);
            WinAppDriver.ClickOn(qtdeCompraPos);
            string className = "Edit";
            ReadOnlyCollection<WindowsElement> editElements = elementHandler.FindElementsByClassName(className);
            WindowsElement qtdeCompraEdit = editElements[9];
            string qtdeComprValue = qtdeCompraEdit.GetAttribute("Value.Value");
            return qtdeComprValue;
        }

        public void ConfirmPedidosWindow()
        {
            string windowName = "Opções de geração do(s) pedido(s)";
            ConfirmWindow(windowName, 1); // Atualiza pedido
            ConfirmWindow(windowName, 0); // Confirma validação

            windowName = "Atenção";
            ConfirmWindow(windowName);
        }

        public void ConfirmConsultaLoteCompraWindow()
        {
            WinAppDriver.PressEnter();
            string windowName = "Consulta Lote de Compra";
            ConfirmWindow(windowName);
        }

        public void OpenLote(string idLote)
        {
            // idLote = 312667
            BoundingRectangle pesquisarButton = new BoundingRectangle(151, 78, 179, 106);
            WinAppDriver.ClickOn(pesquisarButton);

            BoundingRectangle sequenciaqLoteField = new BoundingRectangle(114, 137, 165, 157);
            WinAppDriver.ClickOn(sequenciaqLoteField);
            WinAppDriver.FillField(idLote);

            WinAppDriver.SendKey(KeyboardKey.F8);
        }

        public string GetIdLoteDeCompra()
        {
            ReadOnlyCollection<WindowsElement> editList = elementHandler.FindElementsByClassName("Edit");
            return editList[5].GetAttribute("Value.Value");
        }

        public void UpdateLoteDeCompra()
        {
            BoundingRectangle refreshButton = new BoundingRectangle(179, 78, 207, 106);
            WinAppDriver.ClickOn(refreshButton);

            ConfirmWindow("Atenção");
        }
    }
}