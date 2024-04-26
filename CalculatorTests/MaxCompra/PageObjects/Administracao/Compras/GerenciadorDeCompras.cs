using Consinco.Helpers;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using System.Collections.ObjectModel;
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
            AppiumWebElement exitButton = buttons[buttonIndex];
            exitButton.Click();
        }

        public void FillFornecedor(string codFornecedor)
        {
            ReadOnlyCollection<AppiumWebElement> paneEditFields = pane.FindElementsByClassName("Edit");
            AppiumWebElement fornecedorField = paneEditFields[2];
            fornecedorField.SendKeys(codFornecedor);
            WinAppDriver.PressEnter();
        }

        public void AddLojas(int qtdLojas)
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

        public void EnableCheckBoxesSugestaoDeCompras(string className)
        {
            ReadOnlyCollection<WindowsElement> checkboxes = elementHandler.FindElementsByClassName(className);

            for (int i = 14; i <= 18; i++)
            {
                WindowsElement checkbox = checkboxes[i];
                if (elementHandler.VerifyCheckBoxIsOn(checkbox))
                {
                    continue;
                }
                else
                {
                    checkbox.Click();
                }
            }
        }

        public void IncluirLote()
        {
            BoundingRectangle incluirButton = new BoundingRectangle(67, 78, 95, 106);
            WinAppDriver.ClickOn(incluirButton);

            elementHandler.FindElementByClassName("Centura:Dialog");
            BoundingRectangle exitButton = new BoundingRectangle(528, 468, 610, 493);
            WinAppDriver.ClickOn(exitButton);
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

        public void ConfirmProdutosInativosWindow()
        {
            ConfirmWindow("Produtos Inativos");
        }

        public void ConfirmTributacaoWindow()
        {
            ConfirmWindow("Tributação");
        }

        public void FillQtdeCompra(int qtdLojas, int qtdProdutos, int qtdeCompra)
        {
            BoundingRectangle qtdeCompraFirstProduct = new BoundingRectangle(475, 453, 527, 466);
            WinAppDriver.ClickOn(qtdeCompraFirstProduct);

            for (int i = 0; i < qtdProdutos; i++)
            {
                for (int j = 0; j <= qtdLojas; j++)
                {
                    WinAppDriver.FillField(qtdeCompra.ToString());
                    WinAppDriver.PressEnter();
                }
            }

            BoundingRectangle qtdeCompraTotal = new BoundingRectangle(893, 352, 945, 365);
            WinAppDriver.ClickOn(qtdeCompraTotal);

            AppiumWebElement qtdeCompraTotalEdit = elementHandler.FindElementByClassName("Edit");
            string qtdeCompraTotalValue = qtdeCompraTotalEdit.GetAttribute("Value");

            if (int.Parse(qtdeCompraTotalValue) == (qtdLojas + 1) * qtdProdutos * qtdeCompra)
            {
                Console.WriteLine("deu boa");
            }
            else
            {
                Console.WriteLine("deu ruim");
            }
        }

        public void FillQtdeCompra(int qtdProdutos, int qtdeCompra)
        {
            string gridClassName = "Centura:ChildTable";
            WindowsElement grid = elementHandler.FindElementByClassName(gridClassName);

            if (grid != null)
            {
                BoundingRectangle qtdeCompraFirstProduct = new BoundingRectangle(469, 453, 521, 466);
                WinAppDriver.ClickOn(qtdeCompraFirstProduct);

                for (int i = 0; i < qtdProdutos; i++)
                {
                    WinAppDriver.FillField(qtdeCompra.ToString());
                    
                    WinAppDriver.PressEnter();
                }
            }
            else
            {
                throw new Exception("Grid not found");
            }
        }

        public void ClickGerarPedidos()
        {
            AppiumWebElement geraPedidosButton = elementHandler.FindElementByName("Gera Pedidos");
            geraPedidosButton.Click();
        }

        public string GetQtdeComprValue()
        {
            WinAppDriver.Maximize();
            BoundingRectangle qtdeCompraPos = new BoundingRectangle(1047, 138, 1099, 151);
            string className = "Edit";
            string value = elementHandler.GetElementValue(qtdeCompraPos, className);
            return value;

        }

        public void ConfirmPedidosWindow()
        {
            string windowName = "Opções de geração do(s) pedido(s)";
            ConfirmWindow(windowName, 1); // Atualiza pedido
            ConfirmWindow(windowName, 0); // Confirma validação

            windowName = "Atenção";
            ConfirmWindow(windowName);
        }
    }
}
