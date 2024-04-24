using Consinco.Helpers;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using System.Xml.Linq;

namespace DesktopAppTests.MaxCompra.PageObjects.Administracao.Compras
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
        public void ConfirmWindow(string window, int seconds = 60, int buttonIndex = 0)
        {
            var foundWindow = elementHandler.FindElementByName(window, seconds * 1000 / 10);
            var buttons = foundWindow.FindElementsByClassName("Button");
            var exitButton = buttons[buttonIndex];
            exitButton.Click();
        }

        public void FillFornecedor(int codFornecedor)
        {
            var paneEditFields = pane.FindElementsByClassName("Edit");
            var fornecedorField = paneEditFields[2];
            fornecedorField.SendKeys(codFornecedor.ToString());
            WinAppDriver.SendKey(Helpers.KeyboardKey.Tab);
            ScreenPrinter.CaptureAndSaveScreenshot(Global.screenshotsDirectory, "05-FillFornecedor");
        }

        public void AddLojas(int qtdLojas)
        {
            var empresasButton = pane.FindElementByName("Empresas");
            empresasButton.Click();

            //failed
            //var empresasList = elementHandler.FindElementByName("Empresas");
            //var empresasFirstItem = empresasList.FindElementByXPath(".//[@index='1']");
            //WinAppDriver.DoubleTapOn(empresasFirstItem);

            // Adds first X items from the list of Lojas
            var empresasFirstItem = new ElementHandler.BoundingRectangle(80, 391, 207, 404);
            for (int i = 0; i < qtdLojas; i++)
            {
                WinAppDriver.DoubleClickOn(empresasFirstItem);
            }

            ScreenPrinter.CaptureAndSaveScreenshot(Global.screenshotsDirectory, "06a-SelectLojas");

            var buttons = elementHandler.FindElementsByClassName("Button");
            var returnButton = buttons[1];
            returnButton.Click();
            ScreenPrinter.CaptureAndSaveScreenshot(Global.screenshotsDirectory, "06b-SelectLojas");
        }

        public void SelectCategoria(string categoria)
        {
            //var compradorComboBox = new ElementHandler.BoundingRectangle(377, 201, 563, 222);
            //WinAppDriver.ClickOn(compradorComboBox);

            var comboxboxes = pane.FindElementsByClassName("ComboBox");
            var compradorComboBox = comboxboxes[1];
            compradorComboBox.SendKeys(categoria);

            //compradorComboBox.Click();
            //WinAppDriver.FillField(categoria);
            ScreenPrinter.CaptureAndSaveScreenshot(Global.screenshotsDirectory, "07-SelectCategoria");
        }

        public void FillAbastecimentoDias(int dias)
        {
            var abastecimentoDiasField = new ElementHandler.BoundingRectangle(431, 338, 460, 357);
            WinAppDriver.DoubleClickOn(abastecimentoDiasField);
            WinAppDriver.FillField(dias.ToString());
            WinAppDriver.SendKey(Helpers.KeyboardKey.Tab);
            ScreenPrinter.CaptureAndSaveScreenshot(Global.screenshotsDirectory, "08-FillAbastecimentoDias");

        }

        public void EnableCheckBoxesSugestaoDeCompras(string className)
        {
            var checkboxes = elementHandler.FindElementsByClassName(className);

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
            ScreenPrinter.CaptureAndSaveScreenshot(Global.screenshotsDirectory, "09-EnableCheckBoxesSugestaoDeCompras");
        }

        public void IncluirLote()
        {
            var incluirButton = new ElementHandler.BoundingRectangle(67, 78, 95, 106);
            WinAppDriver.ClickOn(incluirButton);

            WinAppDriver.WaitForElementVisibleByClassName("Centura:Dialog", 10);
            ScreenPrinter.CaptureAndSaveScreenshot(Global.screenshotsDirectory, "10-Incluir");
        }

        public void AddCompradores(string comprador)
        {
            var compradoresButton = new ElementHandler.BoundingRectangle(42, 388, 106, 409);
            WinAppDriver.ClickOn(compradoresButton);

            var compradorListItem = elementHandler.FindElementByXPathPartialName(comprador);
            compradorListItem.Click();

            var selecaoDeCompradoresWindow = elementHandler.FindElementByName("Seleção de Compradores");
            var buttons = selecaoDeCompradoresWindow.FindElementsByClassName("Button");
            var addButton = buttons[0];
            addButton.Click();

            WinAppDriver.PressEnter();
            ScreenPrinter.CaptureAndSaveScreenshot(Global.screenshotsDirectory, "11-AddCompradores");

            var confirmButton = elementHandler.FindElementByName("Estoque em Dias");
            confirmButton.Click();
        }

        public void ConfirmProdutosInativosWindow()
        {
            ConfirmWindow("Produtos Inativos", 150);
            ScreenPrinter.CaptureAndSaveScreenshot(Global.screenshotsDirectory, "12-ProdutosInativos");
        }

        public void ConfirmTributacaoWindow()
        {
            ConfirmWindow("Tributação", 120);
            ScreenPrinter.CaptureAndSaveScreenshot(Global.screenshotsDirectory, "13-Tributacao");
        }

        public void FillQtdeCompra(int qtdLojas, int qtdProdutos, int qtdeCompra)
        {
            //var panes = pane.FindElementsByClassName("Centura:ChildTable");
            //var produtosPorLojaPane = panes[1];

            //BoundingRectangle	[l=475,t=453,r=527,b=466]
            var qtdeCompraFirstProduct = new ElementHandler.BoundingRectangle(475, 453, 527, 466);
            WinAppDriver.ClickOn(qtdeCompraFirstProduct);

            for (int i = 0; i < qtdProdutos; i++)
            {
                for (int j = 0; j <= qtdLojas; j++)
                {
                    WinAppDriver.FillField(qtdeCompra.ToString());
                    WinAppDriver.PressEnter();
                    //WinAppDriver.SendKey(Helpers.KeyboardKey.Down);
                }
            }

            //BoundingRectangle	[l=893,t=352,r=945,b=365]
            var qtdeCompraTotal = new ElementHandler.BoundingRectangle(893, 352, 945, 365);
            WinAppDriver.ClickOn(qtdeCompraTotal);

            var qtdeCompraTotalEdit = elementHandler.FindElementByClassName("Edit");
            var qtdeCompraTotalValue = qtdeCompraTotalEdit.GetAttribute("Value");

            if (int.Parse(qtdeCompraTotalValue) == (qtdLojas + 1) * qtdProdutos * qtdeCompra)
            {
                Console.WriteLine("deu boa");
            }
            else
            {
                Console.WriteLine("deu ruim");
            }

            ScreenPrinter.CaptureAndSaveScreenshot(Global.screenshotsDirectory, "14-FillQtdeCompra");
        }

        public void ClickGerarPedidos()
        {
            var geraPedidosButton = elementHandler.FindElementByName("Gera Pedidos");
            geraPedidosButton.Click();

            var windowName = "Atenção";
            ConfirmWindow(windowName);

            windowName = "Opções de geração do(s) pedido(s)";
            WinAppDriver.WaitForElementVisibleByName(windowName, 10);
            ScreenPrinter.CaptureAndSaveScreenshot(Global.screenshotsDirectory, "15-ClickGerarPedidos");
        }

        public void ConfirmPedidosWindow()
        {
            var windowName = "Opções de geração do(s) pedido(s)";
            ConfirmWindow(windowName, 120, 1); // Atualiza pedido
            ConfirmWindow(windowName, 120, 0); // Confirma validação

            windowName = "Atenção";
            ConfirmWindow(windowName);

            windowName = "Lote com Inconsistência";

            WinAppDriver.WaitForElementVisibleByName(windowName, 10);
            ScreenPrinter.CaptureAndSaveScreenshot(Global.screenshotsDirectory, "16-ConfirmPedidosWindow");
        }

        //"PS032528
    }
}
