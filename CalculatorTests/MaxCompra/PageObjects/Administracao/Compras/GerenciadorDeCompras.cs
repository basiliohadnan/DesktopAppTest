﻿using Consinco.Helpers;
using OpenCvSharp;
using OpenQA.Selenium.Appium.Windows;

namespace DesktopAppTests.MaxCompra.PageObjects.Administracao.Compras
{
    public class GerenciadorDeComprasPO
    {
        private ElementHandler elementHandler;

        public GerenciadorDeComprasPO(ElementHandler elementHandler)
        {
            this.elementHandler = elementHandler;
        }

        public void ConfirmWindow(string window, int seconds = 60)
        {
            WinAppDriver.WaitForElementVisibleByName(window, seconds);
            var foundWindow = elementHandler.FindElementByName(window);
            var buttons = foundWindow.FindElementsByClassName("Button");
            var exitButton = buttons[0];
            exitButton.Click();
        }

        public void FillFornecedor(int codFornecedor)
        {
            var fornecedorField = new ElementHandler.BoundingRectangle(125, 233, 191, 253);
            WinAppDriver.ClickOn(fornecedorField);
            WinAppDriver.FillField(codFornecedor.ToString());
            WinAppDriver.SendKey(Helpers.KeyboardKey.Tab);
            ScreenPrinter.CaptureAndSaveScreenshot(Global.screenshotsDirectory, "05-FillFornecedor");
        }

        public void SelectLojas(int qtdLojas)
        {
            var empresasButton = new ElementHandler.BoundingRectangle(43, 171, 120, 189);
            WinAppDriver.WaitSeconds(1);
            WinAppDriver.ClickOn(empresasButton);

            // Adds first X items from the list of Lojas
            var empresasFirstItem = new ElementHandler.BoundingRectangle(80, 391, 207, 404);
            for (int i = 0; i < qtdLojas; i++)
            {
                WinAppDriver.DoubleClickOn(empresasFirstItem);
            }

            ScreenPrinter.CaptureAndSaveScreenshot(Global.screenshotsDirectory, "06a-SelectLojas");

            var confirmButton = new ElementHandler.BoundingRectangle(83, 89, 111, 117);
            WinAppDriver.ClickOn(confirmButton);
            ScreenPrinter.CaptureAndSaveScreenshot(Global.screenshotsDirectory, "06b-SelectLojas");
        }

        public void SelectCategoria(string categoria)
        {
            var compradorComboBox = new ElementHandler.BoundingRectangle(377, 201, 563, 222);
            WinAppDriver.ClickOn(compradorComboBox);

            WinAppDriver.FillField(categoria);
            WinAppDriver.PressEnter();
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
            ConfirmWindow("Produtos Inativos", 120);
            ScreenPrinter.CaptureAndSaveScreenshot(Global.screenshotsDirectory, "12-ProdutosInativos");
        }

        public void ConfirmTributacaoWindow()
        {
            ConfirmWindow("Tributação", 120);
            ScreenPrinter.CaptureAndSaveScreenshot(Global.screenshotsDirectory, "13-Tributacao");
        }

        public void ChecksPaneLojasShown()
        {
            // Pane de Lojas a serem abastecidas
            //BoundingRectangle[l = 14, t = 133, r = 995, b = 388]
            //ClassName Centura:ChildTable
            //WinAppDriver.WaitForElementVisibleByName(window, seconds);
            //var produtosInativosWindow = elementHandler.FindElementByName(window);
        }
        //PS032528
    }
}
