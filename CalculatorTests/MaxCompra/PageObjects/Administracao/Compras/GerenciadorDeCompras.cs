using Consinco.Helpers;
using DesktopAppTests.Helpers;
using OpenQA.Selenium.Appium.Windows;
using System.Windows.Automation;

namespace DesktopAppTests.MaxCompra.PageObjects.Administracao.Compras
{
    public class GerenciadorDeComprasPO
    {
        private ElementHandler elementHandler;

        public GerenciadorDeComprasPO(ElementHandler elementHandler)
        {
            this.elementHandler = elementHandler;
        }

        public void FillFornecedor(int codFornecedor)
        {
            var fornecedorField = new ElementHandler.BoundingRectangle(125, 233, 191, 253);
            WinAppDriver.ClickOn(fornecedorField);
            WinAppDriver.FillField(codFornecedor.ToString());
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

            var confirmButton = new ElementHandler.BoundingRectangle(83, 89, 111, 117);
            WinAppDriver.ClickOn(confirmButton);
            ScreenPrinter.CaptureAndSaveScreenshot(Global.screenshotsDirectory, "06-SelectLojas");
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

        public void Incluir()
        {
            //BoundingRectangle	[l=67,t=78,r=95,b=106]
            var incluirButton = new ElementHandler.BoundingRectangle(67, 78, 95, 106);
            WinAppDriver.ClickOn(incluirButton);

            WinAppDriver.WaitForElementVisibleByClassName("Centura:Dialog", 10);
            ScreenPrinter.CaptureAndSaveScreenshot(Global.screenshotsDirectory, "10-Incluir");
        }
    }
}