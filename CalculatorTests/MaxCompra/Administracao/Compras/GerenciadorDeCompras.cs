using Consinco.Helpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium.Windows;
using System.Xml.Linq;

namespace Consinco.MaxCompra.Administracao.Compras
{
    [TestClass]
    public class GerenciadorDeCompras : MaxCompraInit
    {
        private ElementHandler elementHandler;
        public GerenciadorDeCompras()
        {
            elementHandler = new ElementHandler();
        }

        // Inserir POs
        [TestMethod]
        public void CreateLoteLojaALoja()
        {
            Login();
            OpenMenu("Administração", "02-OpenMenuAdm");
            OpenMenu("Compras", "03-OpenSubMenuCompras");
            OpenMenu("Gerenciador de Compras", "04-OpenSubMenuGerenciadorDeCompras");
            FillFornecedor(478);
            SelectLojas(11);
            SelectCategoria("LIQ2 (SUCOS, AG");
            FillAbastecimentoDias(60);
            EnableCheckBoxesSugestaoDeCompras("Centura:GPCheck");
        }

        public void FillFornecedor(int codFornecedor)
        {
            var fornecedorField = new ElementHandler.BoundingRectangle(125, 233, 191, 253);
            ClickOn(fornecedorField);
            FillField(codFornecedor.ToString());
            ScreenPrinter.CaptureAndSaveScreenshot(screenshotsDirectory, "05-FillFornecedor");
        }

        public void SelectLojas(int qtdLojas)
        {
            var empresasButton = new ElementHandler.BoundingRectangle(43, 171, 120, 189);
            ClickOn(empresasButton);

            // Adds first X items from the list of Lojas
            var empresasFirstItem = new ElementHandler.BoundingRectangle(80, 391, 207, 404);
            for (int i = 0; i < qtdLojas; i++)
            {
                DoubleClickOn(empresasFirstItem);
            }

            var confirmButton = new ElementHandler.BoundingRectangle(83, 89, 111, 117);
            ClickOn(confirmButton);
            ScreenPrinter.CaptureAndSaveScreenshot(screenshotsDirectory, "06-SelectLojas");
        }

        public void SelectCategoria(string categoria)
        {
            var compradorComboBox = new ElementHandler.BoundingRectangle(377, 201, 563, 222);
            ClickOn(compradorComboBox);

            FillField(categoria);
            PressEnter();
            ScreenPrinter.CaptureAndSaveScreenshot(screenshotsDirectory, "07-SelectCategoria");
        }

        public void FillAbastecimentoDias(int dias)
        {
            var abastecimentoDiasField = new ElementHandler.BoundingRectangle(431, 338, 460, 357);
            DoubleClickOn(abastecimentoDiasField);
            FillField(dias.ToString());
            ScreenPrinter.CaptureAndSaveScreenshot(screenshotsDirectory, "08-FillAbastecimentoDias");
        }

        // Sugestão de Compras | Central Abastecimento
        public void EnableCheckBoxesSugestaoDeCompras(string className)
        {
            //var pane = elementHandler.FindElementByClassName("Centura:Form");
            var sugestaoDeCompraCheckboxes = elementHandler.FindElementsByClassName(className);

            // Sugestão de Compras | CD
            //Name Considera Saldo Pend Receber
            //Name Considera Saldo Pend Expedir
            //Name	Considera Qtde a Comprar Lote

            // Sugestão de Compras | Loja
            //Name	Considera Saldo Pend Receber
            //Name Considera Saldo Pend Expedir

            // Start at index 14 and loop until 18
            for (int i = 14; i < sugestaoDeCompraCheckboxes.Count && i <= 18; i++)
            {
                // Access each element using the indexer
                WindowsElement checkbox = sugestaoDeCompraCheckboxes[i];

                // Perform actions with the element
                if (elementHandler.VerifyCheckBoxIsOn(checkbox))
                {
                    continue;
                }
                else
                {
                    checkbox.Click();
                }
            }
            ScreenPrinter.CaptureAndSaveScreenshot(screenshotsDirectory ,"09-EnableCheckBoxesSugestaoDeCompras");
        }
    }
}