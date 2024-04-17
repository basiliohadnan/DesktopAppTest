using Consinco.Helpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;

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
        }

        public void FillFornecedor(int codFornecedor)
        {
            var fornecedorField = new ElementHandler.BoundingRectangle(125, 233, 191, 253);
            ClickOn(fornecedorField);
            FillField(codFornecedor.ToString());
            ScreenPrinter.CaptureAndSaveScreenshot(Global.appSession, screenshotsDirectory + "\\" + Global.app + "\\" + "05-FillFornecedor");
        }

        // Lojas
        public void SelectLojas(int qtdLojas)
        {
            var empresas = new ElementHandler.BoundingRectangle(43, 171, 120, 189);
            ClickOn(empresas);

            // Adds first X items from the list of Lojas
            var empresaField = new ElementHandler.BoundingRectangle(80, 391, 207, 404);
            for (int i = 0; i < qtdLojas; i++)
            {
                DoubleClickOn(empresaField);
            }

            var confirmButton = new ElementHandler.BoundingRectangle(83, 89, 111, 117);
            ClickOn(confirmButton);
            ScreenPrinter.CaptureAndSaveScreenshot(Global.appSession, screenshotsDirectory + "\\" + Global.app + "\\" + "06-SelectLojas");
        }

        public void SelectCategoria(string categoria)
        {
            var comprador = new ElementHandler.BoundingRectangle(377, 201, 563, 222);
            ClickOn(comprador);

            FillField(categoria);
            PressEnter();
            ScreenPrinter.CaptureAndSaveScreenshot(Global.appSession, screenshotsDirectory + "\\" + Global.app + "\\" + "07-SelectCategoria");
        }

        public void FillAbastecimentoDias(int dias)
        {
            var abastecimentoDiasField = new ElementHandler.BoundingRectangle(431, 338, 460, 357);
            DoubleClickOn(abastecimentoDiasField);
            FillField(dias.ToString());
            ScreenPrinter.CaptureAndSaveScreenshot(Global.appSession, screenshotsDirectory + "\\" + Global.app + "\\" + "08-FillAbastecimentoDias");
        }
    }
}
