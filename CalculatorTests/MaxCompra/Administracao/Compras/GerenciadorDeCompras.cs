using Consinco.Helpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Consinco.MaxCompra.Administracao.Compras
{
    [TestClass]
    public class GerenciadorDeCompras : MaxCompraInit
    {
        // Inserir POs
        [TestMethod]
        public void CreateLoteLojaALoja()
        {
            Login();
            OpenMenu("Administração", "02-OpenMenuAdm");
            OpenMenu("Compras", "03-OpenSubMenuCompras");
            OpenMenu("Gerenciador de Compras", "04-OpenSubMenuGerenciadorDeCompras");
            FillFornecedor();
            SelectLojas();
        }

        public void FillFornecedor()
        {
            var fornecedorField = new ElementHandler.BoundingRectangle(125, 233, 191, 253);
            ClickOn(fornecedorField);
            FillField("478");
            ScreenPrinter.CaptureAndSaveScreenshot(Global.appSession, screenshotsDirectory + "\\" + Global.app + "\\" + "05-FillFornecedor");
        }

        // Lojas
        public void SelectLojas()
        {
            var empresas = new ElementHandler.BoundingRectangle(43, 171, 120, 189);
            ClickOn(empresas);

            // Adds 11 first items from the list of Lojas
            int qtyEmpresas = 11;
            var empresaField = new ElementHandler.BoundingRectangle(80, 391, 207, 404);
            for (int i = 0; i < qtyEmpresas; i++)
            {
                DoubleClickOn(empresaField);
            }

            var confirmButton = new ElementHandler.BoundingRectangle(83, 89, 111, 117);
            ClickOn(confirmButton);
            ScreenPrinter.CaptureAndSaveScreenshot(Global.appSession, screenshotsDirectory + "\\" + Global.app + "\\" + "06-SelectLojas");
        }
    }
}
