using Consinco.Helpers;
using Consinco.MaxCompra;
using DesktopAppTests.MaxCompra.PageObjects.Administracao.Compras;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace DesktopAppTests.MaxCompra.Administracao.Compras
{
    [TestClass]
    public class GerenciadorDeComprasTests : MaxCompraInit
    {
        private GerenciadorDeComprasPO gerenciadorPO;

        public GerenciadorDeComprasTests()
        {
            gerenciadorPO = new GerenciadorDeComprasPO(new ElementHandler());
        }

        [TestMethod]
        public void CreateLoteLojaALoja()
        {
            Login();
            OpenMenu("Administração", "02-OpenMenuAdm");
            OpenMenu("Compras", "03-OpenSubMenuCompras");
            OpenMenu("Gerenciador de Compras", "04-OpenSubMenuGerenciadorDeCompras");

            gerenciadorPO.FillFornecedor(478);
            gerenciadorPO.SelectLojas(11);
            gerenciadorPO.SelectCategoria("LIQ2 (SUCOS, AG");
            gerenciadorPO.FillAbastecimentoDias(60);
            gerenciadorPO.EnableCheckBoxesSugestaoDeCompras("Centura:GPCheck");
            gerenciadorPO.IncluirLote();
            gerenciadorPO.AddCompradores("LIQ1");
        }
    }
}