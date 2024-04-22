using Consinco.Helpers;
using Consinco.MaxCompra;
using DesktopAppTests.MaxCompra.PageObjects.Administracao.Compras;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace DesktopAppTests.MaxCompra.Administracao.Compras
{
    [TestClass]
    public class GerenciadorDeComprasTests : MaxCompraInit
    {
        private GerenciadorDeComprasPO gerenciadorDeComprasPO;

        [TestMethod]
        public void CreateLoteLojaALoja()
        {
            Login();
            OpenMenu("Administração", "02-OpenMenuAdm");
            OpenMenu("Compras", "03-OpenSubMenuCompras");
            OpenMenu("Gerenciador de Compras", "04-OpenSubMenuGerenciadorDeCompras");

            gerenciadorDeComprasPO = new GerenciadorDeComprasPO(new ElementHandler());
            gerenciadorDeComprasPO.FillFornecedor(478);
            gerenciadorDeComprasPO.SelectLojas(2);
            gerenciadorDeComprasPO.SelectCategoria("LIQ2");
            gerenciadorDeComprasPO.FillAbastecimentoDias(60);
            gerenciadorDeComprasPO.EnableCheckBoxesSugestaoDeCompras("Centura:GPCheck");
            gerenciadorDeComprasPO.IncluirLote();
            gerenciadorDeComprasPO.AddCompradores("LIQ1");
            gerenciadorDeComprasPO.ConfirmProdutosInativosWindow();
            gerenciadorDeComprasPO.ConfirmTributacaoWindow();
            //SetAppSession("Centura:MDIFrame");

        }
    }
}