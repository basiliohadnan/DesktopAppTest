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
        }

        public void FillFornecedor()
        {
            var fornecedorField = new Helpers.ElementHandler.BoundingRectangle(125,233, 191, 253);
            ClickOn(fornecedorField);
            FillField("478");
        }
    }
}
