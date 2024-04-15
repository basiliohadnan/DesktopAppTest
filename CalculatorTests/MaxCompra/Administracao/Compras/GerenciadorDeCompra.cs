using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;

namespace Consinco.MaxCompra.Administracao.Compras
{
    [TestClass]
    public class GerenciadorDeCompra : MaxCompraInit
    {
        public GerenciadorDeCompra(IWebDriver driver) : base(driver)
        {
        }

        //POs


        [TestMethod]
        public void CriarLoteCDParaLoja()
        {
            Login();
            OpenMenuItem("Administração", "02-OpenMenuAdm");
            OpenMenuItem("Compras", "03-OpenSubMenuCompras");
            OpenMenuItem("Gerenciador de Compras", "04-OpenSubMenuGerenciadorDeCompras");
        }
    }
}
