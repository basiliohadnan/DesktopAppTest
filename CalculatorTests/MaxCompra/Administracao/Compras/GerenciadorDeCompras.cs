using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium.Windows;

namespace Consinco.MaxCompra.Administracao.Compras
{
    [TestClass]
    public class GerenciadorDeCompras : MaxCompraInit
    {
        public GerenciadorDeCompras(WindowsDriver<WindowsElement> appSession) : base(appSession)
        {
        }

        // Inserir POs

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
