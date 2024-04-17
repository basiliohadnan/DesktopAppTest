using Consinco.Helpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Consinco.MaxCompra.Administracao.Compras
{
    [TestClass]
    public class GerenciadorDeCompras : MaxCompraInit
    {

        // Inserir POs

        private ElementHandler elementHandler;

        public GerenciadorDeCompras()
        {
            // Initialize ElementHandler with the app session
            elementHandler = new ElementHandler();
        }

        [TestMethod]
        public void CreateLoteLojaALoja()
        {
            Login();
            OpenMenu("Administração", "02-OpenMenuAdm");
            OpenMenu("Compras", "03-OpenSubMenuCompras");
            OpenMenu("Gerenciador de Compras", "04-OpenSubMenuGerenciadorDeCompras");
            FillFornecedor();
            FillEmpresas();
            WaitSeconds(10);
        }

        public void FillFornecedor()
        {
            var fornecedorField = new ElementHandler.BoundingRectangle(125, 233, 191, 253);
            ClickOn(fornecedorField);
            FillField("478");
        }

        //Lojas
        public void FillEmpresas()
        {
            // tries
            // LegacyIAccessiblePattern.Name	
            // automation id
            // class name
            // name

            //BoundingRectangle	[l=43,t=171,r=120,b=189]
            var empresas = new ElementHandler.BoundingRectangle(43, 171, 120, 189);
            ClickOn(empresas);

            // Adds 11 items from the list
            //int qtyEmpresas = 11;
            //for (int i = 0; i < qtyEmpresas; i++)
            //{
            //    //BoundingRectangle [l=80,t=391,r=207,b=404]
            //    var empresaField = new Helpers.ElementHandler.BoundingRectangle(80, 391, 207, 404);
            //    ClickOn(empresaField);
            //}

            // Click on the first item
            var csVerde = elementHandler.FindElementByClassName("002 Cs Verde");
        }
    }
}
