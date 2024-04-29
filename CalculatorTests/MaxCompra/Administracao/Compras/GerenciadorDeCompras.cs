using Consinco.Helpers;
using Consinco.MaxCompra.PageObjects.Administracao.Compras;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium.Windows;

namespace Consinco.MaxCompra.Administracao.Compras
{
    [TestClass]
    public class GerenciadorDeComprasTests : MaxCompraInit
    {
        private GerenciadorDeComprasPO gerenciadorDeComprasPO;
        private OCRScanner scan = new OCRScanner();

        // 5 lojas e 4 produtos

        [TestMethod]
        public void CriarLoteDeCompraLojaALojaUmaLoja()
        {
            //Console.WriteLine("Global Variables");
            string scenarioName = "Gerenciador de Compras";
            int reportID = 3;
            int lgsID;
            string printFileName;
            int qtdLojas = 1;
            int qtdeCompra = 6;
            int qtdProdutos = 1;

            //Console.WriteLine("Test Variables");
            string welcomeWindowName = "Conexão de Sistemas Consinco";
            string testName = "Criar lote de compras - Loja a loja";
            string testType = "Funcional";
            string analystName = "Hadnan Basilio";
            string testDesc = "Criar lote de compras - Loja a loja";
            Global.processTest.StartTest(Global.customerName, suiteName, scenarioName, testName, testType, analystName, testDesc, reportID);

            //Console.WriteLine("Test Details");
            string preCondition = "App iniciando";
            string postCondition = "Lote criado com sucesso";
            string inputData = "Nenhum";
            Global.processTest.DoTest(preCondition, postCondition, inputData);

            //Console.WriteLine("Steps Definition");
            Global.processTest.DoStep("Abrir app", "Abertura do app com sucesso");
            Global.processTest.DoStep("Login do analista", "Login com sucesso");
            Global.processTest.DoStep("Tela final", "Tela principal exibida com sucesso");
            Global.processTest.DoStep("Abrir menu Administração", "Menu Administracao aberto com sucesso");
            Global.processTest.DoStep("Abrir menu Compras", "Menu Compras aberto com sucesso");
            Global.processTest.DoStep("Abrir menu Gerenciador de Compras", "Gerenciador de Compras aberto com sucesso");
            Global.processTest.DoStep("Preencher fornecedor", "Fornecedor preenchido com sucesso");
            //Global.processTest.DoStep("Adicionar lojas", "Adição de lojas com sucesso");
            //Global.processTest.DoStep("Confirmar janela Seleção de Empresas do Lote", "Confirmação janela Seleção de Empresas do Lote com sucesso");
            Global.processTest.DoStep("Selecionar categoria", "Seleção da categoria com sucesso");
            Global.processTest.DoStep("Preencher dias abastecimento", "Preenchimento dias abastecimento com sucesso");
            Global.processTest.DoStep("Habilitar checkboxes Sugestão de compra", "Habilitação checkboxes sugestão de compra com sucesso");
            Global.processTest.DoStep("Incluir lote de compra", "Inclusão do lote de compra com sucesso");
            Global.processTest.DoStep("Adicionar comprador", "Inclusão de comprador com sucesso");
            Global.processTest.DoStep("Confirmar janela Seleção de Produtos", "Confirmação janela Seleção de Produtos com sucesso");
            //Global.processTest.DoStep("Confirmar janela Produtos Inativos", "Confirmação janela Produtos Inativos com sucesso");
            //Global.processTest.DoStep("Confirmar janela Tributação", "Confirmação janela Tributação com sucesso");
            Global.processTest.DoStep("Preencher quantidade de compra por loja e produto",
                "Preenchimento quantidade de compra por loja e produto com sucesso");
            Global.processTest.DoStep("Gerar Pedidos", "Pedidos gerados com sucesso");
            Global.processTest.DoStep("Confirmar Consulta do lote", "Lote gerado com sucesso");

            //Console.WriteLine("Steps Execution");
            lgsID = Global.processTest.StartStep("Abrir app", logMsg: "Tentando abrir app", paramName: "appPath", paramValue: appPath);
            Initialize();
            printFileName = Global.processTest.PrintScreen();

            WindowsElement welcomeWindow = elementHandler.FindElementByName(welcomeWindowName);
            if (welcomeWindow != null)
            {
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: "app aberto");
            }
            else
            {
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "erro na abertura do app");
            }

            lgsID = Global.processTest.StartStep("Login do analista", logMsg: "Tentando login", paramName: "matricula", paramValue: matricula);
            Login();
            SetAppSession();
            printFileName = Global.processTest.PrintScreen();

            string databaseWarningName = "Não foi definido a versão do módulo no BANCO DE DADOS, para o sistema de Segurança";
            WindowsElement databaseWarning = elementHandler.FindElementByXPathPartialName(databaseWarningName);
            if (databaseWarning != null)
            {
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: "login efetuado");
            }
            else
            {
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "erro no login");
            }

            lgsID = Global.processTest.StartStep("Tela final", logMsg: "Tentando acessar tela principal", paramName: "", paramValue: "");
            databaseWarning.Click();
            PressEnter();

            string mainWindowClassName = "Centura:MDIFrame";
            WindowsElement mainWindow = elementHandler.FindElementByClassName(mainWindowClassName);
            printFileName = Global.processTest.PrintScreen();
            if (mainWindow != null)
            {
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: "tela principal exibida");
            }
            else
            {
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "erro na exibição da tela principal");
            }

            string menuName = "Administração";
            lgsID = Global.processTest.StartStep($"Abrir menu {menuName}", logMsg: $"Abrir menu {menuName}",
                paramName: "menuName", paramValue: menuName);
            try
            {
                OpenMenu(menuName);
                printFileName = Global.processTest.PrintScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"menu {menuName} aberto com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.PrintScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"erro na abertura do menu {menuName}");
            }

            menuName = "Compras";
            lgsID = Global.processTest.StartStep($"Abrir menu {menuName}", logMsg: $"Abrir menu {menuName}", paramName: "menuName",
                paramValue: menuName);
            try
            {
                OpenMenu(menuName);
                printFileName = Global.processTest.PrintScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"menu {menuName} aberto com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.PrintScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"erro na abertura do menu {menuName}");
            }

            menuName = "Gerenciador de Compras";
            lgsID = Global.processTest.StartStep($"Abrir menu {menuName}", logMsg: $"Abrir menu {menuName}", paramName: "menuName",
                paramValue: menuName);
            try
            {
                OpenMenu(menuName);
                WaitSeconds(1);
                int startX = 2;
                int startY = 45;
                int endX = 1023;
                int endY = 785;
                printFileName = Global.processTest.PrintScreen(fullScreen: false, startX: startX, startY: startY, endX: endX, endY: endY);
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"menu {menuName} aberto com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.PrintScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"erro na abertura do menu {menuName}");
            }

            gerenciadorDeComprasPO = new GerenciadorDeComprasPO(new ElementHandler());
            string codFornecedor = "478";
            lgsID = Global.processTest.StartStep("Preencher fornecedor", logMsg: $"Preencher fornecedor {codFornecedor}",
                paramName: "fornecedorId", paramValue: codFornecedor);
            try
            {
                gerenciadorDeComprasPO.FillFornecedor(codFornecedor);
                int startX = 2;
                int startY = 45;
                int endX = 1023;
                int endY = 785;
                printFileName = Global.processTest.PrintScreen(fullScreen: false, startX: startX, startY: startY, endX: endX, endY: endY);
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: "Fornecedor preenchido com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.PrintScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "erro no preenchimento do fornecedor");
            }

            //lgsID = Global.processTest.StartStep("Adicionar lojas", logMsg: $"Adicionar {qtdLojas} lojas", paramName: "qtdLojas",
            //    paramValue: qtdLojas.ToString());
            //try
            //{
            //    gerenciadorDeComprasPO.AddLojas(qtdLojas);
            //    int startX = 50;
            //    int startY = 60;
            //    int endX = 470;
            //    int endY = 591;
            //    printFileName = Global.processTest.PrintScreen(fullScreen: false, startX: startX, startY: startY, endX: endX, endY: endY);
            //    Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: "Adição de lojas com sucesso");
            //}
            //catch
            //{
            //                printFileName = Global.processTest.PrintScreen();
            //    Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "erro na adição de lojas");
            //}

            //string windowName = "Seleção de Empresas do Lote";
            //lgsID = Global.processTest.StartStep($"Confirmar janela {windowName}", logMsg: $"Tentando confirmar janela {windowName}",
            //    paramName: "windowName", paramValue: windowName);
            //try
            //{
            //    gerenciadorDeComprasPO.ConfirmSelecaoDeLojasWindow();
            //    // Verificar motivo estar salvando screenshot após seleção da categoria.
            //    //int startX = 50;
            //    //int startY = 60;
            //    //int endX = 470;
            //    //int endY = 591;
            //    //printFileName = Global.processTest.PrintScreen(fullScreen: false, startX: startX, startY: startY, endX: endX, endY: endY);
            //    printFileName = Global.processTest.PrintScreen();
            //    Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Confirmação janela {windowName} com sucesso");
            //}
            //catch
            //{
            //                printFileName = Global.processTest.PrintScreen();
            //    Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"erro na confirmação janela {windowName}");
            //}

            string categoria = "LIQ2";
            lgsID = Global.processTest.StartStep("Selecionar categoria", logMsg: $"Selecionar categoria {categoria}",
                paramName: "categoria", paramValue: categoria);
            try
            {
                gerenciadorDeComprasPO.SelectCategoria(categoria);
                int startX = 2;
                int startY = 45;
                int endX = 1023;
                int endY = 785;
                printFileName = Global.processTest.PrintScreen(fullScreen: false, startX: startX, startY: startY, endX: endX, endY: endY);
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: "Seleção da categoria com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.PrintScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "erro na seleção da categoria");
            }


            string diasAbastecimento = "60";
            lgsID = Global.processTest.StartStep("Preencher dias abastecimento", logMsg: $"Preencher dias abastecimento com {diasAbastecimento}",
                paramName: "diasAbastecimento", paramValue: diasAbastecimento);
            try
            {
                gerenciadorDeComprasPO.FillAbastecimentoDias(diasAbastecimento);
                int startX = 2;
                int startY = 45;
                int endX = 1023;
                int endY = 785;
                printFileName = Global.processTest.PrintScreen(fullScreen: false, startX: startX, startY: startY, endX: endX, endY: endY);
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: "Preenchimento dias abastecimento com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.PrintScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "erro no preenchimento de dias abastecimento");
            }

            string checkBoxesClass = "Centura:GPCheck";
            lgsID = Global.processTest.StartStep("Habilitar checkboxes Sugestão de compra", logMsg: $"Habilitar checkboxes sugestão de compras",
                paramName: "checkBoxesClass", paramValue: checkBoxesClass);
            try
            {
                gerenciadorDeComprasPO.EnableCheckBoxesSugestaoDeCompras(checkBoxesClass);
                // Verificar motivo estar salvando screenshot após incluir lote.
                int startX = 2;
                int startY = 45;
                int endX = 1023;
                int endY = 785;
                printFileName = Global.processTest.PrintScreen(fullScreen: false, startX: startX, startY: startY, endX: endX, endY: endY);
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: "Habilitação checkboxes sugestão de compra com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.PrintScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName,
                    logMsg: "erro na habilitação dos checkboxes de Sugestão de compra");
            }

            lgsID = Global.processTest.StartStep("Incluir lote de compra", logMsg: $"Incluir lote de compra", paramName: "", paramValue: "");
            try
            {
                gerenciadorDeComprasPO.IncluirLote();
                int startX = 2;
                int startY = 45;
                int endX = 1023;
                int endY = 785;
                printFileName = Global.processTest.PrintScreen(fullScreen: false, startX: startX, startY: startY, endX: endX, endY: endY);
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: "Inclusão do lote de compra com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.PrintScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "erro na inclusão do lote de compra");
            }

            string comprador = "LIQ1";
            lgsID = Global.processTest.StartStep("Adicionar comprador", logMsg: $"Adicionar comprador {comprador}", paramName: "comprador",
                paramValue: comprador);
            try
            {
                gerenciadorDeComprasPO.AddCompradores(comprador);
                int startX = 24;
                int startY = 77;
                int endX = 631;
                int endY = 500;
                printFileName = Global.processTest.PrintScreen(fullScreen: false, startX: startX, startY: startY, endX: endX, endY: endY);
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: "Inclusão de comprador com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.PrintScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "erro na inclusão de comprador");
            }

            string windowName = "Seleção de Produtos";
            lgsID = Global.processTest.StartStep($"Confirmar janela {windowName}", logMsg: $"Tentando confirmar janela {windowName}",
                paramName: "windowName", paramValue: windowName);
            try
            {
                gerenciadorDeComprasPO.ConfirmSelecaoDeProdutosWindow();
                //string produtosInativosWindowName = "Produtos Inativos";
                //WindowsElement foundWindow = elementHandler.FindElementByName(produtosInativosWindowName, 120 * 1000 / 10);
                printFileName = Global.processTest.PrintScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Confirmação janela {windowName} com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.PrintScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"erro na confirmação janela {windowName}");
            }

            //windowName = "Produtos Inativos";
            //lgsID = Global.processTest.StartStep($"Confirmar janela {windowName}", logMsg: $"Tentando confirmar janela {windowName}",
            //    paramName: "windowName", paramValue: windowName);
            //try
            //{
            //    gerenciadorDeComprasPO.ConfirmProdutosInativosWindow();
            //    //printFileName = Global.processTest.PrintScreen();
            //    int startX = 53;
            //    int startY = 63;
            //    int endX = 677;
            //    int endY = 86;
            //    printFileName = Global.processTest.PrintScreen(fullScreen: false, startX: startX, startY: startY, endX: endX, endY: endY);
            //    Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Confirmação janela {windowName} com sucesso");
            //}
            //catch
            //{
            //  printFileName = Global.processTest.PrintScreen();
            //    Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"erro na confirmação janela {windowName}");
            //}

            //windowName = "Tributação";
            //lgsID = Global.processTest.StartStep($"Confirmar janela {windowName}", logMsg: $"Tentando confirmar janela {windowName}",
            //    paramName: "windowName", paramValue: windowName);
            //try
            //{
            //    gerenciadorDeComprasPO.ConfirmTributacaoWindow();
            //    int startX = 2;
            //    int startY = 45;
            //    int endX = 1023;
            //    int endY = 785;
            //    printFileName = Global.processTest.PrintScreen(fullScreen: false, startX: startX, startY: startY, endX: endX, endY: endY);
            //    Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Confirmação janela {windowName} com sucesso");
            //}
            //catch
            //{
            //    printFileName = Global.processTest.PrintScreen();
            //    Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: $"erro na confirmação janela {windowName}");
            //}

            lgsID = Global.processTest.StartStep($"Preencher quantidade de compra por loja e produto",
                logMsg: $"Preencher quantidade de compra por loja e produto",
                paramName: "qtdLojas, qtdProdutos, qtdeCompra", paramValue: $"{qtdLojas}, {qtdProdutos}, {qtdeCompra}");
            try
            {
                gerenciadorDeComprasPO.FillQtdeCompra(qtdProdutos, qtdeCompra);

                //Validação total de compra
                string qtdeComprValue = gerenciadorDeComprasPO.GetQtdeComprValue();
                int qtdeComprasValue = int.Parse(qtdeComprValue);
                if (qtdeComprasValue != qtdLojas * qtdProdutos * qtdeCompra)
                {
                    Console.Write($"Erro no preenchimento: qtdeComprasValue atual: {qtdeComprasValue}, Total esperado: {qtdLojas * qtdProdutos * qtdeCompra}");
                }
                //Maximize();
                printFileName = Global.processTest.PrintScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Preenchimento quantidade de compra por loja e produto com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.PrintScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName,
                    logMsg: $"erro no preenchimento quantidade de compra por loja e produto");
            }

            lgsID = Global.processTest.StartStep($"Gerar Pedidos", logMsg: $"Tentando Gerar Pedidos", paramName: "", paramValue: "");
            try
            {
                gerenciadorDeComprasPO.ClickGerarPedidos();
                gerenciadorDeComprasPO.ConfirmPedidosWindow();
                printFileName = Global.processTest.PrintScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Pedidos gerados com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.PrintScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName,
                    logMsg: $"erro ao clicar no botão Gera Pedidos");
            }

            lgsID = Global.processTest.StartStep($"Confirmar Consulta do lote", logMsg: $"Tentando confirmar Consulta do Lote", paramName: "", paramValue: "");
            try
            {
                gerenciadorDeComprasPO.ConfirmConsultaLoteCompraWindow();
                printFileName = Global.processTest.PrintScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"Lote gerado com sucesso");
            }
            catch
            {
                printFileName = Global.processTest.PrintScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName,
                    logMsg: $"erro ao clicar no confirmar Lote");
            }
            //Teardown function
            PressEnter();
            Global.processTest.EndTest(reportID);
        }
    }
}