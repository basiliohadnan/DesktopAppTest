//using System;

//namespace Starline
//{
//    class Program
//    {
//        static void Main(string[] args)
//        {
//            ProcessTest processTest = new ProcessTest();

//            Console.WriteLine("Global Variables");
//            string customerName = "DPSP";
//            string suiteName = "PDV";
//            string scenarioName = "Venda";
//            int reportID = 3;
//            int lgsID;
//            string printFileName;

//            Console.WriteLine("Test Variables");
//            string testName = "Pagamento no dinheiro";
//            string testType = "Funcional";
//            string analystName = "Yuri Souza";
//            string testDesc = "Realizar fluxo de venda simples de 1 produto com pagamento em dinheiro";
//            processTest.StartTest(customerName, suiteName, scenarioName, testName, testType, analystName, testDesc, reportID);

//            Console.WriteLine("Test Details");
//            string preCondition = "Caixa aberto";
//            string postCondition = "Nenhuma";
//            string inputData = "Código de produto válido";
//            processTest.DoTest(preCondition, postCondition, inputData);

//            Console.WriteLine("Steps Definition");
//            processTest.DoStep("Login do gerente", "Login com sucesso");
//            processTest.DoStep("Login do operador", "Login com sucesso");
//            processTest.DoStep("Entrada de CPF fidelidade", "Cadastro encontrado");
//            processTest.DoStep("Entrada de CPF para nota fiscal paulista", "");
//            processTest.DoStep("Entrada de código do produto", "Produto incluí­do no carrinho");
//            processTest.DoStep("Encerramento da venda", "Venda concluída");
//            processTest.DoStep("Tela final", "Caixa disponível para nova operação");

//            Console.WriteLine("Steps Execution");

//            lgsID = processTest.StartStep("Login do gerente", logMsg: "Tentando login", paramName: "cpf", paramValue: "123.456.789-00");
//            // Step actions
//            printFileName = processTest.PrintScreen(full: false, startX: 400, startY: 140, endX: 730, endY: 320);
//            if (true)
//            {
//                processTest.EndStep(lgsID, printPath: printFileName, logMsg: "login efetuado");
//            }
//            else
//            {
//                processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "erro no login");
//            }

//            lgsID = processTest.StartStep("Login do operador", logMsg: "Tentando login", paramName: "cpf", paramValue: "123.456.789-00");
//            // Step actions
//            printFileName = processTest.PrintScreen();
//            if (true)
//            {
//                processTest.EndStep(lgsID, printPath: printFileName, logMsg: "login efetuado");
//            }
//            else
//            {
//                processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "erro no login");
//            }

//            lgsID = processTest.StartStep("Entrada de CPF fidelidade", logMsg: "Tentando entrar cpf do cliente", paramName: "cpf", paramValue: "123.456.789-00");
//            // Step actions
//            printFileName = processTest.PrintScreen();
//            if (true)
//            {
//                processTest.EndStep(lgsID, printPath: printFileName, logMsg: "cpf aceito");
//            }
//            else
//            {
//                processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg: "erro na entrada de cpf fidelidade");
//            }

//            Console.WriteLine("Teardown function");
//            processTest.EndTest(reportID);
//        }
//    }
//}