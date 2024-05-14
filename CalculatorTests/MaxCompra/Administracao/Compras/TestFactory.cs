using Consinco.Helpers;
using System;
using System.Collections.Generic;
using System.Reflection;

namespace Consinco.MaxCompra.Administracao.Compras
{
    public class TestFactory
    {
        private Dictionary<string, Action<Dictionary<string, string>>> testMethods;

        public TestFactory(Dictionary<string, Action<Dictionary<string, string>>> methods)
        {
            testMethods = methods ?? throw new ArgumentNullException(nameof(methods), "Dictionary of methods cannot be null.");
        }

        public void ExecuteTestStep(string methodName, Dictionary<string, string> parameters)
        {
            // Check if the method exists in the dictionary
            if (testMethods.ContainsKey(methodName))
            {
                // Get the corresponding delegate and execute it with the provided parameters
                Action<Dictionary<string, string>> method = testMethods[methodName];
                ExecuteTestMethod(methodName, method, parameters);
            }
            else
            {
                // Method not found in dictionary
                Console.WriteLine($"Method '{methodName}' not found in the dictionary.");
            }
        }

        private void ExecuteTestMethod(string methodName, Action<Dictionary<string, string>> method, Dictionary<string, string> parameters)
        {
            string printFileName;
            int lgsID = Global.processTest.StartStep($"Executing {methodName}",
                logMsg: $"Executing {methodName} with parameters: {string.Join(", ", parameters.Select(kv => $"{kv.Key}={kv.Value}"))}",
                paramName: "parameters", paramValue: string.Join(", ", parameters.Select(kv => $"{kv.Key}={kv.Value}")));
            try
            {
                method(parameters);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: $"{methodName} executed successfully.");
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "error", printPath: printFileName, logMsg:
                    $"Error occurred while executing {methodName}.");
            }
        }

        public static string GetCurrentMethodName()
        {
            MethodBase method = new System.Diagnostics.StackTrace().GetFrame(1).GetMethod();
            return method.Name;
        }
    }
}