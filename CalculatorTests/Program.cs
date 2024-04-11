//using DesktopAppTests.Tests;
//using DesktopAppTests.Helpers;
//using Microsoft.VisualStudio.TestTools.UnitTesting;

//namespace DesktopAppTests
//{
//    class Program
//    {
//        static void Main(string[] args)
//        {
//            RunNoteAppTests();
//            RunExportTableTests();
//        }

//        static void RunNoteAppTests()
//        {
//            TestRunner<NoteAppTest> testRunner = new TestRunner<NoteAppTest>();
//            Helpers.TestResult result = testRunner.RunTests();
//            DisplayTestResult(result, "NoteApp");
//        }

//        static void RunExportTableTests()
//        {
//            TestRunner<ExportTableTest> testRunner = new TestRunner<ExportTableTest>();
//            Helpers.TestResult result = testRunner.RunTests();
//            DisplayTestResult(result, "ExportTable");
//        }

//        static void DisplayTestResult(Helpers.TestResult result, string testName)
//        {
//            if (result.Outcome == UnitTestOutcome.Passed)
//            {
//                Console.WriteLine($"{testName} tests passed!");
//            }
//            else
//            {
//                Console.WriteLine($"{testName} tests failed:");
//                foreach (var failure in result.TestFailures)
//                {
//                    Console.WriteLine(failure.Message);
//                }
//            }
//        }
//    }
//}
