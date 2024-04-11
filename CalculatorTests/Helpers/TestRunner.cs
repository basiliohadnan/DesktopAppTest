using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Reflection;

namespace DesktopAppTests.Helpers
{
    public class TestRunner<T> where T : class
    {
        public TestResult RunTests()
        {
            TestResult result = new TestResult();
            TestClassAttribute testClassAttribute = typeof(T).GetCustomAttribute<TestClassAttribute>();

            if (testClassAttribute != null)
            {
                T testClassInstance = Activator.CreateInstance<T>();

                foreach (var methodInfo in typeof(T).GetMethods().Where(m => m.GetCustomAttribute<TestMethodAttribute>() != null))
                {
                    try
                    {
                        methodInfo.Invoke(testClassInstance, null);
                    }
                    catch (Exception ex)
                    {
                        result.Outcome = UnitTestOutcome.Failed;
                        result.TestFailures.Add(new TestFailure { Message = ex.InnerException.Message });
                    }
                }
            }
            else
            {
                result.Outcome = UnitTestOutcome.Failed;
                result.TestFailures.Add(new TestFailure { Message = "Test class not found." });
            }

            return result;
        }
    }

    public class TestResult
    {
        public UnitTestOutcome Outcome { get; set; } = UnitTestOutcome.Passed;
        public List<TestFailure> TestFailures { get; set; } = new List<TestFailure>();
    }

    public class TestFailure
    {
        public string Message { get; set; }
    }
}
