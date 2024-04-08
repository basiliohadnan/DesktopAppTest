using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Tesseract;

namespace CalculatorTests.Helpers
{
    public class OCRTranslator
    {
        public static string PerformOCR(string imagePath)
        {
            using (var engine = new TesseractEngine(@"C:\Users\Starline\source\repos\CalculatorTests\tessdata", "eng", EngineMode.Default))
            {
                using (var image = Pix.LoadFromFile(imagePath))
                {
                    using (var page = engine.Process(image))
                    {
                        return page.GetText().Trim();
                    }
                }
            }
        }

        public static bool ValidateResult(string screenshotPath, string expectedValue)
        {
            string result = PerformOCR(screenshotPath);
            return result == expectedValue;
        }
    }
}
