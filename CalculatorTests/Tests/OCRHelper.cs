using Tesseract;

namespace CalculatorTests.Tests
{
    public static class OCRHelper
    {
        public static string ReadTextFromImage(string imagePath)
        {
            using (var engine = new TesseractEngine(@"C:\Users\Starline\source\repos\CalculatorTests\tessdata", "eng", EngineMode.Default))
            {
                using (var img = Pix.LoadFromFile(imagePath))
                {
                    using (var page = engine.Process(img))
                    {
                        return page.GetText();
                    }
                }
            }
        }
    }
}
