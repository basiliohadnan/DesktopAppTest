using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Tesseract;

namespace CalculatorTests.Helpers
{
    public class OCRTranslator
    {
        public static string ExtractText(string imagePath, int roiX, int roiY, int roiWidth, int roiHeight)
        {
            // Load the full screenshot image
            using (var fullImage = new Bitmap(imagePath))
            {
                // Define the ROI (region of interest) rectangle
                var roiRect = new Rectangle(roiX, roiY, roiWidth, roiHeight);

                // Crop the full image to get the ROI
                using (var croppedImage = fullImage.Clone(roiRect, fullImage.PixelFormat))
                {
                    // Convert the cropped image to a Pix object
                    using (var memoryStream = new MemoryStream())
                    {
                        croppedImage.Save(memoryStream, System.Drawing.Imaging.ImageFormat.Png);
                        memoryStream.Position = 0; // Reset the position

                        // Perform OCR on the cropped image
                        using (var engine = new TesseractEngine(@"C:\Users\Starline\source\repos\CalculatorTests\tessdata", "eng", EngineMode.Default))
                        using (var pix = Pix.LoadFromMemory(memoryStream.ToArray()))
                        using (var page = engine.Process(pix, PageSegMode.Auto))
                        {
                            return page.GetText().Trim();
                        }
                    }
                }
            }
        }
    }
}
