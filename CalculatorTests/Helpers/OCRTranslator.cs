using CalculatorTests.Helpers;
using System.Drawing;
using System.Drawing.Imaging;

public class OCRTranslator
{
    public static string ExtractText(string imagePath, int roiX, int roiY, int roiWidth, int roiHeight)
    {
        using (var engine = new Tesseract.TesseractEngine(@"C:\Users\Starline\source\repos\CalculatorTests\tessdata", "eng", Tesseract.EngineMode.Default))
        {
            using (var stream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
            {
                using (var image = Image.FromStream(stream))
                {
                    var rect = new Rectangle(roiX, roiY, roiWidth, roiHeight);

                    // Duplicate the size of the image
                    var enlargedImage = ImageEditor.DuplicateSize(image);

                    // Convert the image to grayscale
                    var grayscaleImage = ImageEditor.ConvertToGrayscale(enlargedImage);

                    // Apply thresholding
                    var thresholdedImage = ImageEditor.ApplyThreshold(grayscaleImage, 175);

                    // Save the pre-processed image for inspection
                    string preprocessedImagePath = Path.Combine(Path.GetDirectoryName(imagePath), "preprocessed_image.png");
                    ImageEditor.SaveImage(thresholdedImage, preprocessedImagePath);

                    // Crop the preprocessed image
                    using (var croppedImage = ImageEditor.CropImage(thresholdedImage, rect))
                    {
                        // Save the cropped image for inspection
                        string croppedImagePath = Path.Combine(Path.GetDirectoryName(imagePath), "cropped_image.png");
                        ImageEditor.SaveImage(croppedImage, croppedImagePath);

                        // Load the cropped image directly into a Pix object
                        using (var memoryStream = new MemoryStream())
                        {
                            croppedImage.Save(memoryStream, ImageFormat.Png);
                            memoryStream.Position = 0;
                            using (var pix = Tesseract.Pix.LoadFromMemory(memoryStream.ToArray()))
                            {
                                // Perform OCR on the cropped image
                                using (var page = engine.Process(pix, Tesseract.PageSegMode.Auto))
                                {
                                    // Return the extracted text
                                    var teste = page.GetText();
                                    return page.GetText().Trim();
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
