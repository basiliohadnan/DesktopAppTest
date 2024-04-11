using DesktopAppTests.Helpers;
using System.Drawing;
using Tesseract;

namespace DesktopAppTests.Helpers
{
    public class OCRTranslator
    {
        public static string ExtractText(string imagePath, int roiX, int roiY, int roiWidth, int roiHeight, int threshold = 150)
        {
            using (var engine = new TesseractEngine(@"C:\Users\Starline\source\repos\DesktopAppTest\tessdata", "eng", EngineMode.Default))
            {
                using (var stream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
                {
                    using (var image = Image.FromStream(stream))
                    {
                        var rect = new Rectangle(roiX, roiY, roiWidth, roiHeight);

                        // Convert the image to Bitmap
                        Bitmap bitmapImage = new Bitmap(image);

                        // Convert the Bitmap image to grayscale
                        var grayscaleImage = ImageEditor.ConvertToGrayscale(bitmapImage);

                        // Apply thresholding
                        var thresholdedImage = ImageEditor.ApplyThreshold(grayscaleImage, threshold);

                        // Save the thresholded image for inspection
                        string thresholdedImagePath = Path.Combine(Path.GetDirectoryName(imagePath), $"thresholded{threshold}_.png");
                        ImageEditor.SaveImage(thresholdedImage, thresholdedImagePath);

                        // Crop the thresholded image
                        using (var croppedImage = ImageEditor.CropImage(thresholdedImage, rect))
                        {
                            // Save the cropped image for inspection
                            string croppedImagePath = Path.Combine(Path.GetDirectoryName(imagePath), $"cropped{threshold}_.png");
                            ImageEditor.SaveImage(croppedImage, croppedImagePath);

                            // Load the cropped image directly into a Pix object
                            using (var memoryStream = new MemoryStream())
                            {
                                croppedImage.Save(memoryStream, System.Drawing.Imaging.ImageFormat.Png);
                                memoryStream.Position = 0;
                                using (var pix = Pix.LoadFromMemory(memoryStream.ToArray()))
                                {
                                    // Perform OCR on the cropped image
                                    using (var page = engine.Process(pix, PageSegMode.Auto))
                                    {
                                        // Return the extracted text
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
}
