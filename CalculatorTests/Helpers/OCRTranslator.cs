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
            using (var engine = new TesseractEngine(@"C:\Users\Starline\source\repos\CalculatorTests\tessdata", "eng", EngineMode.Default))
            {
                using (var stream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
                {
                    using (var image = Image.FromStream(stream))
                    {
                        var rect = new Rectangle(roiX, roiY, roiWidth, roiHeight);

                        // Crop the original image
                        using (var croppedImage = CropImage(image, rect))
                        {
                            // Duplicate the size of the cropped image
                            var enlargedImage = EnlargeImage(croppedImage);

                            // Convert the image to grayscale
                            var grayscaleImage = ConvertToGrayscale(enlargedImage);

                            // Apply thresholding (Threshold = 150)
                            var thresholdedImage = ApplyThreshold(grayscaleImage, 150);

                            // Save the pre-processed image for inspection
                            string preprocessedImagePath = Path.Combine(Path.GetDirectoryName(imagePath), "preprocessed_image.png");
                            thresholdedImage.Save(preprocessedImagePath, System.Drawing.Imaging.ImageFormat.Png);
                            Console.WriteLine($"Preprocessed image saved to: {preprocessedImagePath}");

                            // Load the pre-processed image directly into a Pix object
                            using (var memoryStream = new MemoryStream())
                            {
                                thresholdedImage.Save(memoryStream, System.Drawing.Imaging.ImageFormat.Png);
                                memoryStream.Position = 0;
                                using (var pix = Pix.LoadFromMemory(memoryStream.ToArray()))
                                {
                                    // Perform OCR on the pre-processed image
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

        private static Bitmap CropImage(Image image, Rectangle cropRect)
        {
            var croppedImage = new Bitmap(cropRect.Width, cropRect.Height);
            using (var graphics = Graphics.FromImage(croppedImage))
            {
                graphics.DrawImage(image, new Rectangle(0, 0, cropRect.Width, cropRect.Height), cropRect, GraphicsUnit.Pixel);
            }
            return croppedImage;
        }

        private static Bitmap EnlargeImage(Bitmap image)
        {
            var enlargedImage = new Bitmap(image.Width * 2, image.Height * 2);
            using (var graphics = Graphics.FromImage(enlargedImage))
            {
                graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
                graphics.DrawImage(image, 0, 0, enlargedImage.Width, enlargedImage.Height);
            }
            return enlargedImage;
        }

        private static Bitmap ConvertToGrayscale(Bitmap image)
        {
            var grayscaleImage = new Bitmap(image.Width, image.Height, PixelFormat.Format24bppRgb);
            using (var graphics = Graphics.FromImage(grayscaleImage))
            {
                var colorMatrix = new ColorMatrix(new float[][]
                {
                    new float[] {0.299f, 0.299f, 0.299f, 0, 0},
                    new float[] {0.587f, 0.587f, 0.587f, 0, 0},
                    new float[] {0.114f, 0.114f, 0.114f, 0, 0},
                    new float[] {0, 0, 0, 1, 0},
                    new float[] {0, 0, 0, 0, 1}
                });
                using (var attributes = new ImageAttributes())
                {
                    attributes.SetColorMatrix(colorMatrix);
                    graphics.DrawImage(image, new Rectangle(0, 0, grayscaleImage.Width, grayscaleImage.Height), 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, attributes);
                }
            }
            return grayscaleImage;
        }

        private static Bitmap ApplyThreshold(Bitmap image, int threshold)
        {
            var thresholdedImage = new Bitmap(image.Width, image.Height);
            for (int x = 0; x < image.Width; x++)
            {
                for (int y = 0; y < image.Height; y++)
                {
                    Color originalColor = image.GetPixel(x, y);
                    Color thresholdedColor = Color.FromArgb(originalColor.R > threshold ? 255 : 0,
                                                             originalColor.G > threshold ? 255 : 0,
                                                             originalColor.B > threshold ? 255 : 0);
                    thresholdedImage.SetPixel(x, y, thresholdedColor);
                }
            }
            return thresholdedImage;
        }
    }
}
