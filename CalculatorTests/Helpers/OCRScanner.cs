using OpenCvSharp;
using System.Drawing;
using Tesseract;

namespace Consinco.Helpers
{
    public class OCRScanner
    {
        public static string tessPath = $"C:\\Users\\{Global.user}\\source\\repos\\DesktopAppTest\\tessdata";
        public static string ExtractText(string imagePath, int roiX, int roiY, int roiWidth, int roiHeight, int threshold = 150)
        {
            using (var engine = new TesseractEngine(tessPath, "eng", EngineMode.Default))
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
        public string ValidateField(string inputFile, string outputFile, int threshold = 150, int ratio = 2)
        {
            string value = "";
            Image src = Image.FromFile(inputFile);
            using (var image = new Mat(inputFile))
            // bw
            using (var gray = image.CvtColor(ColorConversionCodes.BGR2GRAY))
            {
                // resize
                var faceSize = new OpenCvSharp.Size(src.Width * ratio, src.Height * ratio);
                var resized = gray.Resize(faceSize, 0, 0, InterpolationFlags.Cubic);
                // threshold
                var thre = resized.Threshold(threshold, 255, ThresholdTypes.Binary);
                // check
                if (File.Exists(outputFile))
                {
                    File.Delete(outputFile);
                }
                thre.SaveImage(outputFile);
            }

            try
            {
                using (var engine = new TesseractEngine(tessPath, "eng", EngineMode.Default))
                {
                    using (var img = Pix.LoadFromFile(outputFile))
                    {
                        using (var page = engine.Process(img))
                        {
                            value = page.GetText().Replace("\n", "");
                        }
                    }
                }
            }
            catch (Exception ex)
            {

                throw ex;
            }
            return value;
        }
    }
}
