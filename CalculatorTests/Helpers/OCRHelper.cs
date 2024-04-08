//using Tesseract;
//using OpenCvSharp;
//using System.IO;

//namespace CalculatorTests.Helpers
//POC Example 
//{
//    public static class OCRHelper
//    {
//        public static string ReadTextFromImage(string imagePath)
//        {
//            string result = "";

//            // Load the image using OpenCVSharp
//            using (var image = Cv2.ImRead(imagePath))
//            {
//                // Duplicate the size of the image
//                var doubledSizeImage = new Mat();
//                Cv2.Resize(image, doubledSizeImage, new Size(image.Width * 2, image.Height * 2), interpolation: InterpolationFlags.Cubic);

//                // Convert the image to grayscale
//                var grayImage = doubledSizeImage.CvtColor(ColorConversionCodes.BGR2GRAY);

//                // Apply thresholding
//                var thresholdedImage = grayImage.Threshold(150, 255, ThresholdTypes.Binary);

//                // Save the thresholded image to a temporary file with PNG format
//                string tempImagePath = Path.ChangeExtension(Path.GetTempFileName(), ".png");
//                thresholdedImage.SaveImage(tempImagePath);

//                // Perform OCR on the thresholded image
//                using (var engine = new TesseractEngine(@"C:\Users\Starline\source\repos\CalculatorTests\tessdata", "eng", EngineMode.Default))
//                {
//                    using (var img = Pix.LoadFromFile(tempImagePath))
//                    {
//                        using (var page = engine.Process(img))
//                        {
//                            result = page.GetText();
//                        }
//                    }
//                }

//                // Delete the temporary file
//                File.Delete(tempImagePath);
//            }

//            return result;
//        }
//    }
//}
