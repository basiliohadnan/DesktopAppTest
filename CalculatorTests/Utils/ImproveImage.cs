using System.Drawing;
using System;
using OpenCvSharp;

namespace CalculatorTests.Utils
{
    internal class ImproveImage
    {
        string inputFile = "input.jpg";
        string outputFile = "output.jpg";
        int threshold = 150;

            // Load image
            using (var image = new Mat(inputFile))
            {
                // Convert to grayscale
                var gray = image.CvtColor(ColorConversionCodes.BGR2GRAY);

    // Resize image
    var resized = ResizeImage(gray, 2.0);

    // Apply threshold
    var thresholded = ApplyThreshold(resized, threshold);

    // Save processed image
    thresholded.SaveImage(outputFile);
            }

Console.WriteLine("Image processing complete.");
        }

        // Method to resize the image
        static Mat ResizeImage(Mat image, double scaleFactor)
{
    var newSize = new Size((int)(image.Width * scaleFactor), (int)(image.Height * scaleFactor));
    return image.Resize(newSize, interpolation: InterpolationFlags.Cubic);
}

// Method to apply threshold to the image
static Mat ApplyThreshold(Mat image, int thresholdValue)
{
    return image.Threshold(thresholdValue, 255, ThresholdTypes.Binary);
}
    }
}}
    }
}
