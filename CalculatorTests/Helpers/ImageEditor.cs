﻿using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace CalculatorTests.Helpers
{
    public class ImageEditor
    {
        public static Bitmap CropImage(Bitmap image, Rectangle cropRect)
        {
            var croppedImage = new Bitmap(cropRect.Width, cropRect.Height);
            using (var graphics = Graphics.FromImage(croppedImage))
            {
                graphics.DrawImage(image, new Rectangle(0, 0, cropRect.Width, cropRect.Height), cropRect, GraphicsUnit.Pixel);
            }
            return croppedImage;
        }

        public static Bitmap DuplicateSize(Image image)
        {
            var enlargedImage = new Bitmap(image.Width * 2, image.Height * 2);
            using (var graphics = Graphics.FromImage(enlargedImage))
            {
                graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
                graphics.DrawImage(image, 0, 0, enlargedImage.Width, enlargedImage.Height);
            }
            return enlargedImage;
        }

        public static Bitmap ConvertToGrayscale(Bitmap image)
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

        public static Bitmap ApplyThreshold(Bitmap image, int threshold)
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

        public static void SaveImage(Bitmap image, string imagePath)
        {
            image.Save(imagePath, ImageFormat.Png);
            Console.WriteLine($"Image saved to: {imagePath}");
        }
    }
}