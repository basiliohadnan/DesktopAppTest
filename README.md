# OCR Translator

This project provides a utility for extracting text from images using Optical Character Recognition (OCR) technology. It includes functionalities for preprocessing images before performing OCR.

## Description

The OCR Translator project consists of a class named `OCRTranslator` that allows users to extract text from images. The extraction process involves preprocessing the image by resizing, converting to grayscale, and applying thresholding before performing OCR. Additionally, an `ImageEditor` class is provided to handle image manipulation tasks such as cropping, resizing, converting to grayscale, and applying thresholding.

## How to Run

To use the OCR Translator, follow these steps:

1. Clone or download the project from the repository.

2. Open the project in your preferred IDE or text editor.

3. Ensure that you have the necessary dependencies installed. This project relies on the Tesseract OCR engine. Make sure you have Tesseract installed on your system and the necessary language data downloaded.

4. Include the `CalculatorTests.Helpers` namespace in your project or test suite.

5. Use the `OCRTranslator.ExtractText` method to extract text from an image. Provide the path to the image file, along with the region of interest (ROI) coordinates.

6. Optionally, you can customize the preprocessing steps by directly calling methods from the `ImageEditor` class.

Here's an example of how to use the `OCRTranslator`:

```csharp
using CalculatorTests.Helpers;

class Program
{
    static void Main(string[] args)
    {
        string imagePath = "path/to/your/image.png";
        int roiX = 0;
        int roiY = 0;
        int roiWidth = 100;
        int roiHeight = 100;

        string extractedText = OCRTranslator.ExtractText(imagePath, roiX, roiY, roiWidth, roiHeight);
        Console.WriteLine("Extracted Text:");
        Console.WriteLine(extractedText);
    }
}
```

Make sure to replace `"path/to/your/image.png"` with the actual path to your image file, and adjust the ROI coordinates as needed.

## Dependencies

- .NET Framework
- Tesseract OCR Engine

## License

This project is licensed under the [MIT License](LICENSE).
