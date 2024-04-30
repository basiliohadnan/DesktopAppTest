using OfficeOpenXml;

public class ExcelReader
{
    private string filePath;

    public ExcelReader(string filePath)
    {
        this.filePath = filePath;
    }

    public string ReadCellValue(string columnName, int rowNumber)
    {
        // Set LicenseContext to suppress the LicenseException
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Or LicenseContext.Commercial if you have a commercial license

        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Assuming data is in the first worksheet

            // Find the column index by name
            int colIndex = -1;
            for (int col = 1; col <= worksheet.Dimension.Columns; col++)
            {
                if (worksheet.Cells[1, col].Value?.ToString() == columnName)
                {
                    colIndex = col;
                    break;
                }
            }

            if (colIndex == -1)
            {
                throw new ArgumentException($"Column '{columnName}' not found in the Excel file.");
            }

            // Read the cell value
            return worksheet.Cells[rowNumber, colIndex].Value?.ToString();
        }
    }
}
