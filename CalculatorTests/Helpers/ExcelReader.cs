using OfficeOpenXml;

public class ExcelReader
{
    public ExcelWorksheet OpenWorksheet(string filePath, string worksheetName)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Or LicenseContext.Commercial if you have a commercial license

        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            ExcelWorksheet originalWorksheet = package.Workbook.Worksheets[worksheetName];

            if (originalWorksheet == null)
            {
                throw new ArgumentException($"Worksheet '{worksheetName}' not found in the Excel file.");
            }

            var newPackage = new ExcelPackage();
            ExcelWorksheet clonedWorksheet = newPackage.Workbook.Worksheets.Add(worksheetName, originalWorksheet);

            return clonedWorksheet;
        }
    }

    public string ReadCellValue(ExcelWorksheet worksheet, string columnName, int rowNumber)
    {
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
            throw new ArgumentException($"Column '{columnName}' not found in the worksheet.");
        }

        // Read the cell value
        return worksheet.Cells[rowNumber, colIndex].Value?.ToString();
    }
}
