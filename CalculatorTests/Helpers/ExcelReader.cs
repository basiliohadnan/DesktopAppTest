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

    public string ReadCellValueToString(ExcelWorksheet worksheet, string columnName, int rowNumber)
    {
        // Find the column index by name
        int colIndex = FindColumnIndex(worksheet, columnName);

        // Read the cell value
        return worksheet.Cells[rowNumber, colIndex].Value?.ToString();
    }

    public List<string> ReadCellValueToList(ExcelWorksheet worksheet, string columnName, int rowNumber, char delimiter = ',')
    {
        // Read the cell value as a string
        string cellValue = ReadCellValueToString(worksheet, columnName, rowNumber);

        // Split the cell value using the delimiter
        return SplitCellString(cellValue, delimiter);
    }

    private int FindColumnIndex(ExcelWorksheet worksheet, string columnName)
    {
        for (int col = 1; col <= worksheet.Dimension.Columns; col++)
        {
            if (worksheet.Cells[1, col].Value?.ToString() == columnName)
            {
                return col;
            }
        }

        throw new ArgumentException($"Column '{columnName}' not found in the worksheet.");
    }

    private List<string> SplitCellString(string cellValue, char delimiter)
    {
        List<string> values = new List<string>();

        if (!string.IsNullOrEmpty(cellValue))
        {
            string[] valueArray = cellValue.Split(delimiter);
            foreach (string value in valueArray)
            {
                values.Add(value.Trim()); // Trim to remove any leading or trailing whitespaces
            }
        }

        return values;
    }
}
