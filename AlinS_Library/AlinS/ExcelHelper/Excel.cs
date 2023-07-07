using OfficeOpenXml;

namespace AlinS.ExcelHelper;

public static class Excel
{
    /// <summary>
    /// Reads all the cells from the specified sheet of the specified Excel file and converts them to string.
    /// </summary>
    /// <param name="xlsxFileName"></param>
    /// <param name="sheetName"> The name of the sheet to read. Leave null to use the first sheet in the workbook. </param>
    public static List<List<string>> Read(string xlsxFileName, string? sheetName = null)
    {
        var data = new List<List<string>>();
        var fileInfo = new FileInfo(xlsxFileName);
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using var package = new ExcelPackage(fileInfo);
        var worksheet = (sheetName is null) ? package.Workbook.Worksheets[0]
                                            : package.Workbook.Worksheets[sheetName];

        int rowCount = worksheet.Dimension.Rows;
        int colCount = worksheet.Dimension.Columns;

        for (int row = 1; row <= rowCount; row++)
        {
            var rowData = new List<string>();

            for (int col = 1; col <= colCount; col++)
            {
                string cellValue = worksheet.Cells[row, col].Value?.ToString() ?? "";
                rowData.Add(cellValue);
            }

            data.Add(rowData);
        }

        return data;
    }

    /// <summary>
    /// Save the specified table to the specified Excel file. If the file already exists, the specified sheet will be overwritten.
    /// </summary>
    /// <param name="xlsxFileName"></param>
    /// <param name="table"></param>
    /// <param name="sheetName"></param>
    public static void Save(string xlsxFileName, List<List<string>> table, string sheetName = "Sheet1")
    {
        var fileInfo = new FileInfo(xlsxFileName);
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using var package = new ExcelPackage(xlsxFileName);

        if (package.Workbook.Worksheets.Any(w => w.Name == sheetName))
            package.Workbook.Worksheets.Delete(sheetName);
        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(sheetName);

        int rowCount = table.Count;
        int colCount = table.Count > 0 ? table[0].Count : 0;

        for (int row = 1; row <= rowCount; row++)
        {
            for (int col = 1; col <= colCount; col++)
                worksheet.Cells[row, col].Value = table[row - 1][col - 1];
        }

        package.SaveAs(fileInfo);
    }
}
