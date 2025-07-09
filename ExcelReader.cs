// Logic to read from Excel
using System;
using System.Data;
using System.IO;
using OfficeOpenXml;

namespace UserStorySimilarityAddIn
{
    public static class ExcelReader
    {
        public static DataTable ReadExcelToDataTable(string filePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var dataTable = new DataTable();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Read the first sheet

                if (worksheet.Dimension == null)
                    return dataTable;

                // Add columns
                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    var colName = worksheet.Cells[1, col].Text.Trim();
                    if (string.IsNullOrEmpty(colName))
                        colName = $"Column{col}";
                    dataTable.Columns.Add(colName);
                }

                // Add rows
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    var dataRow = dataTable.NewRow();
                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        dataRow[col - 1] = worksheet.Cells[row, col].Text;
                    }
                    dataTable.Rows.Add(dataRow);
                }
            }

            return dataTable;
        }
    }
}
