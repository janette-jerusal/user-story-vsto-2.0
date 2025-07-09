// Logic to read from Excel
using System;
using System.Data;
using System.IO;
using OfficeOpenXml;

namespace UserStorySimilarityAddIn
{
    public static class ExcelReader
    {
        public static DataTable ReadExcel(string filePath)
        {
            var table = new DataTable();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];

                // Add columns
                for (int col = worksheet.Dimension.Start.Column; col <= worksheet.Dimension.End.Column; col++)
                {
                    table.Columns.Add(worksheet.Cells[1, col].Text);
                }

                // Add rows
                for (int row = worksheet.Dimension.Start.Row + 1; row <= worksheet.Dimension.End.Row; row++)
                {
                    var newRow = table.NewRow();
                    for (int col = worksheet.Dimension.Start.Column; col <= worksheet.Dimension.End.Column; col++)
                    {
                        newRow[col - 1] = worksheet.Cells[row, col].Text;
                    }
                    table.Rows.Add(newRow);
                }
            }

            return table;
        }
    }
}

