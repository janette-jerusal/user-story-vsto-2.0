// Logic to write to Excel
using System;
using System.Data;
using System.IO;
using OfficeOpenXml;

namespace UserStorySimilarityAddIn
{
    public static class ExcelWriter
    {
        public static void WriteExcel(DataTable dataTable, string filePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Matches");

                // Add headers
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = dataTable.Columns[i].ColumnName;
                }

                // Add data
                for (int row = 0; row < dataTable.Rows.Count; row++)
                {
                    for (int col = 0; col < dataTable.Columns.Count; col++)
                    {
                        worksheet.Cells[row + 2, col + 1].Value = dataTable.Rows[row][col];
                    }
                }

                // Save to file
                FileInfo file = new FileInfo(filePath);
                package.SaveAs(file);
            }
        }
    }
}

