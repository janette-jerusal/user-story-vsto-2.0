using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace UserStorySimilarityAddIn
{
    public static class ExcelReaderWriter
    {
        public static List<UserStory> ReadUserStories(string filePath)
        {
            var userStories = new List<UserStory>();

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);
            Excel._Worksheet worksheet = workbook.Sheets[1];
            Excel.Range range = worksheet.UsedRange;

            int rowCount = range.Rows.Count;

            for (int i = 2; i <= rowCount; i++)
            {
                string id = Convert.ToString((range.Cells[i, 1] as Excel.Range)?.Value2)?.Trim();
                string desc = Convert.ToString((range.Cells[i, 2] as Excel.Range)?.Value2)?.Trim();

                if (!string.IsNullOrEmpty(id) && !string.IsNullOrEmpty(desc))
                {
                    userStories.Add(new UserStory { ID = id, Desc = desc });
                }
            }

            workbook.Close(false);
            excelApp.Quit();

            return userStories;
        }

        public static void WriteResultsToNewSheet(List<SimilarityResult> results)
        {
            Excel.Worksheet newSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets.Add();
            newSheet.Name = "Similarity Results";

            newSheet.Cells[1, 1] = "Story A ID";
            newSheet.Cells[1, 2] = "Story A Desc";
            newSheet.Cells[1, 3] = "Story B ID";
            newSheet.Cells[1, 4] = "Story B Desc";
            newSheet.Cells[1, 5] = "Similarity Score";

            int row = 2;
            foreach (var result in results)
            {
                newSheet.Cells[row, 1] = result.StoryA_ID;
                newSheet.Cells[row, 2] = result.StoryA_Desc;
                newSheet.Cells[row, 3] = result.StoryB_ID;
                newSheet.Cells[row, 4] = result.StoryB_Desc;
                newSheet.Cells[row, 5] = result.Score;
                row++;
            }

            newSheet.Columns.AutoFit();
        }
    }

    public class UserStory
    {
        public string ID { get; set; }
        public string Desc { get; set; }
    }

    public class SimilarityResult
    {
        public string StoryA_ID { get; set; }
        public string StoryA_Desc { get; set; }
        public string StoryB_ID { get; set; }
        public string StoryB_Desc { get; set; }
        public double Score { get; set; }
    }
}
