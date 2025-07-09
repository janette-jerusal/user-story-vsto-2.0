using Microsoft.Office.Tools.Ribbon;
using System;
using System.Data;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace UserStorySimilarityAddIn
{
    public partial class MyRibbon
    {
        private void MyRibbon_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void CompareButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog
                {
                    Filter = "Excel Files|*.xlsx;*.xls",
                    Multiselect = true,
                    Title = "Select 2 Excel files with User Stories"
                };

                if (openFileDialog.ShowDialog() != DialogResult.OK || openFileDialog.FileNames.Length != 2)
                {
                    MessageBox.Show("Please select exactly 2 Excel files.", "Input Error");
                    return;
                }

                DataTable tableA = ExcelReader.ReadUserStories(openFileDialog.FileNames[0]);
                DataTable tableB = ExcelReader.ReadUserStories(openFileDialog.FileNames[1]);

                double threshold = 0.75; // This can be customized later through a settings UI if needed

                DataTable results = UserStoryComparer.CompareUserStories(tableA, tableB, threshold);

                if (results.Rows.Count == 0)
                {
                    MessageBox.Show("No matches found above the similarity threshold.", "Result");
                    return;
                }

                string tempPath = Path.Combine(Path.GetTempPath(), "UserStoryMatches.xlsx");
                ExcelWriter.WriteToExcel(results, tempPath);

                Excel.Application app = Globals.ThisAddIn.Application;
                app.Workbooks.Open(tempPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred:\n" + ex.Message, "Error");
            }
        }
    }
}
