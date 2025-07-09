// Ribbon1 UI logic
using System;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;

namespace UserStorySimilarityAddIn
{
    public partial class Ribbon1 : RibbonBase
    {
        public Ribbon1() : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void compareButton_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            if (workbook == null)
            {
                MessageBox.Show("No active workbook.");
                return;
            }

            Excel.Sheets sheets = workbook.Sheets;
            if (sheets.Count < 2)
            {
                MessageBox.Show("Workbook must have at least 2 sheets.");
                return;
            }

            Excel.Worksheet sheet1 = sheets[1];
            Excel.Worksheet sheet2 = sheets[2];

            // Placeholder logic — replace with actual comparison logic
            MessageBox.Show($"Comparing Sheet1: {sheet1.Name} with Sheet2: {sheet2.Name}");
        }
    }
}
