using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools.Excel.Extensions;

namespace ExcelAddIn
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void createList_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet worksheet = Globals.Factory.GetVstoObject(
              Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1]);

            Excel.Range selection = Globals.ThisAddIn.Application.Selection as Excel.Range;

            string text = selection.Value.ToString();

            var range = worksheet.get_Range("F4", "F4");
            range.Select();
            range.Value = text;
            string[] cellValues = text.Split('\n');

            //range = worksheet.get_Range("G4", "G4");
            //range = worksheet.get_Range()
            worksheet.Cells[6, 6] = cellValues[0];
            //range.Select();
            //range.Value = cellValues[0];

        }
    }
}
