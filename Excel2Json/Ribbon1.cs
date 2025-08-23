using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Xml;
using System.Threading.Tasks;
using System.Drawing.Text;
using ClosedXML.Excel;
using Newtonsoft.Json;
using Microsoft.Office.Interop.Excel;
using JsonFormatting = Newtonsoft.Json.Formatting;  // バージョン間の曖昧さを回避
using Excel = Microsoft.Office.Interop.Excel;



namespace Excel2Json
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void checkBox1_Click(object sender, RibbonControlEventArgs e)
        {
            bool isChecked = ((RibbonCheckBox)sender).Checked;
            Globals.ThisAddIn.SetExportEnabled(isChecked);

            if (isChecked)
            {
                // 現在開いているブックを取得してエクスポート
                Excel.Workbook Wb = Globals.ThisAddIn.Application.ActiveWorkbook;
                Wb.Save();
            }
        }

        // Export Test Button
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook Wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Globals.ThisAddIn.ExportJson(Wb);
        }
    }
}