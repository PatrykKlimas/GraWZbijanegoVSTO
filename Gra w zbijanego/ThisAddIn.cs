using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace Gra_w_zbijanego
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        public Excel.Workbook GetActiveWorkbook()
        {
            return Application.ActiveWorkbook;
        }
        public Excel.Range GetActiveCell()
        {
            return Application.ActiveCell;
        }

    }
}
