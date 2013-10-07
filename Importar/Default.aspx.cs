using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;


using System.Data.OleDb;
using System.Configuration;



namespace Importar
{
    //Este es un cambio en el código
    public partial class _Default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }


        public DataTable Import(String sExcelFile)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workBook = app.Workbooks.Open(sExcelFile, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            
            Microsoft.Office.Interop.Excel.Worksheet workSheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.ActiveSheet;

            int index = 0;
            object rowIndex = 1;

            DataTable dt = new DataTable();
//------------------------------------------------------
            int _Rows = 0;

            while ((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[rowIndex,1]).Value2 != null)
            {
                Range _WorkRange = (Microsoft.Office.Interop.Excel.Range).workSheet.Cells(string.Format("A{0}", _Rows));

                if (_WorkRange == null)
                    break;

                _Rows++;

            }


            DataRow row;
            
            while (((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[rowIndex,1]).Value2 != null)
            {
                
                rowIndex = 1 + index;
                row = dt.NewRow();
                for (int i=1; i<=18;i++)
                {
                row[i-1] = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[rowIndex, i]).Value2);
                }
                index++;
                dt.Rows.Add(row);
            }
            app.Workbooks.Close();
            return dt;
        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            Import("F:\\O1110713.xls");
        }



    }
}
