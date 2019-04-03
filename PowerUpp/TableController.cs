using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace PowerUpp
{
    class TableController
    {
        string filePath = @"C:\Users\Robert Woodhouse\Google Drive\PowerUpp\PowerUppXL.xlsx";

        Excel.Application xlApp = new Excel.Application(); // Create new excel app in background process
        Excel.Workbook xlWorkbook; // New workbook
        Excel.Worksheet xlWorksheet; // New worksheet  
        Excel.Range range;

        public DataView Data
        {
            get
            {
                xlWorkbook = xlApp.Workbooks.Open(filePath);
                xlWorksheet = xlWorkbook.Worksheets[1];

                int column = 0;
                int row = 0;

                range = xlWorksheet.UsedRange;
                DataTable dt = new DataTable();
                dt.Columns.Add("Exercise");
                dt.Columns.Add("3 Sets");
                dt.Columns.Add("2 Sets");
                dt.Columns.Add("1 Set (Default)");
                dt.Columns.Add("Misc");
                for (row = 2; row <= range.Rows.Count; row++)
                {
                    DataRow dr = dt.NewRow();
                    for (column = 1; column <= range.Columns.Count; column++)
                    {
                        dr[column - 1] = (range.Cells[row, column] as Excel.Range).Value2;
                    }
                    dt.Rows.Add(dr);
                    dt.AcceptChanges();
                }
                xlWorkbook.Close(true, Missing.Value, Missing.Value);
                xlApp.Quit();
                return dt.DefaultView;
            }
        }
    }
}
