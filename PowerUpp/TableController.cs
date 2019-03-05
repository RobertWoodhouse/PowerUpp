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

        public DataView Data
        {
            get
            {
                Excel.Application xlApp = new Excel.Application(); // Create new excel app
                Excel.Workbook xlWorkbook; // New workbook
                Excel.Worksheet xlWorksheet; // New worksheet  
                Excel.Range range;
                //xlApp.Visible = true;
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
                        //dr[column - 1] = (range.Cells[row, column] as Excel.Range).Value2.ToString();
                        dr[column - 1] = (range.Cells[row, column] as Excel.Range);
                        //dr[column - 1] = 2;
                    }
                    dt.Rows.Add(dr);
                    dt.AcceptChanges();
                }
                xlWorkbook.Close(true, Missing.Value, Missing.Value);
                xlApp.Quit();
                return dt.DefaultView;
            }
        }
        /*
        public static bool loadFile;


        public void CreateWorkbook() // TEMP
        {
            xlApp.Visible = true;
            xlWorkbook = xlApp.Workbooks.Add();
            xlWorksheet = xlWorkbook.Worksheets[1]; // Worksheet the data is written onto

            try
            {
                // Column 1
                xlWorksheet.Cells[1, 1] = "Exercise";
                xlWorksheet.Cells[1, 2] = "3 Sets";
                xlWorksheet.Cells[1, 3] = "2 Sets";
                xlWorksheet.Cells[1, 4] = "1 Set (Default)";
                xlWorksheet.Cells[1, 5] = "Misc";

                Excel.Range range = (Excel.Range)xlWorksheet.Columns[1];

                range.Font.Bold = true;

                // Row A
                xlWorksheet.Cells[2, 1] = "Push Ups";
                xlWorksheet.Cells[3, 1] = "Squats";
                xlWorksheet.Cells[4, 1] = "Reverse Leg Lifts";
                xlWorksheet.Cells[5, 1] = "Dumbbell Side Bend";
                xlWorksheet.Cells[6, 1] = "Dumbbell Curls";
                xlWorksheet.Cells[7, 1] = "Standing Lunges";
                xlWorksheet.Cells[8, 1] = "Boxing";
                xlWorksheet.Cells[9, 1] = "Just Dance";
                xlWorksheet.Cells[10, 1] = "Sit Ups";
                xlWorksheet.Cells[11, 1] = "Shoulder Press";

                range = (Excel.Range)xlWorksheet.Rows[1];
                range.Font.Bold = true;
                range.Font.Color = System.Drawing.Color.Purple;

                //range.BorderAround2(Excel.XlLineStyle.xlContinuous); // Border around cells
            }
            catch (Exception exHandle)
            {
                Console.WriteLine("Exception: " + exHandle.Message);
            }
        }

        public void SaveAndQuit(bool saveFile)
        {
            if (saveFile)
            {
                xlWorkbook.SaveAs(filePath);
                Console.WriteLine("Spreadsheet saved");
            }

            xlWorkbook.Close();
            xlApp.Quit();
        }

        */
    }
}
