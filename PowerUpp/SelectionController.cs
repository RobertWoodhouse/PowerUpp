using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace PowerUpp
{
    enum Exercise { Push_Ups = 2, Squats = 3, Reverse_Leg_Lift = 4, Dumbbell_Side_Bend = 5, Dumbbell_Curls = 6, Standing_Lunges = 7, Boxing = 8, Just_Dance = 9, Sit_Ups = 10, Shoulder_Press = 11 }; // Column A
    enum Sets { Three_Sets = 2, Two_Sets, One_Set, Misc }; // Row 1
    
    class SelectionController
    {
        string filePath = @"C:\Users\Robert Woodhouse\Google Drive\PowerUpp\PowerUppXL.xlsx";
        public static bool loadFile;

        Excel.Application xlApp = new Excel.Application(); // Create new excel app
        Excel.Workbook xlWorkbook; // New workbook
        Excel.Worksheet xlWorksheet; // New worksheet
        Excel.Worksheet xlWorksheetEx; // New worksheet
        Excel.Range xlRange; // Worksheet column - row range

        int colRange = 0;
        int rowRange = 0;


        public void LoadWorkbook(string var)
        {
            xlApp.Visible = true; // Stops Excel app from loading
            xlWorkbook = xlApp.Workbooks.Open(var);
            //xlWorksheet = xlWorkbook.Worksheets[1];
            xlWorksheet = xlWorkbook.Worksheets["Exercise Table"];
        }

        public void CreateWorkbookTable() // TEMP
        {
            xlApp.Visible = true; // Stops Excel app from loading
            xlWorkbook = xlApp.Workbooks.Add();
            //xlWorksheet = xlWorkbook.Worksheets[1]; // Worksheet the data is written onto
            xlWorksheet = (Excel.Worksheet)xlApp.ActiveSheet;
            xlWorksheet.Name = "Exercise Table";

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

        public void CreateEditWorksheet(Enum exercise) // TODO Change var to int if enum doesn't work
        {
            try
            {
                // Open existing worksheet
                xlWorksheetEx = xlApp.Worksheets[exercise.ToString()];
                xlWorksheetEx.Select(true);
            }
            catch (COMException ex)
            {
                // Add worksheet if it does not exist in workbook
                xlWorksheetEx = xlWorkbook.Sheets.Add();
                xlWorksheetEx.Name = exercise.ToString();
                Console.WriteLine("Exception: " + ex.Message);
            }

            try
            {
                // Column 1
                xlWorksheetEx.Cells[1, 1] = "Date";
                xlWorksheetEx.Cells[1, 2] = exercise.ToString();

                Excel.Range range = (Excel.Range)xlWorksheetEx.Columns[1];

                range.Font.Bold = true;

                // Row A
                //xlWorksheet.Cells[2, 1] = "20/10/2018";

                range = (Excel.Range)xlWorksheetEx.Rows[1];
                range.Font.Bold = true;
                range.Font.Color = System.Drawing.Color.Crimson;

                //range.BorderAround2(Excel.XlLineStyle.xlContinuous); // Border around cells
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
            }
        }
        
        /*
        public void SaveAndQuit(bool saveFile)
        {
            if (saveFile)
            {
                //xlWorkbook.SaveAs(filePath);
                xlWorkbook.Save();
                Console.WriteLine("Spreadsheet saved");
            }
            xlWorkbook.Close();
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorksheet);
            Marshal.ReleaseComObject(xlWorkbook);
            Marshal.ReleaseComObject(xlApp);
        }
        */

        async Task SaveAndQuitAsync (bool saveFile)
        {
            if (saveFile)
            {
                //xlWorkbook.SaveAs(filePath);
                xlWorkbook.Save();
                Console.WriteLine("Spreadsheet saved");
            }
            /*
            xlWorkbook.Close();
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorksheet);
            Marshal.ReleaseComObject(xlWorkbook);
            Marshal.ReleaseComObject(xlApp);
            */
        }

        public async Task EditTableCellAsync(Enum exercise, Enum sets, string updateCell)
        {
            try
            {
                Console.WriteLine("Enter cell... ");
                xlWorksheet.Cells[exercise, sets] = updateCell;
            }
            catch (Exception exMessage)
            {
                Console.WriteLine("Error message " + exMessage);
            }
            finally // TODO: sort save
            {
                //SaveAndQuit(true); // Saves file before closing, allowing data to be loaded into WPF table
                await SaveAndQuitAsync(true);
                //SaveAndQuit(false);
            }
        }

        public async Task EditExerciseCellAsync(string updateCell)
        {
            xlRange = xlWorksheetEx.UsedRange;

            //colRange = xlRange.Columns.Count;
            rowRange = xlRange.Rows.Count;

            DateTime dateToday = DateTime.UtcNow.Date;
            string date = dateToday.ToString("dd/MM/yyyy");
            string cellValue = ((Excel.Range)xlWorksheetEx.Cells[rowRange, 1]).Text;

            try
            {
                if (cellValue == date || cellValue == "" || cellValue == null) // If cell date == date today
                {
                    xlWorksheetEx.Cells[rowRange, 1].NumberFormat = "@"; // Prevent autoformat by setting cell format to "text"
                    xlWorksheetEx.Cells[rowRange, 1] = date;
                    xlWorksheetEx.Cells[rowRange, 2] = updateCell;
                }
                else 
                {
                    xlWorksheetEx.Cells[rowRange+1, 1].NumberFormat = "@"; // Prevent autoformat by setting cell format to "text"
                    xlWorksheetEx.Cells[rowRange+1, 1] = date;
                    xlWorksheetEx.Cells[rowRange+1, 2] = updateCell;
                }
            }
            catch (AggregateException exMessage)
            {
                Console.WriteLine("Aggrage error message " + exMessage);
            }
            catch (Exception exMessage)
            {
                Console.WriteLine("Error message " + exMessage);
            }
            finally // TODO: sort save
            {
                //SaveAndQuit(true); // Saves file before closing, allowing data to be loaded into WPF table
                await SaveAndQuitAsync(true);
                //SaveAndQuit(false);
            }
        }

        public void OpenWorkbook(bool var)
        {
            if (var == true) LoadWorkbook(filePath);
            else CreateWorkbookTable();
        }
    }
}
