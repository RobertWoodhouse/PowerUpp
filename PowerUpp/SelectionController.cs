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
        //Excel.Worksheet xlWorksheet; // New worksheet


        public void LoadWorkbook(string var)
        {
            xlApp.Visible = true; // Stops Excel app from loading
            xlWorkbook = xlApp.Workbooks.Open(var);
            xlWorksheet = xlWorkbook.Worksheets[1];
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
            //xlApp.Visible = true; // Stops Excel app from loading
            //xlWorkbook = xlApp.Workbooks.Add();
            //xlWorksheet = xlWorkbook.Worksheets[1]; // Worksheet the data is written onto, TODO change to exercise
            //xlWorksheet = (Excel.Worksheet)xlApp.ActiveSheet;
            //xlWorksheet = xlApp.Sheets.Add(1);
            xlWorksheet = xlWorkbook.Sheets.Add(); // Add new worksheet to workbook
            xlWorksheet.Name = exercise.ToString(); // TODO: fix ERROR where name of worksheet already exists

            try
            {
                // Column 1
                xlWorksheet.Cells[1, 1] = "Date";
                xlWorksheet.Cells[1, 2] = exercise.ToString();

                Excel.Range range = (Excel.Range)xlWorksheet.Columns[1];

                range.Font.Bold = true;

                // Row A
                //xlWorksheet.Cells[2, 1] = "20/10/2018";
                //xlWorksheet.Cells[3, 1] = "21/10/2018";
                //xlWorksheet.Cells[4, 1] = "22/10/2018";

                range = (Excel.Range)xlWorksheet.Rows[1];
                range.Font.Bold = true;
                range.Font.Color = System.Drawing.Color.Crimson;

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

        public void EditTableCell(Enum exercise, Enum sets, string updateCell)
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
                SaveAndQuit(true); // Saves file before closing, allowing data to be loaded into WPF table
                //SaveAndQuit(false);
            }
        }

        public void EditExerciseCell(string updateCell)
        {
            DateTime date = DateTime.UtcNow.Date;
            date.ToString("dd/MM/yyyy");

            try
            {
                Console.WriteLine("Enter cell... ");
                //xlWorksheet.Cells[exercise, sets] = updateCell;
                //for (int i = 2; )
                //xlWorksheet.Cells[2, 1] = thisDay.ToString();
                xlWorksheet.Cells[2, 1] = date;
                xlWorksheet.Cells[2, 2] = updateCell;
            }
            catch (Exception exMessage)
            {
                Console.WriteLine("Error message " + exMessage);
            }
            finally // TODO: sort save
            {
                SaveAndQuit(true); // Saves file before closing, allowing data to be loaded into WPF table
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
