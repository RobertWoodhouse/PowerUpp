using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace PowerUpp
{
    enum Exercise { Push_Ups = 2, Squats = 3, Reverse_Leg_Lift = 4, Dumbbell_Side_Bend = 5, Dumbbell_Curls = 6, Standing_Lunges = 7, Boxing = 8, Just_Dance = 9, Sit_Ups = 10, Shoulder_Press = 11 }; // Column A
    enum Sets { Three_Sets = 2, Two_Sets, One_Set, Misc }; // Row 1

    class TableController
    {
        string filePath = @"C:\Users\Robert Woodhouse\Google Drive\PowerUpp\PowerUppXL.xlsx";

        Excel.Application xlApp = new Excel.Application(); // Create new excel app
        Excel.Workbook xlWorkbook; // New workbook
        Excel.Worksheet xlWorksheet; // New worksheet  

        public void LoadWorkbook(string loadFile)
        {
            xlApp.Visible = true;
            xlWorkbook = xlApp.Workbooks.Open(loadFile);
        }

        public void CreateWorkbook() // TEMP
        {
            xlApp.Visible = true;
            xlWorkbook = xlApp.Workbooks.Add();
            xlWorksheet = xlWorkbook.Worksheets[1]; // Worksheet the data is written onto
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

        public void EditWorksheetCell(Enum exercise, Enum sets, string updateCell)
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
            finally
            {
                SaveAndQuit(true);
                //SaveAndQuit(false);
            }
        }

        public void OpenWorkbook()
        {
            LoadWorkbook(filePath);
            //CreateWorkbook();

            try
            {
                /* Column 1 */
                xlWorksheet.Cells[1, 1] = "Exercise";
                xlWorksheet.Cells[1, 2] = "3 Sets";
                xlWorksheet.Cells[1, 3] = "2 Sets";
                xlWorksheet.Cells[1, 4] = "1 Set (Default)";
                xlWorksheet.Cells[1, 5] = "Misc";

                Excel.Range range = (Excel.Range)xlWorksheet.Columns[1];

                range.Font.Bold = true;

                /* Row A */
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
            finally
            {
                //xlWorkbook.Close();
                //xlApp.Quit();
            }
        }
    }
}
