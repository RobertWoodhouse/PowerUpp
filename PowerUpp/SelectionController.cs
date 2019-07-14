using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace PowerUpp
{
    enum Exercise { Push_Ups = 2, Squats = 3, Reverse_Leg_Lift = 4, Dumbbell_Side_Bend = 5, Dumbbell_Curls = 6, Standing_Lunges = 7, Sit_Ups, Shoulder_Press }; // Column A
    enum Sets { Three_Sets = 2, Two_Sets, One_Set}; // Row 1
    
    class SelectionController
    {
        static string currentDirectory = Directory.GetCurrentDirectory();
        string filePath = System.IO.Path.Combine(currentDirectory, "Resources", "PowerUppXL.xlsx");
        public static bool loadFile;

        public static Excel.Application xlApp; // Create new excel app
        public static Excel.Workbook xlWorkbook; // New workbook
        public static Excel.Worksheet xlWorksheet; // New worksheet
        public static Excel.Worksheet xlWorksheetEx; // New worksheet
        static Excel.Range xlRange; // Worksheet column - row xlRange

        private int rowRange;

        public delegate void ExcelControllerEventHandler(object source, EventArgs args);

        public event ExcelControllerEventHandler ExcelControl;

        public void CloseExcel()
        {
            xlWorkbook.Close();
            xlApp.Quit();

            OnClosedExcel();
        }

        protected virtual void OnClosedExcel()
        {
            ExcelControl?.Invoke(this, EventArgs.Empty);
        }

        public void OpenWorkbook(bool var)
        {
            if (var == true) LoadWorkbook(filePath);
            else CreateWorkbookTable();
        }

        public void LoadWorkbook(string var)
        {
            xlApp.Visible = false; // Stops Excel app from loading
            xlWorkbook = xlApp.Workbooks.Open(var);
            xlWorksheet = xlWorkbook.Worksheets["Exercise Table"]; // Worksheet the data is written onto
        }

        public void CreateWorkbookTable()
        {
            xlApp.Visible = false; // Stops Excel app from loading
            xlWorkbook = xlApp.Workbooks.Add();
            xlWorksheet = (Excel.Worksheet)xlApp.ActiveSheet; // Worksheet the data is written onto
            xlWorksheet.Name = "Exercise Table";

            try
            {
                // Column 1
                xlWorksheet.Cells[1, 1] = "Exercise";
                xlWorksheet.Cells[1, 2] = "3 Sets";
                xlWorksheet.Cells[1, 3] = "2 Sets";
                xlWorksheet.Cells[1, 4] = "1 Set";

                Excel.Range range = (Excel.Range)xlWorksheet.Columns[1];

                range.Font.Bold = true;

                // Row A
                xlWorksheet.Cells[2, 1] = "Push Ups";
                xlWorksheet.Cells[3, 1] = "Squats";
                xlWorksheet.Cells[4, 1] = "Reverse Leg Lifts";
                xlWorksheet.Cells[5, 1] = "Dumbbell Side Bend";
                xlWorksheet.Cells[6, 1] = "Dumbbell Curls";
                xlWorksheet.Cells[7, 1] = "Standing Lunges";
                xlWorksheet.Cells[8, 1] = "Sit Ups";
                xlWorksheet.Cells[9, 1] = "Shoulder Press";

                range = (Excel.Range)xlWorksheet.Rows[1];
                range.Font.Bold = true;
                range.Font.Color = System.Drawing.Color.Purple;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message + " thrown @ SelectionController/CreateWorkbookTable()");
            }
        }

        public void CreateEditWorksheet(Enum exercise)
        {
            try
            {
                // Open existing worksheet
                xlWorksheetEx = xlApp.Worksheets[exercise.ToString()];
                xlWorksheetEx.Select(true);
            }
            catch (COMException ex)
            {
                xlWorksheetEx = xlWorkbook.Sheets.Add();
                xlWorksheetEx.Name = exercise.ToString();
                Console.WriteLine("COM Exception: " + ex.Message + " thrown @ SelectionController/CreateEditWorksheet()");
            }

            try
            {
                // Column 1
                xlWorksheetEx.Cells[1, 1] = "Date";
                xlWorksheetEx.Cells[1, 2] = "3 Sets";
                xlWorksheetEx.Cells[1, 3] = "2 Sets";
                xlWorksheetEx.Cells[1, 4] = "1 Set";

                Excel.Range range = (Excel.Range)xlWorksheetEx.Columns[1];

                range.Font.Bold = true;

                // Row A
                range = (Excel.Range)xlWorksheetEx.Rows[1];
                range.Font.Bold = true;
                range.Font.Color = System.Drawing.Color.Crimson;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message + " thrown @ SelectionController/CreateEditWorksheet()");
            }
        }

        public static async Task SaveFileAsync (bool saveFile)
        {
            if (saveFile)
            {
                xlWorkbook.Save();
                Console.WriteLine("Spreadsheet saved");
            }
        }

        public static async Task StartExcelAppAsync() 
        {
            if (xlApp == null) xlApp = new Excel.Application(); // Create new excel app
        }

        public static async Task StopExcelAppAsync()
        {
            xlWorkbook.Close();
            xlApp.Quit();
        }

        public async Task EditTableCellAsync(Enum exercise, Enum sets, string updateCell)
        {
            try
            {
                xlWorksheet.Cells[exercise, sets] = updateCell;
            }
            catch (Exception exMessage)
            {
                Console.WriteLine("Exception: " + exMessage + " thrown @ SelectionController/EditTableCellAsync()");
            }
            finally
            {
                await SaveFileAsync(true);
            }
        }

        public async Task EditExerciseCellAsync(Enum sets, string updateCell)
        {
            xlRange = xlWorksheetEx.UsedRange;
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
                    xlWorksheetEx.Cells[rowRange, sets] = updateCell;
                    ChartController.BottomRight = "D" + rowRange; // Set Chart Row Range var in ChartController

                }
                else
                {
                    xlWorksheetEx.Cells[rowRange + 1, 1].NumberFormat = "@"; // Prevent autoformat by setting cell format to "text"
                    xlWorksheetEx.Cells[rowRange + 1, 1] = date;
                    xlWorksheetEx.Cells[rowRange + 1, sets] = updateCell;
                    ChartController.BottomRight = "D" + (rowRange + 1); // Set Chart Row Range var in ChartController
                }
            }
            catch (AggregateException exMessage)
            {
                Console.WriteLine("Aggregate Exception: " + exMessage + " thrown @ SelectionController/EditExerciseCellAsync()");
            }
            catch (Exception exMessage)
            {
                Console.WriteLine("Exception: " + exMessage + " thrown @ SelectionController/EditExerciseCellAsync()");
            }
            finally // TODO: sort save
            {
                await SaveFileAsync(true);
            }
        }

        public async Task UpdateExerciseCellsAsync()
        {
            string[] temp = new string[] { "0", "0", "0" };
            string cellValue = "";
            xlRange = xlWorksheetEx.UsedRange;
            rowRange = xlRange.Rows.Count;

            try
            {
                for (int row = 2; row <= rowRange; row++)
                {
                    for (int col = 2; col <= 4; col++)
                    {
                        cellValue = ((Excel.Range)xlWorksheetEx.Cells[row, col]).Text;

                        if(!string.IsNullOrEmpty(cellValue)) temp[col - 2] = cellValue;

                        else xlWorksheetEx.Cells[row, col] = temp[col - 2];
                    }
                }
            }
            catch (AggregateException exMessage) 
            {
                Console.WriteLine("Aggregate Exception: " + exMessage + " thrown @ SelectionController/UpdateExerciseCellsAsync()");
            }
            catch (Exception exMessage)
            {
                Console.WriteLine("Exception: " + exMessage + " thrown @ SelectionController/UpdateExerciseCellsAsync()");
            }
            finally // TODO: sort save
            {
                await SaveFileAsync(true);
            }
        }
    }
}
