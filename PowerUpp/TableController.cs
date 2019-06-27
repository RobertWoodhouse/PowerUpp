using System;
using System.Data;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace PowerUpp
{
    class TableController
    {
        private static Enum selectedExercise; // Get selected exercise from SelectionView in cboExercise_SelectionChanged()

        public static Enum SelectedExercise
        {
            get { return selectedExercise; }
            set { selectedExercise = value; }
        }

        static Excel.Worksheet xlWorksheet1; // New worksheet  
        static Excel.Worksheet xlWorksheet2; // New worksheet  
        static Excel.Range xlRange;

        public DataView DataTable
        {
            get
            {
                DataTable dt = new DataTable();
                dt.Columns.Add("Exercise");
                dt.Columns.Add("3 Sets");
                dt.Columns.Add("2 Sets");
                dt.Columns.Add("1 Set");
                try
                {
                    xlRange = xlWorksheet1.UsedRange;

                    for (int row = 2; row <= xlRange.Rows.Count; row++)
                    {
                        DataRow dr = dt.NewRow();
                        for (int column = 1; column <= xlRange.Columns.Count; column++)
                        {
                            dr[column - 1] = (xlRange.Cells[row, column] as Excel.Range).Value2;
                        }
                        dt.Rows.Add(dr);
                        dt.AcceptChanges();
                    }
                }
                catch (IndexOutOfRangeException exMessage)
                {
                    Console.WriteLine("IndexOutOfRange Exception: " + exMessage + " thrown @ TableController/DataTable/get");
                }
                catch (COMException ex)
                {
                    Console.WriteLine("COMException Exception: " + ex + " thrown @ TableController/DataTable/get");
                }
                catch (NullReferenceException ex)
                {
                    Console.WriteLine("COMException Exception: " + ex + " thrown @ TableController/DataTable/get");
                }
                return dt.DefaultView;
            }
        }
        
        public DataView DataExercise
        {
            get
            {
                DataTable dt = new DataTable();
                dt.Columns.Add("Date");
                dt.Columns.Add("3 Sets");
                dt.Columns.Add("2 Sets");
                dt.Columns.Add("1 Set");
                
                try
                {
                    xlRange = xlWorksheet2.UsedRange;

                    for (int row = 2; row <= xlRange.Rows.Count; row++)
                    {
                        DataRow dr = dt.NewRow();
                        for (int column = 1; column <= xlRange.Columns.Count; column++)
                        {
                            dr[column - 1] = (xlRange.Cells[row, column] as Excel.Range).Value2;
                        }
                        dt.Rows.Add(dr);
                        dt.AcceptChanges();
                    }
                }
                catch (IndexOutOfRangeException exMessage)
                {
                    Console.WriteLine("IndexOutOfRange Exception: " + exMessage + " thrown @ TableController/DataExercise/get");
                }
                catch (COMException ex)
                {
                    Console.WriteLine("COMException Exception: " + ex + " thrown @ TableController/DataTable/get");
                }
                catch (NullReferenceException ex)
                {
                    Console.WriteLine("COMException Exception: " + ex + " thrown @ TableController/DataTable/get");
                }
                return dt.DefaultView;
            }
        }

        public async Task OpenExcelWorksheetAsync()
        {
            xlWorksheet1 = SelectionController.xlWorkbook.Worksheets["Exercise Table"];
            xlWorksheet2 = SelectionController.xlWorkbook.Worksheets[SelectedExercise.ToString()]; // Select specific exercise table
        }
    }
}
