using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace PowerUpp
{
    class ChartController
    {
        string filePath = @"C:\Users\Robert Woodhouse\Google Drive\PowerUpp\PowerUppXL.xlsx";
        private static Enum selectedExercise; // Get selected exercise from SelectionView in cboExercise_SelectionChanged()

        public static Enum SelectedExercise
        {
            get { return selectedExercise; }
            set { selectedExercise = value; }
        }

        string topLeft = "A1";

        private static string bottomRight = "B6";

        public static string BottomRight
        {
            get { return bottomRight; }
            set { bottomRight = value; }
        }

        string graphTitle = "<Exercise> Chart";
        string xAxis = "Date";
        string yAxis = "Reps";

        public async Task CreateChart()
        {
            // Open Excel and get first worksheet.
            Excel.Application xlApp = new Excel.Application(); // Create new excel app in background process
            Excel.Workbook xlWorkbook; // New workbook
            Excel.Worksheet xlWorksheet; // New worksheet   
            Excel.Range range;

            xlWorkbook = xlApp.Workbooks.Open(filePath);
            xlWorksheet = xlWorkbook.Worksheets[SelectedExercise.ToString()]; // Select specific exercise table

            // Add chart.
            var charts = xlWorksheet.ChartObjects() as Excel.ChartObjects;
            var chartObject = charts.Add(60, 20, 600, 300) as Excel.ChartObject;
            var chart = chartObject.Chart;

            graphTitle = selectedExercise.ToString() + " Chart";
            Console.WriteLine("Chart Row Range = {0}", bottomRight);

            range = xlWorksheet.get_Range(topLeft, bottomRight);
            chart.SetSourceData(range);

            // Set chart properties.
            chart.ChartType = Excel.XlChartType.xlLine;
            chart.ChartWizard(Source: range,
                Title: graphTitle,
                CategoryTitle: xAxis,
                ValueTitle: yAxis);

            //export chart as picture file
            chart.Export(@"C:\Users\Robert Woodhouse\Google Drive\PowerUpp\Images\ChartPic.jpg", "JPG", System.Reflection.Missing.Value);

            // Save.
            xlWorkbook.Save();
        }
    }
}
