using System;
using System.IO;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace PowerUpp
{
    class ChartController
    {
        static string currentDirectory = Directory.GetCurrentDirectory();
        public static string filePath = System.IO.Path.Combine(currentDirectory, "Images", "ChartPic.jpg");
        private static Enum selectedExercise; // Get selected exercise from SelectionView in cboExercise_SelectionChanged()

        public static Enum SelectedExercise
        {
            get => selectedExercise;
            set => selectedExercise = value;
        }

        const string topLeft = "A1";

        private static string bottomRight = "B6";

        public static string BottomRight
        {
            get => bottomRight;
            set => bottomRight = value;
        }

        string graphTitle = "<Exercise> Chart";
        const string xAxis = "Date";
        const string yAxis = "Reps";

        public async Task CreateChart()
        {
            // Open Excel and get first worksheet.
            Excel.Worksheet xlWorksheet; // New worksheet   
            Excel.Range range;

            xlWorksheet = SelectionController.xlWorkbook.Worksheets[SelectedExercise.ToString()]; // Select specific exercise table

            // Add chart.
            var charts = xlWorksheet.ChartObjects() as Excel.ChartObjects;
            var chartObject = charts.Add(60, 20, 600, 300) as Excel.ChartObject;
            var chart = chartObject.Chart;

            graphTitle = selectedExercise.ToString().Replace("_"," ") + " Chart";

            try
            {
                range = xlWorksheet.get_Range(topLeft, bottomRight);
                chart.SetSourceData(range);

                // Set chart properties.
                chart.ChartType = Excel.XlChartType.xlLine;
                chart.ChartWizard(Source: range,
                    Title: graphTitle,
                    CategoryTitle: xAxis,
                    ValueTitle: yAxis);

                //export chart as picture file
                chart.Export(filePath, "JPG", System.Reflection.Missing.Value);

                // Save.
                await SelectionController.SaveFileAsync(true); // TEST
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                Console.WriteLine("Exception: " + ex + " thrown @ ChartController/CreateChart()");
            }
        }
    }
}
