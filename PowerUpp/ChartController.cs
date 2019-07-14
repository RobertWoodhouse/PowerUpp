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

        public static Enum SelectedExercise { get; set; } // Get selected exercise from SelectionView in cboExercise_SelectionChanged()

        const string TopLeft = "A1";

        public static string BottomRight { get; set; } = "B6";

        string graphTitle = "<Exercise> Chart";
        const string XAxis = "Date";
        const string YAxis = "Reps";

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

            graphTitle = SelectedExercise.ToString().Replace("_"," ") + " Chart";

            try
            {
                range = xlWorksheet.get_Range(TopLeft, BottomRight);
                chart.SetSourceData(range);

                // Set chart properties.
                chart.ChartType = Excel.XlChartType.xlLine;
                chart.ChartWizard(Source: range,
                    Title: graphTitle,
                    CategoryTitle: XAxis,
                    ValueTitle: YAxis);

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
