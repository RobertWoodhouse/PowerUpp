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
        //const string filePath = "C:\\Book1.xlsx";
        string filePath = @"C:\Users\Robert Woodhouse\Google Drive\PowerUpp\PowerUppXL.xlsx";
        public static Enum selectedExercise; // Get selected exercise from SelectionView in cboExercise_SelectionChanged()

        const string topLeft = "A1";
        const string bottomRight = "A4";
        const string graphTitle = "Graph Title";
        const string xAxis = "Time";
        const string yAxis = "Value";
 
        public void CreateChart()
        {
            // Open Excel and get first worksheet.
            //var xlApp = new Application();
            Excel.Application xlApp = new Excel.Application(); // Create new excel app in background process
            //var xlWorkbook = xlApp.Workbooks.Open(filePath);
            Excel.Workbook xlWorkbook; // New workbook
            //var xlWorksheet = xlWorkbook.Worksheets[1] as Microsoft.Office.Interop.Excel.Worksheet;
            Excel.Worksheet xlWorksheet; // New worksheet   
            Excel.Range range;

            xlWorkbook = xlApp.Workbooks.Open(filePath);
            xlWorksheet = xlWorkbook.Worksheets[selectedExercise.ToString()]; // Select specific exercise table

            // Add chart.
            var charts = xlWorksheet.ChartObjects() as
                Microsoft.Office.Interop.Excel.ChartObjects;
            var chartObject = charts.Add(60, 10, 300, 300) as
                Microsoft.Office.Interop.Excel.ChartObject;
            var chart = chartObject.Chart;

            // Set chart range.
            //var range = xlWorksheet.get_Range(topLeft, bottomRight);
            range = xlWorksheet.get_Range(topLeft, bottomRight);
            chart.SetSourceData(range);

            // Set chart properties.
            chart.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlLine;
            chart.ChartWizard(Source: range,
                Title: graphTitle,
                CategoryTitle: xAxis,
                ValueTitle: yAxis);

            // Save.
            xlWorkbook.Save();
        }
    }
}
