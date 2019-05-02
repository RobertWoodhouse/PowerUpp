using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace PowerUpp
{
    /// <summary>
    /// Interaction logic for ChartView.xaml
    /// </summary>
    public partial class ChartView : Page
    {
        public ChartView()
        {
            InitializeComponent();
            lblHeader.Content = SelectionView.exerciseTitle + " Chart";
            ChartController chartCtl = new ChartController();
            chartCtl.CreateChart();
            this.imgChart.Source = new BitmapImage(new Uri(@"C:\Users\Robert Woodhouse\Google Drive\PowerUpp\Images\ChartPic.jpg"));
            //this.dgTable.DataContext = excelData; // Load data from spreadsheet into exercises table
            //this.dgExTable.DataContext = excelData; // Load data from spreadsheet into specicic exercises
        }

        private void btnPrevious_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Content = new TableView();
        }
    }
}
