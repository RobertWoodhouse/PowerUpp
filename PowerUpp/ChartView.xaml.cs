using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;

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
            ChartController chartCtl = new ChartController();
            chartCtl.CreateChart();
            this.imgChart.Source = new BitmapImage(new Uri(ChartController.filePath));
        }

        private void btnPrevious_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Content = new TableView();
        }
    }
}
