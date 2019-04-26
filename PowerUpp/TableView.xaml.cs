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
    /// Interaction logic for TableView.xaml
    /// </summary>
    public partial class TableView : Page
    {
        public TableView()
        {
            InitializeComponent();
            lblHeaderEx.Content = SelectionView.exerciseTitle;
            TableController excelData = new TableController();
            this.dgTable.DataContext = excelData; // Load data from spreadsheet into exercises table
            this.dgExTable.DataContext = excelData; // Load data from spreadsheet into specicic exercises
        }

        private void btnPrevious_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Content = new SelectionView();
        }

        private void btnNext_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Content = new ChartView();
        }
    }
}
