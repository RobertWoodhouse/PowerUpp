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
    /// Interaction logic for SelectionView.xaml
    /// </summary>
    public partial class SelectionView : Page
    {
        public SelectionView()
        {
            InitializeComponent();
        }

        Enum selectedExercise;
        Enum selectedSets;
        string updateCells;

        //SelectionController selectCtrl = new SelectionController(); // CAUTION causes infinte Excel load

        private void cboExercise_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            selectedExercise = (Enum)cboExercise.SelectedItem;
        }

        private void cboSets_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            selectedSets = (Enum)cboSets.SelectedItem;
        }

        private void txbSets_TextChanged(object sender, TextChangedEventArgs e)
        {
            updateCells = txbSets.Text;
        }

        private void btnNext_Click(object sender, RoutedEventArgs e)
        {
            // Error message box for incorrect values
            bool nan = int.TryParse(updateCells, out int tempInt);

            if (selectedExercise == null)
            {
                MessageBox.Show("Invalid selection, please select Exercise", "Invalid selection");
                return;
            }

            if (selectedSets == null)
            {
                MessageBox.Show("Invalid selection, please select Sets", "Invalid selection");
                return;
            }

            if (!nan)
            {
                MessageBox.Show('"' + updateCells + '"' + " is not a valid number, please enter number of sets", "Invalid value");
                return;
            }

            if (tempInt <= 0)
            {
                MessageBox.Show('"' + updateCells + '"' + " in not a valid number, please enter number larger than 0", "Invalid value");
                return;
            }

            // Open Excel table file

            //selectCtrl.OpenWorkbook(SelectionController.loadFile);
            //selectCtrl.EditWorksheetCell((Enum)selectedExercise, (Enum)selectedSets, updateCells);


            // Open content into TableView with table
            //this.Content = new TableView();

            // TEST Open content into Frame with table
            //frmMainMenu.Content = new TableView();
            NavigationService.Content = new TableView();
        }

        private void btnPrevious_Click(object sender, RoutedEventArgs e)
        {
            //this.Content = new MainWindow();
            NavigationService.Content = new StartView();
        }
        
    }
}
