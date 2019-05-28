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
        string filePath = @"C:\Users\Robert Woodhouse\Google Drive\PowerUpp\Images\Watermark.jpg";


        public static string ExerciseTitle { get; set; }

        SelectionController selectCtrl = new SelectionController(); // CAUTION causes infinte Excel load

        private void cboExercise_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            selectedExercise = (Enum)cboExercise.SelectedItem;
            TableController.SelectedExercise = (Enum)cboExercise.SelectedItem; // Set selected exercise for TableController
            ChartController.SelectedExercise = (Enum)cboExercise.SelectedItem; // Set selected exercise for ChartController
        }

        private void cboSets_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            selectedSets = (Enum)cboSets.SelectedItem;
        }

        private void txbSets_TextChanged(object sender, TextChangedEventArgs e)
        {
            //updateCells = txbSets.Text;

            if (txbSets.Text == "")
            {
                // Create an ImageBrush.
                ImageBrush textImageBrush = new ImageBrush();
                //textImageBrush.ImageSource = new BitmapImage(new Uri(@"Images\Watermark.jpg", UriKind.Relative));
                textImageBrush.ImageSource = new BitmapImage(new Uri(filePath, UriKind.Relative));
                textImageBrush.AlignmentX = AlignmentX.Left;
                textImageBrush.Stretch = Stretch.None;
                // Use the brush to paint the button's background.
                txbSets.Background = textImageBrush;

            }
            else
            {
                txbSets.Background = null;
                updateCells = txbSets.Text;
            }
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

            ExerciseTitle = selectedExercise.ToString(); // Assign string value for label title in TableView

            // Open Excel table file
            selectCtrl.OpenWorkbook(SelectionController.loadFile);
            selectCtrl.EditTableCellAsync((Enum)selectedExercise, (Enum)selectedSets, updateCells).Wait();

            selectCtrl.CreateEditWorksheet((Enum)selectedExercise); //TODO see if new worksheet is created and updated
            //selectCtrl.EditExerciseCellAsync(updateCells).Wait(); // TODO: fix data loaded into wrong WPF table
            selectCtrl.EditExerciseCellAsync((Enum)selectedSets, updateCells).Wait();


            // Open content into frame with table
            NavigationService.Content = new TableView();
        }

        private void btnPrevious_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Content = new StartView();
        }
        
    }
}
