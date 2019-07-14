using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;

namespace PowerUpp
{
    /// <summary>
    /// Interaction logic for SelectionView.xaml
    /// </summary>
    public partial class SelectionView : Page
    {
        Enum selectedExercise;
        Enum selectedSets;
        string updateCells;
        static string currentDirectory = Directory.GetCurrentDirectory();
        string filePath = System.IO.Path.Combine(currentDirectory, "Images", "Watermark.jpg");
        public static string ExerciseTitle { get; set; }

        SelectionController selectionCtrl = new SelectionController(); // CAUTION can cause infinte Excel load

        public SelectionView()
        {
            InitializeComponent();
        }

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
            if (txbSets.Text == "")
            {
                // Create an ImageBrush.
                ImageBrush textImageBrush = new ImageBrush();
                textImageBrush.ImageSource = new BitmapImage(new Uri(filePath, UriKind.Relative));
                textImageBrush.AlignmentX = AlignmentX.Left;
                textImageBrush.Stretch = Stretch.None;
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
                MessageBox.Show("Invalid selectionCtrl, please select Exercise", "Invalid selectionCtrl");
                return;
            }

            if (selectedSets == null)
            {
                MessageBox.Show("Invalid selectionCtrl, please select Sets", "Invalid selectionCtrl");
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

            SelectionController.StartExcelAppAsync();

            // Open Excel table file
            selectionCtrl.OpenWorkbook(SelectionController.loadFile);
            selectionCtrl.EditTableCellAsync((Enum)selectedExercise, (Enum)selectedSets, updateCells).Wait();

            selectionCtrl.CreateEditWorksheet((Enum)selectedExercise);
            selectionCtrl.EditExerciseCellAsync((Enum)selectedSets, updateCells).Wait();
            selectionCtrl.UpdateExerciseCellsAsync().Wait(); // Updates all blank cells on worksheet

            // Open content into frame with table
            NavigationService.Content = new TableView();
        }

        private void btnPrevious_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Content = new StartView();
        }
    }
}
