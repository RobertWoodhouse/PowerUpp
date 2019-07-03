using System.Windows;

namespace PowerUpp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>

    public partial class MainWindow : Window
    {
        MessageBoxController controller = new MessageBoxController();

        public MainWindow()
        {
            InitializeComponent();
            frmMainMenu.Content = new StartView();
        }

        private void NewBtnMenu_Click(object sender, RoutedEventArgs e)
        {
            controller.NewMenu(frmMainMenu, this);
        }

        private void ExitBtnMenu_Click(object sender, RoutedEventArgs e)
        {
            SelectionController.StopExcelAppAsync();
            Application.Current.Shutdown();
        }

        private void AboutBtnMenu_Click(object sender, RoutedEventArgs e)
        {
            controller.AboutMenu();
        }
    }
}
