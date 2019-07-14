using System.ComponentModel;
using System.Windows;
using System;

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

            this.Closing += new CancelEventHandler(MainWindow_Closing);
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

        private void MainWindow_Closing(object sender, CancelEventArgs e)
        {
            SelectionController.StopExcelAppAsync();
            Application.Current.Shutdown();
        }
    }
}
