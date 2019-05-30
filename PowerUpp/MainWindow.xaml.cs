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
using Gat.Controls;

namespace PowerUpp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>

    public partial class MainWindow : Window
    {
        About about = new About();

        public MainWindow()
        {
            InitializeComponent();
            frmMainMenu.Content = new StartView();
            //About about = new About();
        }

        private void NewBtnMenu_Click(object sender, RoutedEventArgs e)
        {
            // Add pop up to prompt new spreadsheet
            frmMainMenu.Content = new SelectionView();
            SelectionController.loadFile = false;
        }

        private void ExitBtnMenu_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        private void AboutBtnMenu_Click(object sender, RoutedEventArgs e)
        {
            //About about = new About();
            about.ApplicationLogo = new BitmapImage(new Uri(@"C:\Users\Robert Woodhouse\Google Drive\PowerUpp\Images\SquatsIcon.ico"));
            about.Version = "v1.05";
            about.AdditionalNotes = "Power Upp is an application used to track your resistance exercise data over the course of time and present it visually as a chart";
            about.PublisherLogo = new BitmapImage(new Uri(@"C:\Users\Robert Woodhouse\Google Drive\PowerUpp\Images\DumbbellIcon.ico"));
            about.Copyright = "© 2019 Robert Woodhouse \nAll rights reserved";
            about.HyperlinkText = "https://github.com/RobertWoodhouse/PowerUpp";
            about.Show();
        }
    }
}
