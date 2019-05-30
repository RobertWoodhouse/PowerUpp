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
        //AboutControlView about = new AboutControlView();
        //AboutControlViewModel vmAbout = new AboutControlViewModel();
        MessageBoxView messageBox = new MessageBoxView();
        MessageBoxViewModel vmMessageBox = new MessageBoxViewModel();


        public MainWindow()
        {
            InitializeComponent();
            frmMainMenu.Content = new StartView();
            //About about = new About();
        }

        private void NewBtnMenu_Click(object sender, RoutedEventArgs e)
        {
            // Add pop up to prompt new spreadsheet
            //MessageBoxViewModel vmMessageBox = (MessageBoxViewModel)messageBox.FindResource("ViewModel");
            vmMessageBox = (MessageBoxViewModel)messageBox.FindResource("ViewModel");

            vmMessageBox.Caption = "Confirm New Record";
            vmMessageBox.Message = "A Record already exists.\nDo you want to replace it?";
            vmMessageBox.Ok = "Yes";
            vmMessageBox.Cancel = "No";
            //vmMessageBox.Yes = "Top";
            //vmMessageBox.No = "Bottom";
            vmMessageBox.OkVisibility = true;
            vmMessageBox.CancelVisibility = true;
            ///vm.YesVisibility = true;
            //vmMessageBox.NoVisibility = true;
            vmMessageBox.Image = new BitmapImage(new Uri(@"C:\Users\Robert Woodhouse\Google Drive\PowerUpp\Images\SaveIcon.ico"));

            // Center functionality
            vmMessageBox.Position = MessageBoxPosition.CenterOwner;
            vmMessageBox.Owner = this;

            //vmMessageBox.Show();

            Gat.Controls.MessageBoxResult result = vmMessageBox.Show();
            if (result == Gat.Controls.MessageBoxResult.Ok)
            {
                frmMainMenu.Content = new SelectionView();
                SelectionController.loadFile = false;
            }
        }

        private void ExitBtnMenu_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        private void AboutBtnMenu_Click(object sender, RoutedEventArgs e)
        {
            //About about = new About();
            about.ApplicationLogo = new BitmapImage(new Uri(@"C:\Users\Robert Woodhouse\Google Drive\PowerUpp\Images\SquatsIcon.ico"));
            about.Title = "Power Upp";
            about.Version = "v1.05";
            about.AdditionalNotes = "Power Upp is an application used to track your resistance exercise data over the course of time and present it visually as a chart";
            about.PublisherLogo = new BitmapImage(new Uri(@"C:\Users\Robert Woodhouse\Google Drive\PowerUpp\Images\DumbbellIcon.ico"));
            about.Copyright = "© 2019 Robert Woodhouse \nAll rights reserved";
            about.HyperlinkText = "https://github.com/robertwoodhouse/powerupp";
            about.Show();
            
            /*
            vmAbout.

            vmAbout.Window.Content = about;
            vmAbout.Window.Show();

            vmAbout = (AboutControlViewModel)about.FindResource("ViewModel");
            */
        }
    }
}
