using Gat.Controls;
using System;
using System.IO;
using System.Windows.Controls;
using System.Windows.Media.Imaging;

namespace PowerUpp
{
    class MessageBoxController
    {
        About about = new About();
        MessageBoxView messageBox = new MessageBoxView();
        MessageBoxViewModel vmMessageBox = new MessageBoxViewModel();


        public void NewMenu(Frame frame, MainWindow window)
        {
            vmMessageBox = (MessageBoxViewModel)messageBox.FindResource("ViewModel");

            vmMessageBox.Caption = "Confirm New Record";
            vmMessageBox.Message = "A Record already exists.\nDo you want to replace it?";
            vmMessageBox.Ok = "Yes";
            vmMessageBox.Cancel = "No";
            vmMessageBox.OkVisibility = true;
            vmMessageBox.CancelVisibility = true;
            vmMessageBox.Image = new BitmapImage(new Uri(IconURI("Images", "SaveIcon.ico")));


            // Center functionality
            vmMessageBox.Position = MessageBoxPosition.CenterOwner;
            vmMessageBox.Owner = window;

            Gat.Controls.MessageBoxResult result = vmMessageBox.Show();
            if (result == Gat.Controls.MessageBoxResult.Ok)
            {
                frame.Content = new SelectionView();
                SelectionController.loadFile = false;
            }
        }

        public void AboutMenu()
        {
            about.ApplicationLogo = new BitmapImage(new Uri(IconURI("Images", "SquatsIcon.ico")));
            about.Title = "Power Upp";
            about.Version = "v1.07";
            about.AdditionalNotes = "Power Upp is an application used to track your resistance exercise data over the course of time and present it visually as a chart";
            about.PublisherLogo = new BitmapImage(new Uri(IconURI("Images", "DumbbellIcon.ico")));
            about.Copyright = "© 2019 Robert Woodhouse \nAll rights reserved";
            about.HyperlinkText = "https://github.com/robertwoodhouse/powerupp";
            about.Show();
        }

        private string IconURI(string folder, string ico)
        {
            var currentDirectory = Directory.GetCurrentDirectory();
            return Path.Combine(currentDirectory, folder, ico);
        }
    }
}
