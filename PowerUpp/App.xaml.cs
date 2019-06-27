using System.Windows;

namespace PowerUpp
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        void AppExit(object sender, ExitEventArgs e)
        {
            SelectionController.StopExcelAppAsync();
        }
    }
}
