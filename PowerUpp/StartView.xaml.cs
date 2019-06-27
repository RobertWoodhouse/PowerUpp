using System.Windows;
using System.Windows.Controls;
using System.Windows.Navigation;

namespace PowerUpp
{
    /// <summary>
    /// Interaction logic for StartView.xaml
    /// </summary>
    public partial class StartView : Page
    {
        public StartView()
        {
            InitializeComponent();
        }

        private void btnStart_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Content = new SelectionView();
            SelectionController.loadFile = true;
        }
    }
}
