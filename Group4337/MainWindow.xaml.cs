using System.Windows;

namespace Group4337
{
    public partial class MainWindow : Window
    {
        public MainWindow()
            => InitializeComponent();

        private void _4337_Titova_Click(object sender, RoutedEventArgs e)
        {
            _4337_Titova infoWindow = new _4337_Titova();
            infoWindow.ShowDialog();
        }
    }
}