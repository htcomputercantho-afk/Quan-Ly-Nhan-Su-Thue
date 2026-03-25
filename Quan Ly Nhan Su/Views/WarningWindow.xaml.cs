using System.Windows;

namespace TaxPersonnelManagement.Views
{
    public partial class WarningWindow : Window
    {
        public WarningWindow(string message, string title = "Cảnh Báo")
        {
            InitializeComponent();
            txtMessage.Text = message;
            txtTitle.Text = title;
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
