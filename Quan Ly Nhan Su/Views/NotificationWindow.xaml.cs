using System.Windows;

namespace TaxPersonnelManagement.Views
{
    public partial class NotificationWindow : Window
    {
        public NotificationWindow(string message, string title = "THÔNG BÁO")
        {
            InitializeComponent();
            txtMessage.Text = message;
            txtTitle.Text = title?.ToUpper() ?? "THÔNG BÁO";
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
