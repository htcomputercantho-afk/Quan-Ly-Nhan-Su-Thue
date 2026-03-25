using System.Windows;

namespace TaxPersonnelManagement.Views
{
    public partial class NotificationWindow : Window
    {
        public NotificationWindow(string message, string title = "Thông báo")
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
