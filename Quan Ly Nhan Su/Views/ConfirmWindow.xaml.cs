using System.Windows;

namespace TaxPersonnelManagement.Views
{
    public partial class ConfirmWindow : Window
    {
        public ConfirmWindow(string message, string title = "Xác Nhận")
        {
            InitializeComponent();
            txtMessage.Text = message;
            txtTitle.Text = title;
        }

        private void Yes_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
            this.Close();
        }

        private void No_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }
    }
}
