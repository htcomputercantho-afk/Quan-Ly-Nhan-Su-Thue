using System.Windows;

namespace TaxPersonnelManagement.Views
{
    public partial class ConfirmDialog : Window
    {
        public ConfirmDialog(string message)
        {
            InitializeComponent();
            txtMessage.Text = message;
        }

        private void btnConfirm_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
        }
    }
}
