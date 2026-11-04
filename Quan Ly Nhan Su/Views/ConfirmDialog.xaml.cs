using System.Windows;

namespace TaxPersonnelManagement.Views
{
    public partial class ConfirmDialog : Window
    {
        public ConfirmDialog(string message, string title = "Xác Nhận Thao Tác", string confirmText = "XÁC NHẬN")
        {
            InitializeComponent();
            txtMessage.Text = message;
            if (!string.IsNullOrWhiteSpace(title))
            {
                txtTitle.Text = title;
            }
            if (!string.IsNullOrWhiteSpace(confirmText))
            {
                btnConfirm.Content = confirmText;
            }
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
