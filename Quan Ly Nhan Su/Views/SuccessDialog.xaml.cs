using System.Windows;

namespace TaxPersonnelManagement.Views
{
    public partial class SuccessDialog : Window
    {
        public SuccessDialog(string message)
        {
            InitializeComponent();
            txtMessage.Text = message;
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
