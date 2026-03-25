using System.Windows;

namespace TaxPersonnelManagement.Views
{
    public partial class SuccessWindow : Window
    {
        public SuccessWindow(string message = null)
        {
            InitializeComponent();
            if (!string.IsNullOrEmpty(message))
            {
                txtMessage.Text = message;
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
