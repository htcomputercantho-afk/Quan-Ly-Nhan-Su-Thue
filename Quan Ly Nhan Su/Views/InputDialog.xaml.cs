using System.Windows;

namespace TaxPersonnelManagement.Views
{
    public partial class InputDialog : Window
    {
        public string InputText { get; private set; } = string.Empty;

        public InputDialog()
        {
            InitializeComponent();
            txtInput.Focus();
        }

        private void btnConfirm_Click(object sender, RoutedEventArgs e)
        {
            string val = txtInput.Text.Trim();
            if (string.IsNullOrEmpty(val))
            {
                MessageBox.Show("Vui lòng nhập giá trị nhiệm kỳ!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            InputText = val;
            DialogResult = true;
            Close();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
