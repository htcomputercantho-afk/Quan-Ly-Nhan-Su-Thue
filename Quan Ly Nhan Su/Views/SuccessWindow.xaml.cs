using System.Windows;

namespace TaxPersonnelManagement.Views
{
    public partial class SuccessWindow : Window
    {
        public string FilePath { get; private set; }

        public SuccessWindow(string message = null, string filePath = null)
        {
            InitializeComponent();
            
            if (!string.IsNullOrEmpty(message))
            {
                txtMessage.Text = message;
            }

            if (!string.IsNullOrEmpty(filePath))
            {
                this.FilePath = filePath;
                pnlActions.Visibility = Visibility.Visible;
                txtSubMessage.Visibility = Visibility.Visible;
                txtSubMessage.Text = System.IO.Path.GetFileName(filePath);
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void OpenFile_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!string.IsNullOrEmpty(FilePath) && System.IO.File.Exists(FilePath))
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = FilePath,
                        UseShellExecute = true
                    });
                    this.Close();
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Không thể mở tệp: " + ex.Message);
            }
        }

        private void OpenFolder_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!string.IsNullOrEmpty(FilePath) && System.IO.File.Exists(FilePath))
                {
                    string folder = System.IO.Path.GetDirectoryName(FilePath);
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = folder,
                        UseShellExecute = true
                    });
                    this.Close();
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Không thể mở thư mục: " + ex.Message);
            }
        }
    }
}
