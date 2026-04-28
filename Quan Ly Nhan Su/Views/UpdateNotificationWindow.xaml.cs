using System.Windows;
using AutoUpdaterDotNET;

namespace TaxPersonnelManagement.Views
{
    public partial class UpdateNotificationWindow : Window
    {
        private readonly UpdateInfoEventArgs _args;

        public UpdateNotificationWindow(UpdateInfoEventArgs args)
        {
            InitializeComponent();
            _args = args;

            txtCurrentVersion.Text = args.InstalledVersion.ToString();
            txtNewVersion.Text = args.CurrentVersion.ToString();
            
            if (!string.IsNullOrEmpty(args.ChangelogURL))
            {
                // In a real app, you might fetch and show changelog, 
                // for now we use the description from XML or a default message
                txtChangelog.Text = "Hệ thống đã sẵn sàng bản cập nhật mới với nhiều cải tiến về giao diện và hiệu năng. Vui lòng cập nhật để có trải nghiệm tốt nhất.";
            }
        }

        private void btnLater_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            // Tự động sao lưu dữ liệu trước khi thực hiện cập nhật
            App.PerformBackup("auto_update");

            if (AutoUpdater.DownloadUpdate(_args))
            {
                Application.Current.Shutdown();
            }
        }

    }
}
