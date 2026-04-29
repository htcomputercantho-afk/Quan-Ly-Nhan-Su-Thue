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
            
            this.Loaded += UpdateNotificationWindow_Loaded;
        }

        private async void UpdateNotificationWindow_Loaded(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(_args.ChangelogURL))
            {
                try 
                {
                    txtChangelog.Text = "Đang tải thông tin bản cập nhật...";
                    using var client = new System.Net.Http.HttpClient();
                    string url = _args.ChangelogURL;
                    if (!url.Contains("?")) url += $"?t={System.DateTime.Now.Ticks}";
                    
                    var changelogText = await client.GetStringAsync(url);
                    if (!string.IsNullOrWhiteSpace(changelogText))
                    {
                        txtChangelog.Text = changelogText;
                    }
                    else
                    {
                        txtChangelog.Text = "Không có thông tin chi tiết cho bản cập nhật này.";
                    }
                }
                catch 
                {
                    txtChangelog.Text = "Hệ thống đã sẵn sàng bản cập nhật mới với nhiều cải tiến. Vui lòng cập nhật để trải nghiệm.";
                }
            }
            else
            {
                txtChangelog.Text = "Hệ thống đã sẵn sàng bản cập nhật mới với nhiều cải tiến.";
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
