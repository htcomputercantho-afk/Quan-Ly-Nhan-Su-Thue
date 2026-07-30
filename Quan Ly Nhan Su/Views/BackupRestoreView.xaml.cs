using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Win32;
using TaxPersonnelManagement.Services;

namespace TaxPersonnelManagement.Views
{
    public partial class BackupRestoreView : UserControl
    {
        private string _dbPath;
        private string? _selectedRestoreFile;
        private readonly GoogleDriveSyncService _driveService = App.DriveSync;

        public BackupRestoreView()
        {
            InitializeComponent();
            _dbPath = Path.Combine(System.AppContext.BaseDirectory, "tax_personnel.db");
            LoadDbInfo();
            // Tự động tải token đã lưu (không mở trình duyệt)
            _ = TryAutoLoadDriveTokenAsync();
        }

        /// <summary>
        /// Hiển thị thông tin CSDL hiện tại (đường dẫn, kích thước).
        /// </summary>
        private void LoadDbInfo()
        {
            txtDbPath.Text = _dbPath;

            if (File.Exists(_dbPath))
            {
                var fileInfo = new FileInfo(_dbPath);
                double sizeKb = fileInfo.Length / 1024.0;
                if (sizeKb >= 1024)
                {
                    txtDbSize.Text = $"{sizeKb / 1024.0:F2} MB";
                }
                else
                {
                    txtDbSize.Text = $"{sizeKb:F1} KB";
                }
            }
            else
            {
                txtDbSize.Text = "Không tìm thấy file CSDL";
            }
        }

        /// <summary>
        /// Sao lưu CSDL hiện tại ra file do người dùng chọn.
        /// </summary>
        private void btnBackup_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!File.Exists(_dbPath))
                {
                    ShowWarning("Không tìm thấy file cơ sở dữ liệu!", "Lỗi");
                    return;
                }

                var dlg = new SaveFileDialog
                {
                    Title = "Chọn nơi lưu bản sao lưu",
                    FileName = $"QLNS_Backup_{DateTime.Now:yyyyMMdd_HHmmss}.db",
                    DefaultExt = ".db",
                    Filter = "SQLite Database (.db)|*.db"
                };

                if (dlg.ShowDialog() == true)
                {
                    File.Copy(_dbPath, dlg.FileName, overwrite: true);

                    var success = new SuccessWindow("Sao lưu thành công!", null, dlg.FileName, true);
                    success.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                ShowWarning($"Lỗi khi sao lưu: {ex.Message}", "Lỗi Sao Lưu");
            }
        }

        /// <summary>
        /// Chọn file sao lưu để phục hồi.
        /// </summary>
        private void btnSelectFile_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Title = "Chọn file sao lưu để phục hồi",
                DefaultExt = ".db",
                Filter = "SQLite Database (.db)|*.db"
            };

            if (dlg.ShowDialog() == true)
            {
                _selectedRestoreFile = dlg.FileName;
                var fi = new FileInfo(_selectedRestoreFile);
                double sizeKb = fi.Length / 1024.0;
                string sizeText = sizeKb >= 1024 ? $"{sizeKb / 1024.0:F2} MB" : $"{sizeKb:F1} KB";

                txtSelectedFile.Text = $"{fi.Name} ({sizeText}) - {fi.LastWriteTime:dd/MM/yyyy HH:mm}";
                txtSelectedFile.FontStyle = FontStyles.Normal;
                txtSelectedFile.Foreground = System.Windows.Media.Brushes.Black;
                btnRestore.IsEnabled = true;
            }
        }

        /// <summary>
        /// Phục hồi CSDL từ file sao lưu đã chọn.
        /// Tự động sao lưu CSDL hiện tại trước khi ghi đè.
        /// Sau khi phục hồi, ứng dụng sẽ tự khởi động lại.
        /// </summary>
        private void btnRestore_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(_selectedRestoreFile) || !File.Exists(_selectedRestoreFile))
                {
                    ShowWarning("Vui lòng chọn file sao lưu hợp lệ!", "Thông Báo");
                    return;
                }

                // Xác nhận trước khi phục hồi
                var confirm = new ConfirmDialog("⚠ Bạn có chắc muốn phục hồi CSDL?\n\nToàn bộ dữ liệu hiện tại sẽ bị thay thế bằng dữ liệu từ file sao lưu.\nHệ thống sẽ tự động sao lưu CSDL hiện tại trước khi phục hồi.\n\nSau khi phục hồi, ứng dụng sẽ tự khởi động lại.");
                if (confirm.ShowDialog() != true) return;

                // Tự động sao lưu CSDL hiện tại trước khi ghi đè
                string autoBackupPath = Path.Combine(
                    Path.GetDirectoryName(_dbPath)!,
                    $"tax_personnel_auto_backup_{DateTime.Now:yyyyMMdd_HHmmss}.db");

                if (File.Exists(_dbPath))
                {
                    File.Copy(_dbPath, autoBackupPath, overwrite: true);
                }

                // Clear all SQLite connection pools to release file locks
                Microsoft.Data.Sqlite.SqliteConnection.ClearAllPools();
                GC.Collect();
                GC.WaitForPendingFinalizers();

                // Ghi đè CSDL hiện tại bằng file sao lưu
                File.Copy(_selectedRestoreFile, _dbPath, overwrite: true);

                // Xóa các file WAL và SHM nếu có, để tránh lỗi sai lệch dữ liệu do bộ nhớ đệm của SQLite
                string walPath = _dbPath + "-wal";
                string shmPath = _dbPath + "-shm";
                if (File.Exists(walPath)) File.Delete(walPath);
                if (File.Exists(shmPath)) File.Delete(shmPath);

                var success = new SuccessWindow("Phục hồi dữ liệu thành công!\nỨng dụng sẽ tự động khởi động lại.", null, autoBackupPath, showActions: false);
                success.ShowDialog();

                // Khởi động lại ứng dụng
                var exePath = System.Diagnostics.Process.GetCurrentProcess().MainModule?.FileName;
                if (!string.IsNullOrEmpty(exePath))
                {
                    System.Diagnostics.Process.Start(exePath);
                }
                Application.Current.Shutdown();
            }
            catch (Exception ex)
            {
                ShowWarning($"Lỗi khi phục hồi: {ex.Message}", "Lỗi Phục Hồi");
            }
        }

        // ─────────────────────────────────────────────────────────────────────
        // GOOGLE DRIVE SYNC
        // ─────────────────────────────────────────────────────────────────────

        private async Task TryAutoLoadDriveTokenAsync()
        {
            bool loaded = await _driveService.TryLoadSavedTokenAsync();
            UpdateDriveStatusUI(loaded);
            if (loaded)
            {
                var cloudTime = await _driveService.GetCloudModifiedTimeAsync();
                if (cloudTime.HasValue)
                    txtLastSync.Text = $"Drive: {cloudTime.Value:dd/MM/yyyy HH:mm}";
            }
        }

        private void UpdateDriveStatusUI(bool connected)
        {
            if (connected)
            {
                ellipseStatus.Fill     = new SolidColorBrush(Color.FromRgb(46, 125, 50));
                txtDriveStatus.Text    = "Đã kết nối";
                txtDriveEmail.Text     = _driveService.GetConnectedEmail() ?? "";
                txtDriveConnect.Text   = "Ngắt kết nối";
                iconDriveConnect.Kind  = MaterialDesignThemes.Wpf.PackIconKind.LinkVariantOff;
                btnDriveConnect.Background = new SolidColorBrush(Color.FromRgb(198, 40, 40));
                btnDrivePush.IsEnabled = true;
                btnDrivePull.IsEnabled = true;
                borderDriveStatus.Background = new SolidColorBrush(Color.FromRgb(232, 245, 233));
            }
            else
            {
                ellipseStatus.Fill     = new SolidColorBrush(Color.FromRgb(158, 158, 158));
                txtDriveStatus.Text    = "Chưa kết nối";
                txtDriveEmail.Text     = "";
                txtLastSync.Text       = "";
                txtDriveConnect.Text   = "Kết nối Google Drive";
                iconDriveConnect.Kind  = MaterialDesignThemes.Wpf.PackIconKind.LinkVariant;
                btnDriveConnect.Background = new SolidColorBrush(Color.FromRgb(21, 101, 192));
                btnDrivePush.IsEnabled = false;
                btnDrivePull.IsEnabled = false;
                borderDriveStatus.Background = new SolidColorBrush(Color.FromRgb(236, 239, 241));
            }
        }

        private async void btnDriveConnect_Click(object sender, RoutedEventArgs e)
        {
            if (!_driveService.IsConfigured)
            {
                ShowWarning("Tính năng Google Drive chưa được cấu hình.\nVui lòng liên hệ nhà phát triển để được hỗ trợ.", "Chưa Cấu Hình");
                return;
            }

            if (_driveService.IsConnected)
            {
                // Ngắt kết nối
                var confirm = new ConfirmDialog("Bạn có muốn ngắt kết nối Google Drive?\nToken xác thực sẽ bị xóa khỏi máy này.");
                if (confirm.ShowDialog() != true) return;
                _driveService.Disconnect();
                UpdateDriveStatusUI(false);
                return;
            }

            // Kết nối mới
            try
            {
                btnDriveConnect.IsEnabled = false;
                txtDriveConnect.Text = "Đang mở trình duyệt...";
                await _driveService.ConnectAsync();
                UpdateDriveStatusUI(true);
                var cloudTime = await _driveService.GetCloudModifiedTimeAsync();
                if (cloudTime.HasValue)
                    txtLastSync.Text = $"Drive: {cloudTime.Value:dd/MM/yyyy HH:mm}";
            }
            catch (Exception ex)
            {
                ShowWarning($"Không thể kết nối Google Drive:\n{ex.Message}", "Lỗi Kết Nối");
                UpdateDriveStatusUI(false);
            }
            finally
            {
                btnDriveConnect.IsEnabled = true;
            }
        }

        private async void btnDrivePush_Click(object sender, RoutedEventArgs e)
        {
            await RunDriveSyncAsync(isUpload: true);
        }

        private async void btnDrivePull_Click(object sender, RoutedEventArgs e)
        {
            // Kiểm tra xem trên Drive đã có bản sao lưu nào chưa
            panelDriveLoading.Visibility = Visibility.Visible;
            txtDriveLoading.Text = "Đang kiểm tra dữ liệu trên Drive...";
            btnDrivePush.IsEnabled = false;
            btnDrivePull.IsEnabled = false;

            bool hasBackup = false;
            try
            {
                hasBackup = await _driveService.HasCloudBackupAsync();
            }
            finally
            {
                panelDriveLoading.Visibility = Visibility.Collapsed;
                btnDrivePush.IsEnabled = _driveService.IsConnected;
                btnDrivePull.IsEnabled = _driveService.IsConnected;
            }

            if (!hasBackup)
            {
                ShowWarning("Không tìm thấy bản sao lưu nào trên Google Drive!\n\nVui lòng bấm nút 'Đẩy lên' trước để tạo bản sao lưu dữ liệu.", "Chưa Có Bản Sao Lưu");
                return;
            }

            var confirm = new ConfirmDialog(
                "Tải dữ liệu từ Google Drive về sẽ GHI ĐÈ CSDL hiện tại.\n" +
                "Bản backup cục bộ sẽ được tự động tạo trước khi ghi đè.\n\nBạn có chắc muốn tiếp tục?");
            if (confirm.ShowDialog() != true) return;

            await RunDriveSyncAsync(isUpload: false);
        }

        private async Task RunDriveSyncAsync(bool isUpload)
        {
            panelDriveLoading.Visibility = Visibility.Visible;
            txtDriveLoading.Text = isUpload ? "Đang tải lên Drive..." : "Đang tải về từ Drive...";
            btnDrivePush.IsEnabled = false;
            btnDrivePull.IsEnabled = false;
            btnDriveConnect.IsEnabled = false;

            try
            {
                bool success;
                if (isUpload)
                {
                    success = await _driveService.PushAsync();
                }
                else
                {
                    // Release SQLite locks trước khi ghi đè
                    Microsoft.Data.Sqlite.SqliteConnection.ClearAllPools();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    success = await _driveService.PullAsync();
                }

                if (success)
                {
                    var cloudTime = await _driveService.GetCloudModifiedTimeAsync();
                    if (cloudTime.HasValue)
                        txtLastSync.Text = $"Đồng bộ lúc: {cloudTime.Value:dd/MM/yyyy HH:mm}";

                    if (isUpload)
                    {
                        var successWin = new SuccessWindow("Tải lên Google Drive thành công!", "Dữ liệu CSDL đã được sao lưu an toàn trên Google Drive.");
                        if (Window.GetWindow(this) is Window parentWin) successWin.Owner = parentWin;
                        successWin.ShowDialog();
                    }
                    else
                    {
                        var successWin = new SuccessWindow("Tải về từ Google Drive thành công!", "Ứng dụng sẽ tự động khởi động lại để áp dụng CSDL mới.");
                        if (Window.GetWindow(this) is Window parentWin) successWin.Owner = parentWin;
                        successWin.ShowDialog();
                    }

                    if (!isUpload)
                    {
                        var exePath = System.Diagnostics.Process.GetCurrentProcess().MainModule?.FileName;
                        if (!string.IsNullOrEmpty(exePath))
                            System.Diagnostics.Process.Start(exePath);
                        Application.Current.Shutdown();
                    }
                }
                else
                {
                    ShowWarning(isUpload ? "Tải lên thất bại. Vui lòng kiểm tra kết nối mạng." : "Tải về thất bại. Vui lòng kiểm tra kết nối mạng.", "Lỗi Đồng Bộ");
                }
            }
            catch (Exception ex)
            {
                ShowWarning($"Lỗi khi đồng bộ: {ex.Message}", "Lỗi Đồng Bộ");
            }
            finally
            {
                panelDriveLoading.Visibility = Visibility.Collapsed;
                btnDrivePush.IsEnabled = _driveService.IsConnected;
                btnDrivePull.IsEnabled = _driveService.IsConnected;
                btnDriveConnect.IsEnabled = true;
            }
        }

        private void ShowWarning(string message, string title = "Cảnh Báo")
        {
            var win = new WarningWindow(message, title);
            if (Window.GetWindow(this) is Window parentWin) win.Owner = parentWin;
            win.ShowDialog();
        }
    }
}
