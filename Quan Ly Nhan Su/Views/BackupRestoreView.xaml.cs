using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;

namespace TaxPersonnelManagement.Views
{
    public partial class BackupRestoreView : UserControl
    {
        private string _dbPath;
        private string? _selectedRestoreFile;

        public BackupRestoreView()
        {
            InitializeComponent();
            _dbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tax_personnel.db");
            LoadDbInfo();
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
                    MessageBox.Show("Không tìm thấy file cơ sở dữ liệu!", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
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

                    var success = new SuccessWindow("Sao lưu thành công!", dlg.FileName);
                    success.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi sao lưu: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
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
                    MessageBox.Show("Vui lòng chọn file sao lưu hợp lệ!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
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

                // Ghi đè CSDL hiện tại bằng file sao lưu
                File.Copy(_selectedRestoreFile, _dbPath, overwrite: true);

                var success = new SuccessWindow("Phục hồi thành công! Ứng dụng sẽ khởi động lại...", autoBackupPath);
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
                MessageBox.Show($"Lỗi khi phục hồi: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
