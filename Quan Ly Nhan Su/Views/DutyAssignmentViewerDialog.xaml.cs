using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Web.WebView2.Core;
using Microsoft.Win32;
using TaxPersonnelManagement.Models;
using TaxPersonnelManagement.Services;

namespace TaxPersonnelManagement.Views
{
    /// <summary>
    /// Xem lịch sử phân công nhiệm vụ của một Tổ: danh sách các lần thay đổi + xem PDF trực tiếp, tải về, sửa, xóa.
    /// </summary>
    public partial class DutyAssignmentViewerDialog : Window
    {
        private readonly string _department;
        private readonly bool _canEdit;
        private readonly string _tempDir;
        private readonly Dictionary<int, string> _tempFiles = new();

        private int? _pendingSelectId;
        private bool _webReady;
        private DutyVersionItem? _current;
        private bool _loadingList;

        /// <summary>True nếu trong lúc xem có thêm / sửa / xóa dữ liệu (để màn hình chính làm mới).</summary>
        public bool Changed { get; private set; }

        public DutyAssignmentViewerDialog(string department, int? selectId = null)
        {
            InitializeComponent();
            _department = department;
            _pendingSelectId = selectId;
            _canEdit = App.CurrentUser?.Role != UserRole.Staff;
            _tempDir = Path.Combine(Path.GetTempPath(), "QLNS_PhanCong_" + Guid.NewGuid().ToString("N"));

            txtDepartment.Text = department;
            if (!_canEdit)
            {
                btnAddVersion.Visibility = Visibility.Collapsed;
                btnEdit.Visibility = Visibility.Collapsed;
                btnDelete.Visibility = Visibility.Collapsed;
            }

            Loaded += async (s, e) =>
            {
                await InitWebViewAsync();
                LoadVersions();
            };
            Closed += (s, e) => Cleanup();
        }

        private async Task InitWebViewAsync()
        {
            try
            {
                string userData = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "QuanLyNhanSu", "WebView2");
                var env = await CoreWebView2Environment.CreateAsync(null, userData);
                await webView.EnsureCoreWebView2Async(env);
                _webReady = true;
            }
            catch
            {
                _webReady = false;
            }
        }

        private void LoadVersions()
        {
            _loadingList = true;
            var versions = DutyAssignmentService.GetVersions(_department);
            lstVersions.ItemsSource = versions;
            txtCount.Text = versions.Count == 0 ? "Chưa có bản phân công nào" : $"{versions.Count} lần phân công";

            DutyVersionItem? toSelect = null;
            if (_pendingSelectId.HasValue)
                toSelect = versions.FirstOrDefault(v => v.Id == _pendingSelectId.Value);
            toSelect ??= versions.FirstOrDefault(v => v.IsCurrent) ?? versions.FirstOrDefault();
            _pendingSelectId = null;
            _loadingList = false;

            lstVersions.SelectedItem = toSelect;
            if (toSelect == null)
                ShowVersion(null);
        }

        private void LstVersions_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (_loadingList) return;
            ShowVersion(lstVersions.SelectedItem as DutyVersionItem);
        }

        private void ShowVersion(DutyVersionItem? item)
        {
            _current = item;
            bool has = item != null;
            btnDownload.IsEnabled = has;
            btnOpenExternal.IsEnabled = has;
            btnEdit.IsEnabled = has;
            btnDelete.IsEnabled = has;

            if (item == null)
            {
                txtVersionTitle.Text = "";
                txtVersionMeta.Text = "";
                SetPlaceholder("Tổ này chưa có file phân công nhiệm vụ.\nBấm \"Thêm bản mới\" để tải file PDF lên.", true);
                return;
            }

            txtVersionTitle.Text = $"{item.MonthText} — {item.DisplayTitle}";
            txtVersionMeta.Text = $"{item.EffectiveText}  •  {item.FileName} ({item.SizeText})  •  {item.UploadInfo}";

            try
            {
                string? path = EnsureTempFile(item);
                if (path == null)
                {
                    SetPlaceholder("Không đọc được nội dung file PDF.", true);
                    return;
                }

                if (_webReady)
                {
                    pnlPlaceholder.Visibility = Visibility.Collapsed;
                    webView.Visibility = Visibility.Visible;
                    webView.CoreWebView2.Navigate(new Uri(path).AbsoluteUri);
                }
                else
                {
                    SetPlaceholder("Không thể hiển thị PDF trực tiếp (thiếu Microsoft Edge WebView2 Runtime).\nHãy bấm \"Mở ngoài\" hoặc \"Tải về\" để xem file.", true);
                }
            }
            catch (Exception ex)
            {
                SetPlaceholder($"Lỗi khi mở file PDF:\n{ex.Message}", true);
            }
        }

        private void SetPlaceholder(string text, bool hideWeb)
        {
            if (hideWeb) webView.Visibility = Visibility.Collapsed;
            pnlPlaceholder.Visibility = Visibility.Visible;
            txtPlaceholder.Text = text;
        }

        /// <summary>Ghi nội dung PDF ra file tạm (một lần cho mỗi phiên bản) để WebView2 / ứng dụng ngoài mở.</summary>
        private string? EnsureTempFile(DutyVersionItem item)
        {
            if (_tempFiles.TryGetValue(item.Id, out var existing) && File.Exists(existing))
                return existing;

            var data = DutyAssignmentService.LoadFileData(item.Id);
            if (data == null || data.Length == 0) return null;

            Directory.CreateDirectory(_tempDir);
            string path = Path.Combine(_tempDir, $"phan-cong-{item.Id}-{DateTime.Now.Ticks}.pdf");
            File.WriteAllBytes(path, data);
            _tempFiles[item.Id] = path;
            return path;
        }

        // ============================================================
        // Thao tác
        // ============================================================
        private void BtnDownload_Click(object sender, RoutedEventArgs e)
        {
            if (_current == null) return;
            try
            {
                var data = DutyAssignmentService.LoadFileData(_current.Id);
                if (data == null) return;

                var dlg = new SaveFileDialog
                {
                    Title = "Tải file phân công nhiệm vụ",
                    Filter = "Tệp PDF (*.pdf)|*.pdf",
                    FileName = _current.FileName
                };
                if (dlg.ShowDialog(this) == true)
                {
                    File.WriteAllBytes(dlg.FileName, data);
                    DutyAssignmentNotificationDialog.ShowSuccess(this, $"Đã lưu file PDF thành công tại:\n{dlg.FileName}", "Tải về thành công");
                }
            }
            catch (Exception ex)
            {
                DutyAssignmentNotificationDialog.ShowError(this, $"Không thể lưu file PDF:\n{ex.Message}", "Lỗi tải về");
            }
        }

        private void BtnOpenExternal_Click(object sender, RoutedEventArgs e)
        {
            if (_current == null) return;
            try
            {
                string? path = EnsureTempFile(_current);
                if (path != null)
                    Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                DutyAssignmentNotificationDialog.ShowError(this, $"Không thể mở file bằng ứng dụng ngoài:\n{ex.Message}", "Lỗi mở file");
            }
        }

        private void BtnAddVersion_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new DutyAssignmentUploadDialog(_department, DateTime.Today.Year) { Owner = this };
            if (dlg.ShowDialog() == true)
            {
                Changed = true;
                var all = DutyAssignmentService.GetVersions(_department);
                _pendingSelectId = all.OrderByDescending(v => v.Id).FirstOrDefault()?.Id;
                LoadVersions();
            }
        }

        private void BtnEdit_Click(object sender, RoutedEventArgs e)
        {
            if (_current == null) return;
            int id = _current.Id;
            var dlg = new DutyAssignmentUploadDialog(_current.Department, _current.Year, _current.Month, id) { Owner = this };
            if (dlg.ShowDialog() == true)
            {
                Changed = true;
                _tempFiles.Remove(id);
                _pendingSelectId = id;
                LoadVersions();
            }
        }

        private void BtnDelete_Click(object sender, RoutedEventArgs e)
        {
            if (_current == null) return;

            bool agreed = DutyAssignmentNotificationDialog.ShowConfirm(
                this,
                $"Bạn có chắc chắn muốn XÓA bản phân công nhiệm vụ:\n\n{_current.MonthText} — {_current.DisplayTitle}\nTổ: {_current.Department}?\n\nThao tác này sẽ xóa vĩnh viễn tệp PDF và dữ liệu liên quan.",
                "Xác nhận xóa văn bản",
                isDestructive: true,
                confirmButtonText: "ĐỒNG Ý XÓA");

            if (!agreed) return;

            try
            {
                DutyAssignmentService.Delete(_current.Id);
                Changed = true;
                DutyAssignmentNotificationDialog.ShowSuccess(this, "Đã xóa bản phân công nhiệm vụ thành công.", "Đã xóa");
                LoadVersions();
            }
            catch (Exception ex)
            {
                DutyAssignmentNotificationDialog.ShowError(this, $"Lỗi khi xóa:\n{ex.Message}", "Lỗi thao tác");
            }
        }

        private void Cleanup()
        {
            try { webView.Dispose(); } catch { }
            try
            {
                if (Directory.Exists(_tempDir))
                    Directory.Delete(_tempDir, true);
            }
            catch { /* file tạm có thể còn bị khóa — bỏ qua */ }
        }
    }
}
