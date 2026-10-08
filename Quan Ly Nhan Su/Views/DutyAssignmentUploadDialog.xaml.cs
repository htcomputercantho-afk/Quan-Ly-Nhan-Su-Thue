using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Win32;
using TaxPersonnelManagement.Data;
using TaxPersonnelManagement.Models;
using TaxPersonnelManagement.Services;

namespace TaxPersonnelManagement.Views
{
    /// <summary>
    /// Hộp thoại tải lên (hoặc chỉnh sửa) một phiên bản phân công nhiệm vụ cho một Tổ.
    /// </summary>
    public partial class DutyAssignmentUploadDialog : Window
    {
        private const long MaxFileBytes = 50L * 1024 * 1024;

        private readonly int? _editId;
        private byte[]? _pickedBytes;
        private string? _pickedName;
        private bool _suppress = true;
        private bool _dateTouched;

        /// <param name="presetDepartment">Tổ chọn sẵn (nếu có)</param>
        /// <param name="year">Năm đang xem</param>
        /// <param name="month">Tháng chọn sẵn (nếu có)</param>
        /// <param name="editId">Id bản ghi cần sửa (null nếu thêm mới)</param>
        public DutyAssignmentUploadDialog(string? presetDepartment, int year, int? month = null, int? editId = null)
        {
            InitializeComponent();
            _editId = editId;

            foreach (var d in DutyAssignmentService.GetDepartments())
                cmbDepartment.Items.Add(d);

            for (int m = 1; m <= 12; m++)
                cmbMonth.Items.Add($"Tháng {m}");

            int baseYear = DateTime.Today.Year;
            for (int y = Math.Min(year, baseYear) - 3; y <= Math.Max(year, baseYear) + 2; y++)
                cmbYear.Items.Add(y);

            if (!string.IsNullOrWhiteSpace(presetDepartment))
                SelectDepartment(presetDepartment!);

            int initMonth = month ?? (year == DateTime.Today.Year ? DateTime.Today.Month : 1);
            cmbMonth.SelectedIndex = initMonth - 1;
            cmbYear.SelectedItem = year;

            if (_editId.HasValue)
                LoadForEdit(_editId.Value);
            else
                ApplyDefaultEffectiveDate();

            _suppress = false;
        }

        private void SelectDepartment(string name)
        {
            var match = cmbDepartment.Items.Cast<string>()
                .FirstOrDefault(x => string.Equals(x, name, StringComparison.OrdinalIgnoreCase));
            if (match == null)
            {
                cmbDepartment.Items.Add(name);
                match = name;
            }
            cmbDepartment.SelectedItem = match;
        }

        private void LoadForEdit(int id)
        {
            txtTitle.Text = "CHỈNH SỬA PHÂN CÔNG NHIỆM VỤ";
            var meta = DutyAssignmentService.GetVersions().FirstOrDefault(v => v.Id == id);
            if (meta == null) return;

            SelectDepartment(meta.Department);
            cmbMonth.SelectedIndex = meta.Month - 1;
            if (!cmbYear.Items.Contains(meta.Year)) cmbYear.Items.Add(meta.Year);
            cmbYear.SelectedItem = meta.Year;
            dpEffective.SelectedDate = meta.EffectiveDate;
            _dateTouched = true;
            txtDecisionTitle.Text = meta.Title ?? "";
            txtNote.Text = meta.Note ?? "";
            txtFileName.Text = meta.FileName;
            txtFileHint.Text = $"{meta.SizeText} • Chọn file khác nếu muốn thay thế (không bắt buộc)";
        }

        /// <summary>Ngày hiệu lực mặc định: hôm nay nếu là tháng hiện tại, ngược lại là ngày 1 của tháng được chọn.</summary>
        private void ApplyDefaultEffectiveDate()
        {
            if (cmbMonth.SelectedIndex < 0 || cmbYear.SelectedItem is not int y) return;
            int m = cmbMonth.SelectedIndex + 1;
            var today = DateTime.Today;
            dpEffective.SelectedDate = (y == today.Year && m == today.Month) ? today : new DateTime(y, m, 1);
        }

        private void MonthYear_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (_suppress || _dateTouched) return;
            _suppress = true;
            ApplyDefaultEffectiveDate();
            _suppress = false;
        }

        private void DpEffective_Changed(object? sender, SelectionChangedEventArgs e)
        {
            if (!_suppress) _dateTouched = true;
        }

        private void Header_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ButtonState == MouseButtonState.Pressed) DragMove();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        // ============================================================
        // Chọn / kéo thả file PDF
        // ============================================================
        private void BtnPickFile_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Title = "Chọn file PDF phân công nhiệm vụ",
                Filter = "Tệp PDF (*.pdf)|*.pdf",
                CheckFileExists = true
            };
            if (dlg.ShowDialog(this) == true)
                TryLoadPdf(dlg.FileName);
        }

        private void DropZone_DragOver(object sender, DragEventArgs e)
        {
            e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
            e.Handled = true;
        }

        private void DropZone_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetData(DataFormats.FileDrop) is string[] files && files.Length > 0)
                TryLoadPdf(files[0]);
        }

        private void TryLoadPdf(string path)
        {
            try
            {
                if (!string.Equals(Path.GetExtension(path), ".pdf", StringComparison.OrdinalIgnoreCase))
                {
                    ShowWarning("Chỉ chấp nhận file có định dạng PDF (.pdf).");
                    return;
                }

                var info = new FileInfo(path);
                if (info.Length == 0)
                {
                    ShowWarning("File PDF rỗng, vui lòng chọn file khác.");
                    return;
                }
                if (info.Length > MaxFileBytes)
                {
                    ShowWarning("File vượt quá dung lượng tối đa 50 MB.");
                    return;
                }

                var bytes = File.ReadAllBytes(path);
                // Kiểm tra chữ ký PDF "%PDF"
                if (bytes.Length < 4 || bytes[0] != 0x25 || bytes[1] != 0x50 || bytes[2] != 0x44 || bytes[3] != 0x46)
                {
                    ShowWarning("File không phải là PDF hợp lệ.");
                    return;
                }

                _pickedBytes = bytes;
                _pickedName = info.Name;
                txtFileName.Text = info.Name;
                txtFileHint.Text = $"{(info.Length >= 1024 * 1024 ? $"{info.Length / 1024d / 1024d:0.0} MB" : $"{Math.Max(1, info.Length / 1024)} KB")} • Đã chọn, bấm \"Lưu thông tin\" để hoàn tất";
            }
            catch (Exception ex)
            {
                ShowWarning($"Không đọc được file:\n{ex.Message}");
            }
        }

        // ============================================================
        // Lưu
        // ============================================================
        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            if (cmbDepartment.SelectedItem is not string department || string.IsNullOrWhiteSpace(department))
            {
                ShowWarning("Vui lòng chọn Tổ / bộ phận.");
                return;
            }
            if (cmbMonth.SelectedIndex < 0 || cmbYear.SelectedItem is not int year)
            {
                ShowWarning("Vui lòng chọn tháng và năm thay đổi.");
                return;
            }
            if (!dpEffective.SelectedDate.HasValue)
            {
                ShowWarning("Vui lòng chọn ngày có hiệu lực.");
                return;
            }
            if (!_editId.HasValue && _pickedBytes == null)
            {
                ShowWarning("Vui lòng chọn file PDF phân công nhiệm vụ.");
                return;
            }

            try
            {
                using var db = new AppDbContext();
                DutyAssignmentRecord record;
                if (_editId.HasValue)
                {
                    record = db.DutyAssignmentRecords.Find(_editId.Value)
                             ?? throw new InvalidOperationException("Bản ghi không còn tồn tại.");
                }
                else
                {
                    record = new DutyAssignmentRecord
                    {
                        UploadedAt = DateTime.Now,
                        UploadedBy = App.CurrentUser?.FullName
                    };
                    db.DutyAssignmentRecords.Add(record);
                }

                record.Department = department;
                record.Month = cmbMonth.SelectedIndex + 1;
                record.Year = year;
                record.EffectiveDate = dpEffective.SelectedDate.Value.Date;
                record.Title = string.IsNullOrWhiteSpace(txtDecisionTitle.Text) ? null : txtDecisionTitle.Text.Trim();
                record.Note = string.IsNullOrWhiteSpace(txtNote.Text) ? null : txtNote.Text.Trim();

                if (_pickedBytes != null)
                {
                    record.FileData = _pickedBytes;
                    record.FileName = _pickedName ?? "phan-cong-nhiem-vu.pdf";
                    record.FileSize = _pickedBytes.LongLength;
                }

                db.SaveChanges();
                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                DutyAssignmentNotificationDialog.ShowError(this, $"Lỗi khi lưu dữ liệu:\n{ex.Message}", "Lỗi lưu dữ liệu");
            }
        }

        private void ShowWarning(string message)
        {
            DutyAssignmentNotificationDialog.ShowWarning(this, message, "Lưu ý");
        }
    }
}
