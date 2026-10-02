using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using TaxPersonnelManagement.Data;
using TaxPersonnelManagement.Models;

namespace TaxPersonnelManagement.Views
{
    public partial class RotationRecordDialog : Window
    {
        private RotationRecord? _existingRecord;
        private Personnel? _selectedPersonnel;
        private readonly string _defaultPlanType;

        /// <summary>
        /// Tạo dialog thêm mới hoặc chỉnh sửa bản ghi luân chuyển/điều động.
        /// </summary>
        /// <param name="record">Bản ghi cần chỉnh sửa (null nếu thêm mới)</param>
        /// <param name="defaultPlanType">Loại kế hoạch mặc định cho bản ghi mới</param>
        public RotationRecordDialog(RotationRecord? record, string defaultPlanType = "Trong kế hoạch")
        {
            InitializeComponent();
            _existingRecord = record;
            _defaultPlanType = defaultPlanType;

            LoadDepartments();

            if (record != null)
            {
                // Chế độ chỉnh sửa
                txtTitle.Text = "Chỉnh sửa Luân chuyển / Điều động";
                btnDelete.Visibility = Visibility.Visible;
                FillForm(record);
            }
            else
            {
                // Chế độ thêm mới — đặt kế hoạch mặc định
                SetComboByText(cmbPlanType, defaultPlanType);
            }
        }

        // ============================================================
        // Khởi tạo form
        // ============================================================
        private void LoadDepartments()
        {
            try
            {
                using var db = new AppDbContext();
                var depts = db.Departments.OrderBy(d => d.Name).Select(d => d.Name).ToList();

                cmbFromDepartment.Items.Clear();
                cmbToDepartment.Items.Clear();
                foreach (var d in depts)
                {
                    cmbFromDepartment.Items.Add(d);
                    cmbToDepartment.Items.Add(d);
                }
            }
            catch { /* Bỏ qua lỗi tải bộ phận */ }
        }

        private void FillForm(RotationRecord r)
        {
            // Tải thông tin cán bộ
            try
            {
                using var db = new AppDbContext();
                _selectedPersonnel = db.Personnel.Find(r.PersonnelId);
            }
            catch { }

            UpdatePersonnelDisplay();

            SetComboByText(cmbRotationType, r.RotationType);
            SetComboByText(cmbPlanType, r.PlanType);

            cmbFromDepartment.Text = r.FromDepartment ?? "";
            cmbToDepartment.Text = r.ToDepartment ?? "";

            chkIsCompleted.IsChecked = r.IsCompleted;
            if (r.IsCompleted)
            {
                txtDecisionNumber.Text = r.DecisionNumber ?? "";
                dpDecisionDate.SelectedDate = r.DecisionDate;
                dpEffectiveDate.SelectedDate = r.EffectiveDate;
                pnlDecision.Visibility = Visibility.Visible;
            }

            txtNote.Text = r.Note ?? "";
        }

        private void UpdatePersonnelDisplay()
        {
            if (_selectedPersonnel == null)
            {
                txtPersonnelName.Text = "";
                pnlPersonnelInfo.Visibility = Visibility.Collapsed;
                return;
            }

            txtPersonnelName.Text = _selectedPersonnel.FullName;
            txtPersonnelDetails.Text =
                $"CCCD: {_selectedPersonnel.IdentityCardNumber ?? "---"}  |  " +
                $"Ngày sinh: {_selectedPersonnel.DateOfBirth?.ToString("dd/MM/yyyy") ?? "---"}  |  " +
                $"Bộ phận: {_selectedPersonnel.Department ?? "---"}";
            pnlPersonnelInfo.Visibility = Visibility.Visible;

            // Tự động điền bộ phận đang công tác vào cmbFromDepartment nếu chưa có
            if (string.IsNullOrEmpty(cmbFromDepartment.Text) && !string.IsNullOrEmpty(_selectedPersonnel.Department))
                cmbFromDepartment.Text = _selectedPersonnel.Department;
        }

        // ============================================================
        // Sự kiện UI
        // ============================================================
        private void BtnSelectPersonnel_Click(object sender, RoutedEventArgs e)
        {
            var selector = new PersonnelSelectorDialog(new System.Collections.Generic.List<int>(), isSingleSelect: true);
            selector.Owner = this;
            if (selector.ShowDialog() == true && selector.SelectedPersonnelIds.Count > 0)
            {
                int selectedId = selector.SelectedPersonnelIds[0];
                using var db = new AppDbContext();
                _selectedPersonnel = db.Personnel.Find(selectedId);
                UpdatePersonnelDisplay();
            }
        }

        private void ChkIsCompleted_Changed(object sender, RoutedEventArgs e)
        {
            pnlDecision.Visibility = chkIsCompleted.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            // Validate
            if (_selectedPersonnel == null)
            {
                var warning = new WarningWindow("Vui lòng chọn cán bộ công chức!", "Thiếu thông tin");
                warning.Owner = this;
                warning.ShowDialog();
                return;
            }

            if (chkIsCompleted.IsChecked == true && string.IsNullOrWhiteSpace(txtDecisionNumber.Text))
            {
                var warning = new WarningWindow("Vui lòng nhập số quyết định khi đã thực hiện điều động!", "Thiếu thông tin");
                warning.Owner = this;
                warning.ShowDialog();
                return;
            }

            try
            {
                using var db = new AppDbContext();

                if (_existingRecord == null)
                {
                    // Thêm mới
                    var newRecord = BuildRecord();
                    db.RotationRecords.Add(newRecord);
                }
                else
                {
                    // Cập nhật
                    var record = db.RotationRecords.Find(_existingRecord.Id);
                    if (record == null)
                    {
                        var warning = new WarningWindow("Không tìm thấy bản ghi để cập nhật.", "Lỗi");
                        warning.Owner = this;
                        warning.ShowDialog();
                        return;
                    }
                    ApplyToRecord(record);
                }

                db.SaveChanges();
                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                var warning = new WarningWindow($"Lỗi khi lưu dữ liệu:\n{ex.Message}", "Lỗi");
                warning.Owner = this;
                warning.ShowDialog();
            }
        }

        private void BtnDelete_Click(object sender, RoutedEventArgs e)
        {
            if (_existingRecord == null) return;

            var confirm = new ConfirmWindow(
                $"Bạn có chắc chắn muốn XÓA bản ghi luân chuyển/điều động của:\n\n{_selectedPersonnel?.FullName ?? "cán bộ này"}?",
                "Xác nhận xóa");
            confirm.Owner = this;
            if (confirm.ShowDialog() != true) return;

            try
            {
                using var db = new AppDbContext();
                var record = db.RotationRecords.Find(_existingRecord.Id);
                if (record != null) db.RotationRecords.Remove(record);
                db.SaveChanges();
                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                var warning = new WarningWindow($"Lỗi khi xóa bản ghi:\n{ex.Message}", "Lỗi");
                warning.Owner = this;
                warning.ShowDialog();
            }
        }

        // ============================================================
        // Helpers
        // ============================================================
        private RotationRecord BuildRecord()
        {
            var r = new RotationRecord { PersonnelId = _selectedPersonnel!.Id };
            ApplyToRecord(r);
            return r;
        }

        private void ApplyToRecord(RotationRecord r)
        {
            r.PersonnelId       = _selectedPersonnel!.Id;
            r.RotationType      = GetComboText(cmbRotationType);
            r.PlanType          = GetComboText(cmbPlanType);
            r.FromDepartment    = cmbFromDepartment.Text.Trim().NullIfEmpty();
            r.ToDepartment      = cmbToDepartment.Text.Trim().NullIfEmpty();
            r.IsCompleted       = chkIsCompleted.IsChecked == true;
            r.DecisionNumber    = r.IsCompleted ? txtDecisionNumber.Text.Trim().NullIfEmpty() : null;
            r.DecisionDate      = r.IsCompleted ? dpDecisionDate.SelectedDate : null;
            r.EffectiveDate     = r.IsCompleted ? dpEffectiveDate.SelectedDate : null;
            r.Note              = txtNote.Text.Trim().NullIfEmpty();
        }

        private string GetComboText(ComboBox combo)
        {
            if (combo.SelectedItem is ComboBoxItem item)
                return item.Content?.ToString() ?? "";
            return combo.Text ?? "";
        }

        private void SetComboByText(ComboBox combo, string text)
        {
            foreach (ComboBoxItem item in combo.Items)
            {
                if (item.Content?.ToString() == text)
                {
                    combo.SelectedItem = item;
                    return;
                }
            }
        }
    }

    /// <summary>
    /// Extension để chuyển chuỗi rỗng thành null
    /// </summary>
    internal static class StringExtensions
    {
        public static string? NullIfEmpty(this string? s)
            => string.IsNullOrWhiteSpace(s) ? null : s;
    }
}
