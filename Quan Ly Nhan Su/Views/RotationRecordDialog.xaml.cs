using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
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
        /// <param name="defaultYear">Năm kế hoạch mặc định (nếu có)</param>
        public RotationRecordDialog(RotationRecord? record, string defaultPlanType = "Trong kế hoạch", int? defaultYear = null)
        {
            InitializeComponent();
            _existingRecord = record;
            _defaultPlanType = defaultPlanType;

            LoadDepartments();

            int initialYear = defaultYear ?? DateTime.Now.Year;
            SetComboByText(cmbPlanYear, initialYear.ToString());

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
                UpdatePersonnelDisplay();
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
                var depts = db.Departments
                    .AsEnumerable()
                    .OrderBy(d => DepartmentSorter.GetSortKey(d.Name).Order)
                    .ThenBy(d => DepartmentSorter.GetSortKey(d.Name).Number)
                    .ThenBy(d => DepartmentSorter.GetSortKey(d.Name).Name)
                    .Select(d => d.Name)
                    .ToList();

                cmbFromDepartment.Items.Clear();
                cmbToDepartment.Items.Clear();

                cmbToDepartment.Items.Add("-- Chưa xác định / Tùy chọn --");

                foreach (var d in depts)
                {
                    cmbFromDepartment.Items.Add(d);
                    cmbToDepartment.Items.Add(d);
                }

                cmbToDepartment.SelectedIndex = 0;
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

            int recYear = r.DecisionDate?.Year ?? r.EffectiveDate?.Year ?? DateTime.Now.Year;
            SetComboByText(cmbPlanYear, recYear.ToString());

            SetComboByText(cmbFromDepartment, r.FromDepartment);
            SetComboByText(cmbToDepartment, r.ToDepartment);

            chkIsCompleted.IsChecked = r.IsCompleted;
            if (r.IsCompleted)
            {
                txtDecisionNumber.Text = r.DecisionNumber ?? "";
                dpDecisionDate.SelectedDate = r.DecisionDate;
                dpEffectiveDate.SelectedDate = r.EffectiveDate;
            }
            ChkIsCompleted_Changed(chkIsCompleted, new RoutedEventArgs());

            txtNote.Text = r.Note ?? "";
        }

        private void UpdatePersonnelDisplay()
        {
            if (_selectedPersonnel == null)
            {
                if (pnlPersonnelEmpty != null) pnlPersonnelEmpty.Visibility = Visibility.Visible;
                if (pnlPersonnelSelected != null) pnlPersonnelSelected.Visibility = Visibility.Collapsed;
                if (imgPersonnelAvatar != null) imgPersonnelAvatar.Visibility = Visibility.Collapsed;
                return;
            }

            if (pnlPersonnelEmpty != null) pnlPersonnelEmpty.Visibility = Visibility.Collapsed;
            if (pnlPersonnelSelected != null) pnlPersonnelSelected.Visibility = Visibility.Visible;

            if (txtPersonnelFullName != null) txtPersonnelFullName.Text = _selectedPersonnel.FullName;
            if (txtPersonnelAvatar != null) txtPersonnelAvatar.Text = GetInitials(_selectedPersonnel.FullName);

            // Hiển thị ảnh đại diện thật nếu có
            if (!string.IsNullOrWhiteSpace(_selectedPersonnel.AvatarBase64))
            {
                try
                {
                    byte[] binaryData = Convert.FromBase64String(_selectedPersonnel.AvatarBase64);
                    var bitmap = TaxPersonnelManagement.Helpers.ImageHelper.LoadAndOrientImage(binaryData);
                    if (bitmap != null && imgPersonnelAvatar != null)
                    {
                        imgPersonnelAvatar.Background = new ImageBrush(bitmap) { Stretch = Stretch.UniformToFill };
                        imgPersonnelAvatar.Visibility = Visibility.Visible;
                    }
                    else if (imgPersonnelAvatar != null)
                    {
                        imgPersonnelAvatar.Visibility = Visibility.Collapsed;
                    }
                }
                catch
                {
                    if (imgPersonnelAvatar != null) imgPersonnelAvatar.Visibility = Visibility.Collapsed;
                }
            }
            else
            {
                if (imgPersonnelAvatar != null) imgPersonnelAvatar.Visibility = Visibility.Collapsed;
            }

            if (txtPersonnelCCCD != null)
                txtPersonnelCCCD.Text = string.IsNullOrEmpty(_selectedPersonnel.IdentityCardNumber) ? "CCCD: ---" : $"CCCD: {_selectedPersonnel.IdentityCardNumber}";
            if (txtPersonnelDOB != null)
                txtPersonnelDOB.Text = _selectedPersonnel.DateOfBirth?.ToString("dd/MM/yyyy") ?? "---";
            if (txtPersonnelDept != null)
                txtPersonnelDept.Text = string.IsNullOrEmpty(_selectedPersonnel.Department) ? "Chưa phân bộ phận" : _selectedPersonnel.Department;

            // (Đã bỏ khối "Đồng bộ dữ liệu tương thích" vì làm hiện lại dòng CCCD | Ngày sinh | Bộ phận trùng lặp với thẻ cán bộ ở trên)

            // Tự động điền bộ phận đang công tác vào cmbFromDepartment nếu đang thêm mới và có thông tin
            if (_existingRecord == null && _selectedPersonnel != null && !string.IsNullOrEmpty(_selectedPersonnel.Department))
            {
                SetComboByText(cmbFromDepartment, _selectedPersonnel.Department);
            }
        }

        private static string GetInitials(string fullName)
        {
            if (string.IsNullOrWhiteSpace(fullName)) return "NV";
            // Bỏ phần ghi chú trong ngoặc, ví dụ "Phan Xuân Triết (Test)" -> "Phan Xuân Triết"
            string clean = System.Text.RegularExpressions.Regex.Replace(fullName, @"\(.*?\)", "").Trim();
            var parts = clean.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0) return "NV";
            if (parts.Length == 1) return parts[0].Substring(0, Math.Min(2, parts[0].Length)).ToUpper();
            return $"{parts[0][0]}{parts[^1][0]}".ToUpper();
        }

        // ============================================================
        // Sự kiện UI
        // ============================================================
        private void Header_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                try { DragMove(); } catch { }
            }
        }

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
                if (!string.IsNullOrEmpty(_selectedPersonnel?.Department))
                {
                    SetComboByText(cmbFromDepartment, _selectedPersonnel.Department);
                }
            }
        }

        private void ChkIsCompleted_Changed(object sender, RoutedEventArgs e)
        {
            bool isDone = chkIsCompleted.IsChecked == true;
            if (pnlDecision != null)
            {
                pnlDecision.Visibility = isDone ? Visibility.Visible : Visibility.Collapsed;
            }

            if (pnlCompletionCard != null)
            {
                if (isDone)
                {
                    pnlCompletionCard.Background = new SolidColorBrush(Color.FromRgb(240, 253, 244)); // #F0FDF4
                    pnlCompletionCard.BorderBrush = new SolidColorBrush(Color.FromRgb(187, 247, 208)); // #BBF7D0
                }
                else
                {
                    pnlCompletionCard.Background = new SolidColorBrush(Color.FromRgb(248, 250, 252)); // #F8FAFC
                    pnlCompletionCard.BorderBrush = new SolidColorBrush(Color.FromRgb(226, 232, 240)); // #E2E8F0
                }
            }
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
            int planYear = DateTime.Now.Year;
            if (int.TryParse(GetComboText(cmbPlanYear), out int py)) planYear = py;

            r.PersonnelId       = _selectedPersonnel!.Id;
            r.RotationType      = GetComboText(cmbRotationType);
            r.PlanType          = GetComboText(cmbPlanType);
            r.FromDepartment    = GetComboText(cmbFromDepartment).NullIfEmpty();
            r.ToDepartment      = GetComboText(cmbToDepartment).NullIfEmpty();
            r.IsCompleted       = chkIsCompleted.IsChecked == true;
            r.DecisionNumber    = r.IsCompleted ? txtDecisionNumber.Text.Trim().NullIfEmpty() : null;
            r.DecisionDate      = r.IsCompleted ? dpDecisionDate.SelectedDate : null;
            r.EffectiveDate     = r.IsCompleted 
                ? (dpEffectiveDate.SelectedDate ?? dpDecisionDate.SelectedDate ?? new DateTime(planYear, 1, 1))
                : (dpEffectiveDate.SelectedDate ?? new DateTime(planYear, 1, 1));
            r.Note              = txtNote.Text.Trim().NullIfEmpty();
        }

        private string GetComboText(ComboBox combo)
        {
            string? text = null;
            if (combo.SelectedItem is ComboBoxItem item)
                text = item.Content?.ToString();
            else if (combo.SelectedItem is string str)
                text = str;
            else
                text = combo.Text;

            if (string.IsNullOrWhiteSpace(text) || text.StartsWith("--"))
                return "";
            return text.Trim();
        }

        private void SetComboByText(ComboBox combo, string? text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                if (combo.Items.Count > 0 && combo.Items[0]?.ToString()?.StartsWith("--") == true)
                    combo.SelectedIndex = 0;
                else
                    combo.SelectedIndex = -1;
                return;
            }

            foreach (var obj in combo.Items)
            {
                if (obj is ComboBoxItem item && string.Equals(item.Content?.ToString(), text, StringComparison.OrdinalIgnoreCase))
                {
                    combo.SelectedItem = item;
                    return;
                }
                else if (obj is string str && string.Equals(str, text, StringComparison.OrdinalIgnoreCase))
                {
                    combo.SelectedItem = obj;
                    return;
                }
            }

            if (combo == cmbPlanYear)
            {
                var newItem = new ComboBoxItem { Content = text };
                combo.Items.Add(newItem);
                combo.SelectedItem = newItem;
            }
            else if (combo == cmbFromDepartment || combo == cmbToDepartment)
            {
                // Thêm vào danh sách để giữ đúng dữ liệu đã lưu
                combo.Items.Add(text);
                combo.SelectedItem = text;
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
