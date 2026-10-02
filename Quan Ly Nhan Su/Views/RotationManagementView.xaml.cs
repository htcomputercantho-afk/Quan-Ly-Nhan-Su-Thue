using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using ClosedXML.Excel;
using Microsoft.EntityFrameworkCore;
using TaxPersonnelManagement.Data;
using TaxPersonnelManagement.Models;

namespace TaxPersonnelManagement.Views
{
    /// <summary>
    /// ViewModel trung gian để bổ sung các thuộc tính tính toán hiển thị trên DataGrid.
    /// </summary>
    public class RotationRowViewModel
    {
        private RotationRecord _record;

        public RotationRowViewModel(RotationRecord record)
        {
            _record = record;
        }

        // --- Chuyển tiếp các thuộc tính từ RotationRecord ---
        public int Id => _record.Id;
        public int STT { get => _record.STT; set => _record.STT = value; }
        public int PersonnelId => _record.PersonnelId;
        public Personnel? Personnel => _record.Personnel;
        public string? FromDepartment => _record.FromDepartment;
        public string? ToDepartment => _record.ToDepartment;
        public string RotationType => _record.RotationType;
        public string PlanType => _record.PlanType;
        public bool IsCompleted => _record.IsCompleted;
        public string? DecisionNumber => _record.DecisionNumber;
        public DateTime? DecisionDate => _record.DecisionDate;
        public DateTime? EffectiveDate => _record.EffectiveDate;
        public string? Note => _record.Note;

        // Initials cho Avatar
        public string Initials
        {
            get
            {
                var name = Personnel?.FullName?.Trim();
                if (string.IsNullOrEmpty(name)) return "CB";
                var parts = name.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length == 1) return parts[0].Substring(0, Math.Min(2, parts[0].Length)).ToUpper();
                return (parts[0][0].ToString() + parts[^1][0].ToString()).ToUpper();
            }
        }

        // Màu nền Avatar phong phú, hiện đại
        public string AvatarBgColor
        {
            get
            {
                var colors = new[] { "#1D4ED8", "#7C3AED", "#0F766E", "#D97706", "#DC2626", "#0369A1", "#4338CA", "#047857" };
                int hash = Math.Abs((Personnel?.FullName ?? "CB").GetHashCode());
                return colors[hash % colors.Length];
            }
        }

        // Huy hiệu giới tính
        public string GenderBadgeBg => Personnel?.Gender == "Nữ" ? "#FDF2F8" : "#EFF6FF";
        public string GenderBadgeFg => Personnel?.Gender == "Nữ" ? "#BE185D" : "#1D4ED8";
        public string GenderBadgeBorder => Personnel?.Gender == "Nữ" ? "#FCE7F3" : "#BFDBFE";
        public string GenderText => string.IsNullOrWhiteSpace(Personnel?.Gender) ? "-" : Personnel.Gender;

        // Trạng thái đã thực hiện / chưa thực hiện
        public string StatusText => IsCompleted ? "Đã thực hiện" : "Chưa thực hiện";
        public string StatusBg => IsCompleted ? "#ECFDF5" : "#FFFBEB";
        public string StatusFg => IsCompleted ? "#047857" : "#B45309";
        public string StatusBorder => IsCompleted ? "#A7F3D0" : "#FDE68A";
        public string StatusIcon => IsCompleted ? "CheckCircle" : "ClockOutline";

        // Chuỗi hiển thị an toàn
        public string DisplayCCCD => string.IsNullOrWhiteSpace(Personnel?.IdentityCardNumber) ? "-" : Personnel.IdentityCardNumber;
        public string DisplayPhone => string.IsNullOrWhiteSpace(Personnel?.PhoneNumber) ? "-" : Personnel.PhoneNumber;
        public string DisplayEmail => string.IsNullOrWhiteSpace(Personnel?.Email) ? "-" : Personnel.Email;
        public string DisplayToDepartment => string.IsNullOrWhiteSpace(ToDepartment) ? "-" : ToDepartment;
        public string DisplayDecisionNumber => string.IsNullOrWhiteSpace(DecisionNumber) ? "-" : DecisionNumber;
        public string DisplayDecisionDate => DecisionDate.HasValue ? DecisionDate.Value.ToString("dd/MM/yyyy") : "-";
        public string DisplayBirthDate => Personnel?.DateOfBirth.HasValue == true ? Personnel.DateOfBirth.Value.ToString("dd/MM/yyyy") : "-";

        // Tham chiếu về bản ghi gốc để dùng khi mở dialog chỉnh sửa
        public RotationRecord Source => _record;
    }

    public partial class RotationManagementView : Page
    {
        // Danh sách gốc (chưa lọc)
        private List<RotationRowViewModel> _allInPlan = new();
        private List<RotationRowViewModel> _allOutPlan = new();

        // Danh sách bộ phận để hiển thị trên combobox lọc
        private List<string> _departments = new();

        private bool _isFilterChanging = true;
        private bool _isLoaded = false;

        public RotationManagementView()
        {
            _isFilterChanging = true;
            InitializeComponent();
            _isFilterChanging = false;
            _isLoaded = true;
            Loaded += (s, e) => LoadData();
        }

        // ============================================================
        // Tải dữ liệu từ cơ sở dữ liệu
        // ============================================================
        public void LoadData()
        {
            try
            {
                using var db = new AppDbContext();

                // Tải tất cả bản ghi luân chuyển điều động, kèm thông tin cán bộ
                var records = db.RotationRecords
                    .Include(r => r.Personnel)
                    .OrderByDescending(r => r.DecisionDate)
                    .ToList();

                records = records.OrderByDescending(r => r.DecisionDate)
                                 .ThenBy(r => r.Personnel?.FullName ?? "")
                                 .ToList();

                // Phân chia theo kế hoạch
                var inPlan = records.Where(r => r.PlanType == "Trong kế hoạch").ToList();
                var outPlan = records.Where(r => r.PlanType == "Ngoài kế hoạch").ToList();

                // Gán STT
                for (int i = 0; i < inPlan.Count; i++) inPlan[i].STT = i + 1;
                for (int i = 0; i < outPlan.Count; i++) outPlan[i].STT = i + 1;

                _allInPlan = inPlan.Select(r => new RotationRowViewModel(r)).ToList();
                _allOutPlan = outPlan.Select(r => new RotationRowViewModel(r)).ToList();

                // Tải danh sách bộ phận để lọc
                _departments = db.Departments
                    .OrderBy(d => d.Name)
                    .Select(d => d.Name)
                    .ToList();

                LoadDepartmentFilters();
                LoadYearFilters();
                ApplyFilters();
                UpdateTotalCount(records.Count);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi tải dữ liệu: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadDepartmentFilters()
        {
            if (cmbFilterDeptInPlan == null || cmbFilterDeptOutPlan == null) return;
            _isFilterChanging = true;

            // Cập nhật combobox lọc bộ phận cho tab "Trong kế hoạch"
            cmbFilterDeptInPlan.Items.Clear();
            cmbFilterDeptInPlan.Items.Add(new ComboBoxItem { Content = "Tất cả bộ phận", IsSelected = true });
            foreach (var d in _departments)
                cmbFilterDeptInPlan.Items.Add(new ComboBoxItem { Content = d });

            // Cập nhật combobox lọc bộ phận cho tab "Ngoài kế hoạch"
            cmbFilterDeptOutPlan.Items.Clear();
            cmbFilterDeptOutPlan.Items.Add(new ComboBoxItem { Content = "Tất cả bộ phận", IsSelected = true });
            foreach (var d in _departments)
                cmbFilterDeptOutPlan.Items.Add(new ComboBoxItem { Content = d });

            cmbFilterDeptInPlan.SelectedIndex = 0;
            cmbFilterDeptOutPlan.SelectedIndex = 0;

            _isFilterChanging = false;
        }

        private void LoadYearFilters()
        {
            if (cmbFilterYearInPlan == null || cmbFilterYearOutPlan == null) return;
            _isFilterChanging = true;

            var years = new HashSet<int> { 2026, 2025, 2024 };
            foreach (var r in _allInPlan.Concat(_allOutPlan))
            {
                if (r.DecisionDate.HasValue) years.Add(r.DecisionDate.Value.Year);
                if (r.EffectiveDate.HasValue) years.Add(r.EffectiveDate.Value.Year);
            }
            var sortedYears = years.OrderByDescending(y => y).ToList();

            cmbFilterYearInPlan.Items.Clear();
            cmbFilterYearInPlan.Items.Add(new ComboBoxItem { Content = "Tất cả" });
            foreach (var y in sortedYears)
                cmbFilterYearInPlan.Items.Add(new ComboBoxItem { Content = y.ToString() });

            cmbFilterYearOutPlan.Items.Clear();
            cmbFilterYearOutPlan.Items.Add(new ComboBoxItem { Content = "Tất cả" });
            foreach (var y in sortedYears)
                cmbFilterYearOutPlan.Items.Add(new ComboBoxItem { Content = y.ToString() });

            SelectYearInCombo(cmbFilterYearInPlan, "2026");
            SelectYearInCombo(cmbFilterYearOutPlan, "2026");

            _isFilterChanging = false;
        }

        private void SelectYearInCombo(ComboBox? combo, string year)
        {
            if (combo == null) return;
            for (int i = 0; i < combo.Items.Count; i++)
            {
                if (combo.Items[i] is ComboBoxItem item && item.Content?.ToString() == year)
                {
                    combo.SelectedIndex = i;
                    return;
                }
            }
            combo.SelectedIndex = 0;
        }

        // ============================================================
        // Áp dụng bộ lọc & tìm kiếm
        // ============================================================
        private void ApplyFilters()
        {
            if (!_isLoaded || dgInPlan == null || dgOutPlan == null) return;

            // --- Tab Trong kế hoạch ---
            var inPlanFiltered = (_allInPlan ?? new List<RotationRowViewModel>()).AsEnumerable();

            string deptInPlan = GetSelectedComboText(cmbFilterDeptInPlan);
            if (deptInPlan != "Tất cả bộ phận" && deptInPlan != "Tất cả" && !string.IsNullOrEmpty(deptInPlan))
                inPlanFiltered = inPlanFiltered.Where(r => r.FromDepartment == deptInPlan || r.ToDepartment == deptInPlan);

            string yearInPlan = GetSelectedComboText(cmbFilterYearInPlan);
            if (yearInPlan != "Tất cả" && int.TryParse(yearInPlan, out int yIn))
                inPlanFiltered = inPlanFiltered.Where(r =>
                    (r.DecisionDate.HasValue && r.DecisionDate.Value.Year == yIn) ||
                    (r.EffectiveDate.HasValue && r.EffectiveDate.Value.Year == yIn));

            string statusInPlan = GetSelectedComboText(cmbFilterStatusInPlan);
            if (statusInPlan == "Đã thực hiện")
                inPlanFiltered = inPlanFiltered.Where(r => r.IsCompleted);
            else if (statusInPlan == "Chưa thực hiện")
                inPlanFiltered = inPlanFiltered.Where(r => !r.IsCompleted);

            if (chkCompletedInPlan?.IsChecked == true)
                inPlanFiltered = inPlanFiltered.Where(r => r.IsCompleted);

            string searchInPlan = txtSearchInPlan?.Text?.Trim().ToLower() ?? "";
            if (!string.IsNullOrEmpty(searchInPlan))
                inPlanFiltered = inPlanFiltered.Where(r =>
                    (r.Personnel?.FullName?.ToLower().Contains(searchInPlan) ?? false) ||
                    (r.Personnel?.IdentityCardNumber?.ToLower().Contains(searchInPlan) ?? false) ||
                    (r.FromDepartment?.ToLower().Contains(searchInPlan) ?? false) ||
                    (r.ToDepartment?.ToLower().Contains(searchInPlan) ?? false) ||
                    (r.DecisionNumber?.ToLower().Contains(searchInPlan) ?? false) ||
                    (r.Personnel?.PhoneNumber?.ToLower().Contains(searchInPlan) ?? false));

            var inList = inPlanFiltered.ToList();
            for (int i = 0; i < inList.Count; i++) inList[i].STT = i + 1;
            dgInPlan.ItemsSource = inList;

            // --- Tab Ngoài kế hoạch ---
            var outPlanFiltered = (_allOutPlan ?? new List<RotationRowViewModel>()).AsEnumerable();

            string deptOutPlan = GetSelectedComboText(cmbFilterDeptOutPlan);
            if (deptOutPlan != "Tất cả bộ phận" && deptOutPlan != "Tất cả" && !string.IsNullOrEmpty(deptOutPlan))
                outPlanFiltered = outPlanFiltered.Where(r => r.FromDepartment == deptOutPlan || r.ToDepartment == deptOutPlan);

            string yearOutPlan = GetSelectedComboText(cmbFilterYearOutPlan);
            if (yearOutPlan != "Tất cả" && int.TryParse(yearOutPlan, out int yOut))
                outPlanFiltered = outPlanFiltered.Where(r =>
                    (r.DecisionDate.HasValue && r.DecisionDate.Value.Year == yOut) ||
                    (r.EffectiveDate.HasValue && r.EffectiveDate.Value.Year == yOut));

            string statusOutPlan = GetSelectedComboText(cmbFilterStatusOutPlan);
            if (statusOutPlan == "Đã thực hiện")
                outPlanFiltered = outPlanFiltered.Where(r => r.IsCompleted);
            else if (statusOutPlan == "Chưa thực hiện")
                outPlanFiltered = outPlanFiltered.Where(r => !r.IsCompleted);

            if (chkCompletedOutPlan?.IsChecked == true)
                outPlanFiltered = outPlanFiltered.Where(r => r.IsCompleted);

            string searchOutPlan = txtSearchOutPlan?.Text?.Trim().ToLower() ?? "";
            if (!string.IsNullOrEmpty(searchOutPlan))
                outPlanFiltered = outPlanFiltered.Where(r =>
                    (r.Personnel?.FullName?.ToLower().Contains(searchOutPlan) ?? false) ||
                    (r.Personnel?.IdentityCardNumber?.ToLower().Contains(searchOutPlan) ?? false) ||
                    (r.FromDepartment?.ToLower().Contains(searchOutPlan) ?? false) ||
                    (r.ToDepartment?.ToLower().Contains(searchOutPlan) ?? false) ||
                    (r.DecisionNumber?.ToLower().Contains(searchOutPlan) ?? false) ||
                    (r.Personnel?.PhoneNumber?.ToLower().Contains(searchOutPlan) ?? false));

            var outList = outPlanFiltered.ToList();
            for (int i = 0; i < outList.Count; i++) outList[i].STT = i + 1;
            dgOutPlan.ItemsSource = outList;
        }

        private string GetSelectedComboText(ComboBox? combo)
        {
            if (combo == null) return "";
            if (combo.SelectedItem is ComboBoxItem item)
                return item.Content?.ToString() ?? "";
            return combo.Text ?? "";
        }

        private void UpdateTotalCount(int total)
        {
            if (txtTotalCount != null)
                txtTotalCount.Text = $"Theo dõi và quản lý tổng cộng {total} lượt luân chuyển & điều động công chức";

            int inPlanCount = _allInPlan.Count;
            int outPlanCount = _allOutPlan.Count;
            int completedCount = _allInPlan.Count(r => r.IsCompleted) + _allOutPlan.Count(r => r.IsCompleted);
            int pendingCount = _allInPlan.Count(r => !r.IsCompleted) + _allOutPlan.Count(r => !r.IsCompleted);

            double inPlanPct = total > 0 ? (double)inPlanCount / total * 100 : 0;
            double outPlanPct = total > 0 ? (double)outPlanCount / total * 100 : 0;
            double completedPct = total > 0 ? (double)completedCount / total * 100 : 0;
            double pendingPct = total > 0 ? (double)pendingCount / total * 100 : 0;

            if (txtTotalKPICount != null) txtTotalKPICount.Text = total.ToString();
            if (txtInPlanKPICount != null) txtInPlanKPICount.Text = $"{inPlanCount} ({inPlanPct:0.#}%)";
            if (txtOutPlanKPICount != null) txtOutPlanKPICount.Text = $"{outPlanCount} ({outPlanPct:0.#}%)";
            if (txtCompletedKPICount != null) txtCompletedKPICount.Text = $"{completedCount} ({completedPct:0.#}%)";
            if (txtPendingKPICount != null) txtPendingKPICount.Text = $"{pendingCount} ({pendingPct:0.#}%)";

            if (badgeInPlan != null) badgeInPlan.Text = inPlanCount.ToString();
            if (badgeOutPlan != null) badgeOutPlan.Text = outPlanCount.ToString();
        }

        // ============================================================
        // Sự kiện KPI Cards & Lọc
        // ============================================================
        private void CardTotal_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            ResetFilters();
        }

        private void CardInPlan_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (tabMain != null && tabInPlan != null)
                tabMain.SelectedItem = tabInPlan;
        }

        private void CardOutPlan_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (tabMain != null && tabOutPlan != null)
                tabMain.SelectedItem = tabOutPlan;
        }

        private void CardCompleted_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            _isFilterChanging = true;
            if (cmbFilterStatusInPlan != null) cmbFilterStatusInPlan.SelectedIndex = 2; // "Đã thực hiện"
            if (cmbFilterStatusOutPlan != null) cmbFilterStatusOutPlan.SelectedIndex = 2;
            _isFilterChanging = false;
            ApplyFilters();
        }

        private void CardPending_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            _isFilterChanging = true;
            if (cmbFilterStatusInPlan != null) cmbFilterStatusInPlan.SelectedIndex = 1; // "Chưa thực hiện"
            if (cmbFilterStatusOutPlan != null) cmbFilterStatusOutPlan.SelectedIndex = 1;
            _isFilterChanging = false;
            ApplyFilters();
        }

        private void BtnResetFilter_Click(object sender, RoutedEventArgs e)
        {
            ResetFilters();
        }

        private void ResetFilters()
        {
            _isFilterChanging = true;
            if (txtSearchInPlan != null) txtSearchInPlan.Text = "";
            if (txtSearchOutPlan != null) txtSearchOutPlan.Text = "";
            if (cmbFilterDeptInPlan != null) cmbFilterDeptInPlan.SelectedIndex = 0;
            if (cmbFilterDeptOutPlan != null) cmbFilterDeptOutPlan.SelectedIndex = 0;
            if (cmbFilterStatusInPlan != null) cmbFilterStatusInPlan.SelectedIndex = 0;
            if (cmbFilterStatusOutPlan != null) cmbFilterStatusOutPlan.SelectedIndex = 0;
            SelectYearInCombo(cmbFilterYearInPlan, "2026");
            SelectYearInCombo(cmbFilterYearOutPlan, "2026");
            if (chkCompletedInPlan != null) chkCompletedInPlan.IsChecked = false;
            if (chkCompletedOutPlan != null) chkCompletedOutPlan.IsChecked = false;
            _isFilterChanging = false;
            ApplyFilters();
        }

        // ============================================================
        // Sự kiện UI
        // ============================================================
        private void TxtSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!_isLoaded || _isFilterChanging) return;
            ApplyFilters();
        }

        private void FilterChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_isLoaded || _isFilterChanging) return;
            ApplyFilters();
        }

        private void ChkCompleted_Changed(object sender, RoutedEventArgs e)
        {
            if (!_isLoaded || _isFilterChanging) return;
            ApplyFilters();
        }

        private void BtnRefresh_Click(object sender, RoutedEventArgs e)
        {
            LoadData();
        }

        /// <summary>Mở dialog thêm mới bản ghi luân chuyển/điều động</summary>
        private void BtnAddRotation_Click(object sender, RoutedEventArgs e)
        {
            // Xác định loại kế hoạch theo tab đang chọn hoặc nút click
            string defaultPlanType = (sender == btnAddRotationOutPlan || tabMain?.SelectedItem == tabOutPlan)
                ? "Ngoài kế hoạch"
                : "Trong kế hoạch";
            var dialog = new RotationRecordDialog(null, defaultPlanType);
            dialog.Owner = Window.GetWindow(this);
            if (dialog.ShowDialog() == true)
            {
                LoadData();
            }
        }

        /// <summary>Double click vào dòng để mở dialog chỉnh sửa</summary>
        private void DgRotation_DoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            var dg = sender as DataGrid;
            if (dg?.SelectedItem is RotationRowViewModel row)
            {
                var dialog = new RotationRecordDialog(row.Source, row.PlanType);
                dialog.Owner = Window.GetWindow(this);
                if (dialog.ShowDialog() == true)
                {
                    LoadData();
                }
            }
        }

        /// <summary>Nút Sửa trên từng dòng</summary>
        private void BtnEditRow_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.DataContext is RotationRowViewModel row)
            {
                var dialog = new RotationRecordDialog(row.Source, row.PlanType);
                dialog.Owner = Window.GetWindow(this);
                if (dialog.ShowDialog() == true)
                {
                    LoadData();
                }
            }
        }

        /// <summary>Nút Xóa trên từng dòng</summary>
        private void BtnDeleteRow_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.DataContext is RotationRowViewModel row)
            {
                var confirm = new ConfirmWindow(
                    $"Bạn có chắc chắn muốn XÓA bản ghi luân chuyển/điều động của cán bộ:\n\n{row.Personnel?.FullName ?? "này"}?",
                    "Xác nhận xóa");
                confirm.Owner = Window.GetWindow(this);
                if (confirm.ShowDialog() == true)
                {
                    try
                    {
                        using var db = new AppDbContext();
                        var record = db.RotationRecords.Find(row.Id);
                        if (record != null)
                        {
                            db.RotationRecords.Remove(record);
                            db.SaveChanges();
                        }
                        LoadData();
                    }
                    catch (Exception ex)
                    {
                        var warning = new WarningWindow($"Lỗi khi xóa:\n{ex.Message}", "Lỗi");
                        if (Window.GetWindow(this) is Window parentWin) warning.Owner = parentWin;
                        warning.ShowDialog();
                    }
                }
            }
        }

        /// <summary>Xuất báo cáo Excel tổng hợp và chi tiết theo kế hoạch / ngoài kế hoạch</summary>
        private async void BtnExportExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Lấy toàn bộ danh sách Trong kế hoạch và Ngoài kế hoạch (tôn trọng bộ lọc nếu đang áp dụng)
                var inPlanItems = (dgInPlan.ItemsSource as IEnumerable<RotationRowViewModel>)?.ToList() 
                                  ?? _allInPlan.ToList();
                var outPlanItems = (dgOutPlan.ItemsSource as IEnumerable<RotationRowViewModel>)?.ToList() 
                                   ?? _allOutPlan.ToList();

                if (inPlanItems.Count == 0 && outPlanItems.Count == 0)
                {
                    var warning = new WarningWindow("Không có dữ liệu trong danh sách để xuất file.", "Thông báo");
                    if (Window.GetWindow(this) is Window parent) warning.Owner = parent;
                    warning.ShowDialog();
                    return;
                }

                var saveDialog = new Microsoft.Win32.SaveFileDialog
                {
                    Filter = "Excel Files (*.xlsx)|*.xlsx",
                    FileName = $"BaoCao_LuanChuyen_DieuDong_{DateTime.Now:yyyyMMdd_HHmmss}",
                    DefaultExt = ".xlsx"
                };

                if (saveDialog.ShowDialog() != true) return;

                string filePath = saveDialog.FileName;

                // Hiển thị hiệu ứng loading đợi xuất file
                if (loadingOverlay != null) loadingOverlay.Visibility = Visibility.Visible;

                await Task.Run(() =>
                {
                    ExportToExcelReport(inPlanItems, outPlanItems, filePath);
                });

                // Khoảng nghỉ nhỏ để hiệu ứng chuyển cảnh mượt mà
                await Task.Delay(350);

                // Ẩn hiệu ứng loading
                if (loadingOverlay != null) loadingOverlay.Visibility = Visibility.Collapsed;

                // Hiển thị cửa sổ thông báo xuất file thành công hiện đại với nút Mở tệp / Mở thư mục
                var successWindow = new SuccessWindow("Xuất báo cáo Excel thành công!", null, filePath, true);
                if (Window.GetWindow(this) is Window parentWin) successWindow.Owner = parentWin;
                successWindow.ShowDialog();
            }
            catch (Exception ex)
            {
                if (loadingOverlay != null) loadingOverlay.Visibility = Visibility.Collapsed;
                var warning = new WarningWindow($"Có lỗi xảy ra khi xuất Excel:\n{ex.Message}", "Lỗi xuất file");
                if (Window.GetWindow(this) is Window parentWin) warning.Owner = parentWin;
                warning.ShowDialog();
            }
        }

        private void ExportToExcelReport(List<RotationRowViewModel> inPlanItems, List<RotationRowViewModel> outPlanItems, string filePath)
        {
            using var workbook = new XLWorkbook();

            // SHEET 1: Báo Cáo Tổng Hợp & Thống Kê (Trả lời trực tiếp các câu hỏi)
            CreateSummaryReportSheet(workbook, inPlanItems, outPlanItems);

            // SHEET 2: Danh sách đầy đủ Trong Kế Hoạch
            CreateDetailedSheet(workbook, "Trong Kế Hoạch", "KẾ HOẠCH ĐIỀU ĐỘNG, LUÂN CHUYỂN CÔNG CHỨC", inPlanItems, "#1976D2");

            // SHEET 3: Danh sách đầy đủ Ngoài Kế Hoạch
            CreateDetailedSheet(workbook, "Ngoài Kế Hoạch", "DANH SÁCH ĐIỀU ĐỘNG, LUÂN CHUYỂN NGOÀI KẾ HOẠCH", outPlanItems, "#7C3AED");

            workbook.SaveAs(filePath);
        }

        private void CreateSummaryReportSheet(XLWorkbook workbook, List<RotationRowViewModel> inPlanItems, List<RotationRowViewModel> outPlanItems)
        {
            var ws = workbook.Worksheets.Add("Tổng Hợp Báo Cáo");
            ws.ShowGridLines = true;

            int totalInPlan = inPlanItems.Count;
            var completedInPlanList = inPlanItems.Where(x => x.IsCompleted).ToList();
            int completedInPlan = completedInPlanList.Count;
            var pendingInPlanList = inPlanItems.Where(x => !x.IsCompleted).ToList();
            int pendingInPlan = pendingInPlanList.Count;

            int totalOutPlan = outPlanItems.Count;
            var completedOutPlanList = outPlanItems.Where(x => x.IsCompleted).ToList();
            int completedOutPlan = completedOutPlanList.Count;

            // 1. Tiêu đề cấp cơ quan & Tiêu đề báo cáo
            ws.Cell("A1").Value = "CỤC THUẾ TỈNH / THÀNH PHỐ";
            ws.Cell("A1").Style.Font.Bold = true;
            ws.Cell("A1").Style.Font.FontSize = 10;
            ws.Cell("A1").Style.Font.FontColor = XLColor.FromHtml("#475569");

            ws.Range("A2:K2").Merge();
            ws.Cell("A2").Value = "BÁO CÁO CÔNG CHỨC LUÂN CHUYỂN, ĐIỀU ĐỘNG";
            ws.Cell("A2").Style.Font.Bold = true;
            ws.Cell("A2").Style.Font.FontSize = 16;
            ws.Cell("A2").Style.Font.FontColor = XLColor.FromHtml("#1E3A8A");
            ws.Cell("A2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Row(2).Height = 28;

            ws.Range("A3:K3").Merge();
            ws.Cell("A3").Value = $"(Thời điểm trích xuất dữ liệu: {DateTime.Now:dd/MM/yyyy HH:mm} - Hệ thống Quản lý Nhân sự)";
            ws.Cell("A3").Style.Font.Italic = true;
            ws.Cell("A3").Style.Font.FontSize = 10;
            ws.Cell("A3").Style.Font.FontColor = XLColor.FromHtml("#64748B");
            ws.Cell("A3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            // 2. KHỐI THỐNG KÊ TỔNG HỢP (DASHBOARD)
            int curRow = 5;
            ws.Range(curRow, 1, curRow, 11).Merge();
            ws.Cell(curRow, 1).Value = "A. TỔNG HỢP SỐ LIỆU LUÂN CHUYỂN, ĐIỀU ĐỘNG";
            ws.Cell(curRow, 1).Style.Font.Bold = true;
            ws.Cell(curRow, 1).Style.Font.FontSize = 12;
            ws.Cell(curRow, 1).Style.Font.FontColor = XLColor.FromHtml("#0F172A");
            ws.Row(curRow).Height = 24;

            curRow++;
            ws.Cell(curRow, 1).Value = "STT";
            ws.Range(curRow, 2, curRow, 5).Merge().Value = "NỘI DUNG THEO DÕI";
            ws.Range(curRow, 6, curRow, 7).Merge().Value = "SỐ LƯỢNG (NGƯỜI)";
            ws.Range(curRow, 8, curRow, 11).Merge().Value = "TIẾN ĐỘ THỰC HIỆN / GHI CHÚ";

            var kpiHeaderRange = ws.Range(curRow, 1, curRow, 11);
            kpiHeaderRange.Style.Font.Bold = true;
            kpiHeaderRange.Style.Font.FontColor = XLColor.White;
            kpiHeaderRange.Style.Fill.BackgroundColor = XLColor.FromHtml("#1E293B");
            kpiHeaderRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            kpiHeaderRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Row(curRow).Height = 26;

            int kpiStartRow = curRow;

            // 1. Tổng số trong kế hoạch
            curRow++;
            ws.Cell(curRow, 1).Value = "1";
            ws.Range(curRow, 2, curRow, 5).Merge().Value = "Tổng số công chức thuộc diện luân chuyển, điều động (Trong kế hoạch)";
            ws.Range(curRow, 6, curRow, 7).Merge().Value = totalInPlan;
            ws.Range(curRow, 8, curRow, 11).Merge().Value = "100% đối tượng theo danh sách kế hoạch đã duyệt";
            ws.Range(curRow, 1, curRow, 11).Style.Font.Bold = true;
            ws.Range(curRow, 1, curRow, 11).Style.Fill.BackgroundColor = XLColor.FromHtml("#EFF6FF");
            ws.Cell(curRow, 6).Style.Font.FontColor = XLColor.FromHtml("#1D4ED8");
            ws.Row(curRow).Height = 22;

            // 1.1 Đã thực hiện
            curRow++;
            ws.Cell(curRow, 1).Value = "1.1";
            ws.Range(curRow, 2, curRow, 5).Merge().Value = "↳ Trong đó: Đã hoàn thành luân chuyển, điều động";
            ws.Range(curRow, 6, curRow, 7).Merge().Value = completedInPlan;
            double pctComp = totalInPlan > 0 ? (double)completedInPlan / totalInPlan * 100 : 0;
            ws.Range(curRow, 8, curRow, 11).Merge().Value = $"Đạt {pctComp:0.#}% (Đã có quyết định điều động chính thức)";
            ws.Cell(curRow, 6).Style.Font.Bold = true;
            ws.Cell(curRow, 6).Style.Font.FontColor = XLColor.FromHtml("#059669");
            ws.Row(curRow).Height = 22;

            // 1.2 Chưa thực hiện
            curRow++;
            ws.Cell(curRow, 1).Value = "1.2";
            ws.Range(curRow, 2, curRow, 5).Merge().Value = "↳ Trong đó: Chưa thực hiện luân chuyển, điều động";
            ws.Range(curRow, 6, curRow, 7).Merge().Value = pendingInPlan;
            double pctPend = totalInPlan > 0 ? (double)pendingInPlan / totalInPlan * 100 : 0;
            ws.Range(curRow, 8, curRow, 11).Merge().Value = $"Còn lại {pctPend:0.#}% (Đang chuẩn bị quy trình / Chờ QĐ)";
            ws.Cell(curRow, 6).Style.Font.Bold = true;
            ws.Cell(curRow, 6).Style.Font.FontColor = XLColor.FromHtml("#D97706");
            ws.Row(curRow).Height = 22;

            // 2. Ngoài kế hoạch đã thực hiện
            curRow++;
            ws.Cell(curRow, 1).Value = "2";
            ws.Range(curRow, 2, curRow, 5).Merge().Value = "Số công chức ngoài kế hoạch đã thực hiện luân chuyển, điều động";
            ws.Range(curRow, 6, curRow, 7).Merge().Value = completedOutPlan;
            ws.Range(curRow, 8, curRow, 11).Merge().Value = $"Phát sinh ngoài kế hoạch (Đã điều động: {completedOutPlan}/{totalOutPlan} người)";
            ws.Range(curRow, 1, curRow, 11).Style.Font.Bold = true;
            ws.Range(curRow, 1, curRow, 11).Style.Fill.BackgroundColor = XLColor.FromHtml("#F5F3FF");
            ws.Cell(curRow, 6).Style.Font.FontColor = XLColor.FromHtml("#7C3AED");
            ws.Row(curRow).Height = 22;

            // Viền và căn lề bảng KPI
            var kpiTable = ws.Range(kpiStartRow, 1, curRow, 11);
            kpiTable.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            kpiTable.Style.Border.SetOutsideBorderColor(XLColor.FromHtml("#64748B"));
            kpiTable.Style.Border.SetInsideBorder(XLBorderStyleValues.Thin);
            kpiTable.Style.Border.SetInsideBorderColor(XLColor.FromHtml("#CBD5E1"));
            for (int r = kpiStartRow + 1; r <= curRow; r++)
            {
                ws.Cell(r, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(r, 6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(r, 8).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                ws.Row(r).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            }

            // ====================================================================
            // 3. CHI TIẾT 3 NHÓM (TRẢ LỜI CỤ THỂ "VÀ LÀ NHỮNG AI")
            // ====================================================================
            curRow += 2;
            ws.Range(curRow, 1, curRow, 11).Merge();
            ws.Cell(curRow, 1).Value = "B. DANH SÁCH CHI TIẾT THEO TỪNG DIỆN LUÂN CHUYỂN, ĐIỀU ĐỘNG";
            ws.Cell(curRow, 1).Style.Font.Bold = true;
            ws.Cell(curRow, 1).Style.Font.FontSize = 12;
            ws.Cell(curRow, 1).Style.Font.FontColor = XLColor.FromHtml("#0F172A");
            ws.Row(curRow).Height = 24;

            // --- BẢNG I: TRONG KẾ HOẠCH ĐÃ THỰC HIỆN ({completedInPlan} NGƯỜI) ---
            curRow++;
            RenderSectionTitle(ws, curRow, $"I. DANH SÁCH CÔNG CHỨC TRONG KẾ HOẠCH ĐÃ HOÀN THÀNH LUÂN CHUYỂN, ĐIỀU ĐỘNG ({completedInPlan} người)", "#047857");
            
            curRow++;
            var headersCompleted = new[] { "STT", "Họ và tên", "Số CCCD", "Ngày sinh", "Giới tính", "Đơn vị chuyển đi", "Bộ phận đến", "Loại hình", "Số Quyết định", "Ngày ra QĐ", "Ghi chú" };
            RenderTableHeader(ws, curRow, headersCompleted, "#065F46");

            int tbl1Start = curRow + 1;
            if (completedInPlanList.Count == 0)
            {
                curRow++;
                ws.Range(curRow, 1, curRow, 11).Merge().Value = "Chưa có cán bộ nào trong kế hoạch hoàn thành điều động, luân chuyển.";
                ws.Cell(curRow, 1).Style.Font.Italic = true;
                ws.Cell(curRow, 1).Style.Font.FontColor = XLColor.FromHtml("#64748B");
                ws.Cell(curRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Row(curRow).Height = 22;
            }
            else
            {
                for (int i = 0; i < completedInPlanList.Count; i++)
                {
                    curRow++;
                    var r = completedInPlanList[i];
                    FillRowData(ws, curRow, i + 1, r, isCompletedTable: true);
                }
            }
            FormatTableRange(ws, tbl1Start - 1, 1, curRow, 11);

            // --- BẢNG II: TRONG KẾ HOẠCH CHƯA THỰC HIỆN ({pendingInPlan} NGƯỜI) ---
            curRow += 2;
            RenderSectionTitle(ws, curRow, $"II. DANH SÁCH CÔNG CHỨC TRONG KẾ HOẠCH CHƯA THỰC HIỆN LUÂN CHUYỂN, ĐIỀU ĐỘNG ({pendingInPlan} người)", "#B45309");

            curRow++;
            var headersPending = new[] { "STT", "Họ và tên", "Số CCCD", "Ngày sinh", "Giới tính", "Đơn vị hiện tại", "Bộ phận dự kiến đến", "Loại hình", "Trạng thái", "Ngày hiệu lực dự kiến", "Ghi chú" };
            RenderTableHeader(ws, curRow, headersPending, "#92400E");

            int tbl2Start = curRow + 1;
            if (pendingInPlanList.Count == 0)
            {
                curRow++;
                ws.Range(curRow, 1, curRow, 11).Merge().Value = "✓ 100% công chức thuộc diện theo kế hoạch đều đã hoàn thành điều động, luân chuyển.";
                ws.Cell(curRow, 1).Style.Font.Italic = true;
                ws.Cell(curRow, 1).Style.Font.FontColor = XLColor.FromHtml("#059669");
                ws.Cell(curRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Row(curRow).Height = 22;
            }
            else
            {
                for (int i = 0; i < pendingInPlanList.Count; i++)
                {
                    curRow++;
                    var r = pendingInPlanList[i];
                    FillRowData(ws, curRow, i + 1, r, isCompletedTable: false);
                }
            }
            FormatTableRange(ws, tbl2Start - 1, 1, curRow, 11);

            // --- BẢNG III: NGOÀI KẾ HOẠCH ĐÃ THỰC HIỆN ({completedOutPlan} NGƯỜI) ---
            curRow += 2;
            RenderSectionTitle(ws, curRow, $"III. DANH SÁCH CÔNG CHỨC NGOÀI KẾ HOẠCH ĐÃ LUÂN CHUYỂN, ĐIỀU ĐỘNG ({completedOutPlan} người)", "#6D28D9");

            curRow++;
            RenderTableHeader(ws, curRow, headersCompleted, "#5B21B6");

            int tbl3Start = curRow + 1;
            if (completedOutPlanList.Count == 0)
            {
                curRow++;
                ws.Range(curRow, 1, curRow, 11).Merge().Value = "Không có công chức nào ngoài kế hoạch đã điều động, luân chuyển.";
                ws.Cell(curRow, 1).Style.Font.Italic = true;
                ws.Cell(curRow, 1).Style.Font.FontColor = XLColor.FromHtml("#64748B");
                ws.Cell(curRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Row(curRow).Height = 22;
            }
            else
            {
                for (int i = 0; i < completedOutPlanList.Count; i++)
                {
                    curRow++;
                    var r = completedOutPlanList[i];
                    FillRowData(ws, curRow, i + 1, r, isCompletedTable: true);
                }
            }
            FormatTableRange(ws, tbl3Start - 1, 1, curRow, 11);

            // Đặt độ rộng các cột tối ưu
            ws.Column(1).Width = 7;     // STT
            ws.Column(2).Width = 24;    // Họ và tên
            ws.Column(3).Width = 16;    // CCCD
            ws.Column(4).Width = 14;    // Ngày sinh
            ws.Column(5).Width = 11;    // Giới tính
            ws.Column(6).Width = 25;    // Đơn vị chuyển đi
            ws.Column(7).Width = 25;    // Bộ phận đến
            ws.Column(8).Width = 15;    // Loại hình
            ws.Column(9).Width = 20;    // Số QĐ / Trạng thái
            ws.Column(10).Width = 15;   // Ngày ra QĐ
            ws.Column(11).Width = 24;   // Ghi chú
        }

        private void CreateDetailedSheet(XLWorkbook workbook, string sheetName, string title, List<RotationRowViewModel> items, string primaryColor)
        {
            var ws = workbook.Worksheets.Add(sheetName);
            ws.ShowGridLines = true;

            // Tiêu đề sheet
            ws.Cell("A1").Value = title;
            ws.Cell("A1").Style.Font.Bold = true;
            ws.Cell("A1").Style.Font.FontSize = 14;
            ws.Cell("A1").Style.Font.FontColor = XLColor.FromHtml(primaryColor);
            ws.Row(1).Height = 26;

            int completedCount = items.Count(x => x.IsCompleted);
            int pendingCount = items.Count - completedCount;
            ws.Cell("A2").Value = $"Tổng số: {items.Count} người  |  Đã hoàn thành: {completedCount} người  |  Chưa thực hiện: {pendingCount} người";
            ws.Cell("A2").Style.Font.Italic = true;
            ws.Cell("A2").Style.Font.FontSize = 10;
            ws.Cell("A2").Style.Font.FontColor = XLColor.FromHtml("#64748B");
            ws.Row(2).Height = 18;

            int headerRow = 4;
            var headers = new[] { "STT", "Họ và tên", "Số CCCD", "Ngày sinh", "Giới tính", "Đơn vị hiện tại / Chuyển đi", "Bộ phận dự kiến / Đến", "Loại hình", "Trạng thái", "Số Quyết định", "Ngày ra QĐ", "Ghi chú" };
            
            for (int col = 1; col <= headers.Length; col++)
            {
                var cell = ws.Cell(headerRow, col);
                cell.Value = headers[col - 1];
                cell.Style.Font.Bold = true;
                cell.Style.Font.FontSize = 10;
                cell.Style.Font.FontColor = XLColor.White;
                cell.Style.Fill.BackgroundColor = XLColor.FromHtml(primaryColor);
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            }
            ws.Row(headerRow).Height = 26;

            for (int i = 0; i < items.Count; i++)
            {
                int row = headerRow + 1 + i;
                var r = items[i];

                ws.Cell(row, 1).Value = i + 1;
                ws.Cell(row, 2).Value = r.Personnel?.FullName ?? "";
                ws.Cell(row, 2).Style.Font.Bold = true;
                ws.Cell(row, 3).Value = r.Personnel?.IdentityCardNumber ?? "";
                ws.Cell(row, 4).Value = r.Personnel?.DateOfBirth?.ToString("dd/MM/yyyy") ?? "";
                ws.Cell(row, 5).Value = r.Personnel?.Gender ?? "";
                ws.Cell(row, 6).Value = !string.IsNullOrEmpty(r.FromDepartment) ? r.FromDepartment : (r.Personnel?.Department ?? "");
                ws.Cell(row, 7).Value = r.ToDepartment ?? "";
                ws.Cell(row, 8).Value = r.RotationType;

                var statusCell = ws.Cell(row, 9);
                statusCell.Value = r.IsCompleted ? "Đã thực hiện" : "Chưa thực hiện";
                statusCell.Style.Font.Bold = true;
                statusCell.Style.Font.FontColor = r.IsCompleted ? XLColor.FromHtml("#059669") : XLColor.FromHtml("#D97706");

                ws.Cell(row, 10).Value = r.DecisionNumber ?? "";
                ws.Cell(row, 11).Value = r.DecisionDate?.ToString("dd/MM/yyyy") ?? "";
                ws.Cell(row, 12).Value = r.Note ?? "";

                // Căn chỉnh
                ws.Cell(row, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(row, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(row, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(row, 5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(row, 8).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(row, 9).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(row, 10).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(row, 11).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                ws.Row(row).Height = 22;
                ws.Row(row).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                if (i % 2 == 1)
                    ws.Row(row).Style.Fill.BackgroundColor = XLColor.FromHtml("#F8FAFC");
            }

            int endRow = headerRow + items.Count;
            if (items.Count > 0)
            {
                var tableRange = ws.Range(headerRow, 1, endRow, headers.Length);
                tableRange.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                tableRange.Style.Border.SetOutsideBorderColor(XLColor.FromHtml("#94A3B8"));
                tableRange.Style.Border.SetInsideBorder(XLBorderStyleValues.Thin);
                tableRange.Style.Border.SetInsideBorderColor(XLColor.FromHtml("#CBD5E1"));
                tableRange.SetAutoFilter();
            }

            ws.Column(1).Width = 7;
            ws.Column(2).Width = 24;
            ws.Column(3).Width = 16;
            ws.Column(4).Width = 14;
            ws.Column(5).Width = 11;
            ws.Column(6).Width = 25;
            ws.Column(7).Width = 25;
            ws.Column(8).Width = 14;
            ws.Column(9).Width = 16;
            ws.Column(10).Width = 18;
            ws.Column(11).Width = 14;
            ws.Column(12).Width = 24;
        }

        private void RenderSectionTitle(IXLWorksheet ws, int row, string title, string colorHex)
        {
            ws.Range(row, 1, row, 11).Merge().Value = title;
            var range = ws.Range(row, 1, row, 11);
            range.Style.Font.Bold = true;
            range.Style.Font.FontSize = 11;
            range.Style.Font.FontColor = XLColor.White;
            range.Style.Fill.BackgroundColor = XLColor.FromHtml(colorHex);
            range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            range.Style.Alignment.Indent = 1;
            ws.Row(row).Height = 25;
        }

        private void RenderTableHeader(IXLWorksheet ws, int row, string[] headers, string colorHex)
        {
            for (int col = 1; col <= headers.Length; col++)
            {
                var cell = ws.Cell(row, col);
                cell.Value = headers[col - 1];
                cell.Style.Font.Bold = true;
                cell.Style.Font.FontSize = 9.5;
                cell.Style.Font.FontColor = XLColor.White;
                cell.Style.Fill.BackgroundColor = XLColor.FromHtml(colorHex);
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            }
            ws.Row(row).Height = 24;
        }

        private void FillRowData(IXLWorksheet ws, int row, int stt, RotationRowViewModel r, bool isCompletedTable)
        {
            ws.Cell(row, 1).Value = stt;
            ws.Cell(row, 2).Value = r.Personnel?.FullName ?? "";
            ws.Cell(row, 2).Style.Font.Bold = true;
            ws.Cell(row, 3).Value = r.Personnel?.IdentityCardNumber ?? "";
            ws.Cell(row, 4).Value = r.Personnel?.DateOfBirth?.ToString("dd/MM/yyyy") ?? "";
            ws.Cell(row, 5).Value = r.Personnel?.Gender ?? "";
            ws.Cell(row, 6).Value = !string.IsNullOrEmpty(r.FromDepartment) ? r.FromDepartment : (r.Personnel?.Department ?? "");
            ws.Cell(row, 7).Value = r.ToDepartment ?? "";
            ws.Cell(row, 8).Value = r.RotationType;

            if (isCompletedTable)
            {
                ws.Cell(row, 9).Value = r.DecisionNumber ?? "";
                ws.Cell(row, 10).Value = r.DecisionDate?.ToString("dd/MM/yyyy") ?? "";
                ws.Cell(row, 11).Value = r.Note ?? "";
            }
            else
            {
                var statusCell = ws.Cell(row, 9);
                statusCell.Value = "Chưa thực hiện";
                statusCell.Style.Font.Bold = true;
                statusCell.Style.Font.FontColor = XLColor.FromHtml("#D97706");

                ws.Cell(row, 10).Value = r.EffectiveDate?.ToString("dd/MM/yyyy") ?? "";
                ws.Cell(row, 11).Value = r.Note ?? "";
            }

            // Căn chỉnh
            ws.Cell(row, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell(row, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell(row, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell(row, 5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell(row, 8).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell(row, 9).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell(row, 10).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            ws.Row(row).Height = 21;
            ws.Row(row).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            if (stt % 2 == 0)
                ws.Row(row).Style.Fill.BackgroundColor = XLColor.FromHtml("#F8FAFC");
        }

        private void FormatTableRange(IXLWorksheet ws, int startRow, int startCol, int endRow, int endCol)
        {
            var range = ws.Range(startRow, startCol, endRow, endCol);
            range.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            range.Style.Border.SetOutsideBorderColor(XLColor.FromHtml("#94A3B8"));
            range.Style.Border.SetInsideBorder(XLBorderStyleValues.Thin);
            range.Style.Border.SetInsideBorderColor(XLColor.FromHtml("#CBD5E1"));
        }
    }
}
