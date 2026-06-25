using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using Microsoft.EntityFrameworkCore;
using Microsoft.Win32;
using TaxPersonnelManagement.Data;
using TaxPersonnelManagement.Models;
using TaxPersonnelManagement.Services;

namespace TaxPersonnelManagement.Views
{
    public partial class PlanningManagementView : Page
    {
        private List<PlanningRecord> _allProfRecords = new();
        private List<PlanningRecord> _allPartyRecords = new();
        private List<string> _departments = new();
        private bool _isLoaded = false;

        public PlanningManagementView()
        {
            InitializeComponent();
            LoadFilterData();
            LoadData();
            _isLoaded = true;
        }

        /// <summary>
        /// Tải danh mục phòng ban để đưa vào bộ lọc ComboBox
        /// </summary>
        private void LoadFilterData()
        {
            try
            {
                using (var db = new AppDbContext())
                {
                    _departments = db.Departments
                                     .OrderBy(d => d.Name)
                                     .Select(d => d.Name)
                                     .ToList();

                    var profList = new List<string> { "Tất cả bộ phận" };
                    profList.AddRange(_departments);

                    var partyList = new List<string> { "Tất cả bộ phận" };
                    partyList.AddRange(_departments);

                    cboDeptProf.ItemsSource = profList;
                    cboDeptProf.SelectedIndex = 0;

                    cboDeptParty.ItemsSource = partyList;
                    cboDeptParty.SelectedIndex = 0;

                    // Tải danh mục nhiệm kỳ quy hoạch động
                    var terms = db.PlanningTerms
                                  .OrderByDescending(t => t.TermName)
                                  .Select(t => t.TermName)
                                  .ToList();

                    var termListProf = new List<string> { "Tất cả nhiệm kỳ" };
                    termListProf.AddRange(terms);

                    var termListParty = new List<string> { "Tất cả nhiệm kỳ" };
                    termListParty.AddRange(terms);

                    cboTermProf.ItemsSource = termListProf;
                    cboTermProf.SelectedIndex = 0;

                    cboTermParty.ItemsSource = termListParty;
                    cboTermParty.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                var warning = new WarningWindow($"Lỗi tải danh mục bộ phận và nhiệm kỳ: {ex.Message}", "Lỗi");
                warning.Owner = Window.GetWindow(this);
                warning.ShowDialog();
            }
        }

        /// <summary>
        /// Tải toàn bộ dữ liệu quy hoạch từ database
        /// </summary>
        public void LoadData()
        {
            try
            {
                using (var db = new AppDbContext())
                {
                    // Cập nhật lại danh mục bộ lọc nhiệm kỳ quy hoạch động (Mới nhất lên đầu)
                    var currentProfSelected = cboTermProf.SelectedItem as string;
                    var currentPartySelected = cboTermParty.SelectedItem as string;

                    var terms = db.PlanningTerms
                                  .OrderByDescending(t => t.TermName)
                                  .Select(t => t.TermName)
                                  .ToList();

                    var termListProf = new List<string> { "Tất cả nhiệm kỳ" };
                    termListProf.AddRange(terms);

                    var termListParty = new List<string> { "Tất cả nhiệm kỳ" };
                    termListParty.AddRange(terms);

                    cboTermProf.ItemsSource = termListProf;
                    if (currentProfSelected != null && termListProf.Contains(currentProfSelected))
                        cboTermProf.SelectedItem = currentProfSelected;
                    else
                        cboTermProf.SelectedIndex = 0;

                    cboTermParty.ItemsSource = termListParty;
                    if (currentPartySelected != null && termListParty.Contains(currentPartySelected))
                        cboTermParty.SelectedItem = currentPartySelected;
                    else
                        cboTermParty.SelectedIndex = 0;

                    // Lấy tất cả bản ghi quy hoạch kèm thông tin Personnel liên quan
                    var records = db.PlanningRecords
                                    .Include(r => r.Personnel)
                                    .ToList();

                    foreach (var r in records)
                    {
                        if (!string.IsNullOrEmpty(r.Evaluation3Years))
                        {
                            r.Evaluation3Years = System.Text.RegularExpressions.Regex.Replace(r.Evaluation3Years, @",\s*(?=\d{4}:)", Environment.NewLine);
                        }
                    }

                    _allProfRecords = records.Where(r => r.PlanningType == "Chuyên môn").ToList();
                    _allPartyRecords = records.Where(r => r.PlanningType == "Đảng").ToList();

                    ApplyFiltersProf();
                    ApplyFiltersParty();
                }
            }
            catch (Exception ex)
            {
                var warning = new WarningWindow($"Lỗi tải dữ liệu quy hoạch: {ex.Message}", "Lỗi");
                warning.Owner = Window.GetWindow(this);
                warning.ShowDialog();
            }
        }

        #region Quy hoạch Chuyên môn (Tab 1)

        private void FilterProf_Changed(object sender, object e)
        {
            if (!_isLoaded) return;
            ApplyFiltersProf();
        }

        private void ApplyFiltersProf()
        {
            string search = txtSearchProf.Text.Trim().ToLower();
            string dept = cboDeptProf.SelectedItem as string ?? "Tất cả bộ phận";
            string term = cboTermProf.SelectedItem as string ?? "Tất cả nhiệm kỳ";
            string status = (cboStatusProf.SelectedItem as ComboBoxItem)?.Content.ToString() ?? "Tất cả trạng thái";

            var filtered = _allProfRecords.AsEnumerable();

            // Lọc theo tìm kiếm (Tên, CCCD)
            if (!string.IsNullOrEmpty(search))
            {
                filtered = filtered.Where(r => (r.Personnel != null && r.Personnel.FullName.ToLower().Contains(search)) ||
                                               (r.Personnel != null && r.Personnel.IdentityCardNumber != null && r.Personnel.IdentityCardNumber.Contains(search)));
            }

            // Lọc theo bộ phận
            if (dept != "Tất cả bộ phận")
            {
                filtered = filtered.Where(r => r.Personnel != null && r.Personnel.Department == dept);
            }

            // Lọc theo nhiệm kỳ
            if (term != "Tất cả nhiệm kỳ")
            {
                filtered = filtered.Where(r => r.PlanningTerm == term);
            }

            // Lọc theo trạng thái
            if (status != "Tất cả trạng thái")
            {
                // status có dạng "1. Tiếp tục quy hoạch" -> Cần cắt bỏ phần số để lấy text gốc
                string cleanStatus = status.Contains(".") ? status.Substring(status.IndexOf(".") + 1).Trim() : status;
                filtered = filtered.Where(r => r.Status == cleanStatus);
            }

            // Sắp xếp theo thứ tự quy định: 1. Tiếp tục quy hoạch, 2. Đưa ra khỏi quy hoạch, 3. Bổ sung quy hoạch
            // Sau đó sắp xếp theo Bộ phận, rồi Họ tên
            var sortedList = filtered
                .OrderBy(r => GetStatusOrder(r.Status))
                .ThenBy(r => r.Personnel?.Department)
                .ThenBy(r => r.Personnel?.FullName)
                .ToList();

            // Đánh số thứ tự STT chạy riêng
            int stt = 1;
            foreach (var r in sortedList)
            {
                r.STT = stt++;
            }

            // Cập nhật lên UI với WPF Grouping
            var view = new ListCollectionView(sortedList);
            view.GroupDescriptions.Add(new PropertyGroupDescription("Status"));

            dgProf.ItemsSource = view;
            txtTotalProf.Text = $"Hiển thị {sortedList.Count} bản ghi quy hoạch chuyên môn";
        }

        private void btnAddProf_Click(object sender, RoutedEventArgs e)
        {
            OpenPlanningDialog("Chuyên môn");
        }

        private void btnExportProf_Click(object sender, RoutedEventArgs e)
        {
            ExportToExcel("Chuyên môn");
        }

        #endregion

        #region Quy hoạch Đảng (Tab 2)

        private void FilterParty_Changed(object sender, object e)
        {
            if (!_isLoaded) return;
            ApplyFiltersParty();
        }

        private void ApplyFiltersParty()
        {
            string search = txtSearchParty.Text.Trim().ToLower();
            string dept = cboDeptParty.SelectedItem as string ?? "Tất cả bộ phận";
            string term = cboTermParty.SelectedItem as string ?? "Tất cả nhiệm kỳ";
            string status = (cboStatusParty.SelectedItem as ComboBoxItem)?.Content.ToString() ?? "Tất cả trạng thái";

            var filtered = _allPartyRecords.AsEnumerable();

            // Lọc theo tìm kiếm (Tên, CCCD)
            if (!string.IsNullOrEmpty(search))
            {
                filtered = filtered.Where(r => (r.Personnel != null && r.Personnel.FullName.ToLower().Contains(search)) ||
                                               (r.Personnel != null && r.Personnel.IdentityCardNumber != null && r.Personnel.IdentityCardNumber.Contains(search)));
            }

            // Lọc theo bộ phận
            if (dept != "Tất cả bộ phận")
            {
                filtered = filtered.Where(r => r.Personnel != null && r.Personnel.Department == dept);
            }

            // Lọc theo nhiệm kỳ
            if (term != "Tất cả nhiệm kỳ")
            {
                filtered = filtered.Where(r => r.PlanningTerm == term);
            }

            // Lọc theo trạng thái
            if (status != "Tất cả trạng thái")
            {
                string cleanStatus = status.Contains(".") ? status.Substring(status.IndexOf(".") + 1).Trim() : status;
                filtered = filtered.Where(r => r.Status == cleanStatus);
            }

            // Sắp xếp
            var sortedList = filtered
                .OrderBy(r => GetStatusOrder(r.Status))
                .ThenBy(r => r.Personnel?.Department)
                .ThenBy(r => r.Personnel?.FullName)
                .ToList();

            // Đánh số thứ tự STT
            int stt = 1;
            foreach (var r in sortedList)
            {
                r.STT = stt++;
            }

            var view = new ListCollectionView(sortedList);
            view.GroupDescriptions.Add(new PropertyGroupDescription("Status"));

            dgParty.ItemsSource = view;
            txtTotalParty.Text = $"Hiển thị {sortedList.Count} bản ghi quy hoạch đảng";
        }

        private void btnAddParty_Click(object sender, RoutedEventArgs e)
        {
            OpenPlanningDialog("Đảng");
        }

        private void btnExportParty_Click(object sender, RoutedEventArgs e)
        {
            ExportToExcel("Đảng");
        }

        #endregion

        #region Trợ giúp & Thao tác Chung

        private int GetStatusOrder(string status)
        {
            return status switch
            {
                "Tiếp tục quy hoạch" => 1,
                "Đưa ra khỏi quy hoạch" => 2,
                "Bổ sung quy hoạch" => 3,
                _ => 4
            };
        }

        private void tcPlanning_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Reset bộ lọc khi đổi Tab để giao diện sạch sẽ
            if (e.Source is TabControl)
            {
                LoadData();
            }
        }

        private void OpenPlanningDialog(string planningType)
        {
            try
            {
                var dialog = new PlanningRecordDialog(planningType);
                dialog.Owner = Window.GetWindow(this);
                if (dialog.ShowDialog() == true && dialog.ResultRecord != null)
                {
                    using (var db = new AppDbContext())
                    {
                        db.PlanningRecords.Add(dialog.ResultRecord);
                        db.SaveChanges();
                    }
                    LoadData();
                    var success = new SuccessWindow("Thêm thông tin quy hoạch thành công!");
                    success.Owner = Window.GetWindow(this);
                    success.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                var warning = new WarningWindow($"Lỗi lưu thông tin quy hoạch: {ex.Message}", "Lỗi");
                warning.Owner = Window.GetWindow(this);
                warning.ShowDialog();
            }
        }

        private void btnEdit_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (sender is Button btn && btn.DataContext is PlanningRecord record)
                {
                    // Lấy đầy đủ thông tin từ database để sửa đổi
                    using (var db = new AppDbContext())
                    {
                        var fullRecord = db.PlanningRecords
                                           .Include(r => r.Personnel)
                                           .FirstOrDefault(r => r.Id == record.Id);

                        if (fullRecord == null)
                        {
                            var warning = new WarningWindow("Bản ghi không tồn tại hoặc đã bị xóa!", "Lỗi");
                            warning.Owner = Window.GetWindow(this);
                            warning.ShowDialog();
                            return;
                        }

                        var dialog = new PlanningRecordDialog(fullRecord);
                        dialog.Owner = Window.GetWindow(this);
                        if (dialog.ShowDialog() == true && dialog.ResultRecord != null)
                        {
                            // Cập nhật thông tin bản ghi
                            var entity = db.PlanningRecords.Find(fullRecord.Id);
                            if (entity != null)
                            {
                                entity.Status = dialog.ResultRecord.Status;
                                entity.PlanningTerm = dialog.ResultRecord.PlanningTerm;
                                entity.CurrentPosition = dialog.ResultRecord.CurrentPosition;
                                entity.PlannedPosition = dialog.ResultRecord.PlannedPosition;
                                entity.PlannedTransitionPosition = dialog.ResultRecord.PlannedTransitionPosition;
                                entity.TrainingLevel = dialog.ResultRecord.TrainingLevel;
                                entity.PoliticalTheoryLevel = dialog.ResultRecord.PoliticalTheoryLevel;
                                entity.DecisionNumber = dialog.ResultRecord.DecisionNumber;
                                entity.DecisionDate = dialog.ResultRecord.DecisionDate;
                                entity.DecisionUnit = dialog.ResultRecord.DecisionUnit;
                                entity.Evaluation3Years = dialog.ResultRecord.Evaluation3Years;
                                entity.Note = dialog.ResultRecord.Note;

                                db.SaveChanges();
                            }
                            LoadData();
                            var success = new SuccessWindow("Cập nhật thông tin quy hoạch thành công!");
                            success.Owner = Window.GetWindow(this);
                            success.ShowDialog();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var warning = new WarningWindow($"Lỗi cập nhật thông tin quy hoạch: {ex.Message}", "Lỗi");
                warning.Owner = Window.GetWindow(this);
                warning.ShowDialog();
            }
        }

        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (sender is Button btn && btn.DataContext is PlanningRecord record)
                {
                    var confirm = new ConfirmWindow($"Bạn có chắc chắn muốn xóa bản ghi quy hoạch của cán bộ '{record.Personnel?.FullName}' không?", "Xác nhận xóa");
                    confirm.Owner = Window.GetWindow(this);

                    if (confirm.ShowDialog() == true)
                    {
                        using (var db = new AppDbContext())
                        {
                            var entity = db.PlanningRecords.Find(record.Id);
                            if (entity != null)
                            {
                                db.PlanningRecords.Remove(entity);
                                db.SaveChanges();
                            }
                        }
                        LoadData();
                        var success = new SuccessWindow("Xóa bản ghi quy hoạch thành công!");
                        success.Owner = Window.GetWindow(this);
                        success.ShowDialog();
                    }
                }
            }
            catch (Exception ex)
            {
                var warning = new WarningWindow($"Lỗi xóa bản ghi quy hoạch: {ex.Message}", "Lỗi");
                warning.Owner = Window.GetWindow(this);
                warning.ShowDialog();
            }
        }

        private async void ExportToExcel(string planningType)
        {
            try
            {
                // Lấy danh sách theo tab đang xuất
                List<PlanningRecord> records;
                if (planningType == "Chuyên môn")
                {
                    records = dgProf.ItemsSource is ICollectionView cvProf
                        ? cvProf.Cast<PlanningRecord>().ToList()
                        : _allProfRecords;
                }
                else
                {
                    records = dgParty.ItemsSource is ICollectionView cvParty
                        ? cvParty.Cast<PlanningRecord>().ToList()
                        : _allPartyRecords;
                }

                if (!records.Any())
                {
                    var warning = new WarningWindow("Không có dữ liệu hiển thị để xuất Excel!", "Thông báo");
                    warning.Owner = Window.GetWindow(this);
                    warning.ShowDialog();
                    return;
                }

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
                string loai = planningType == "Chuyên môn" ? "ChuyenMon" : "Dang";
                saveFileDialog.FileName = $"QuyHoach_{loai}_{DateTime.Now:yyyyMMdd}.xlsx";
                
                if (saveFileDialog.ShowDialog() == true)
                {
                    string filePath = saveFileDialog.FileName;
                    LoadingOverlay.Visibility = Visibility.Visible;
                    try
                    {
                        await Task.Run(() =>
                        {
                            PlanningExcelExporter.Export(records, filePath);
                        });

                        LoadingOverlay.Visibility = Visibility.Collapsed;
                        var success = new SuccessWindow("Xuất file Excel thành công!", null, filePath, true);
                        success.Owner = Window.GetWindow(this);
                        success.ShowDialog();
                    }
                    catch (Exception ex)
                    {
                        LoadingOverlay.Visibility = Visibility.Collapsed;
                        var warning = new WarningWindow($"Lỗi xuất Excel: {ex.Message}", "Lỗi");
                        warning.Owner = Window.GetWindow(this);
                        warning.ShowDialog();
                    }
                }
            }
            catch (Exception ex)
            {
                var warning = new WarningWindow($"Lỗi xuất Excel: {ex.Message}", "Lỗi");
                warning.Owner = Window.GetWindow(this);
                warning.ShowDialog();
            }
        }

        #endregion
    }
}
