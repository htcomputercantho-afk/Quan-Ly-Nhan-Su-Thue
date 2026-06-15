using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Microsoft.EntityFrameworkCore;
using TaxPersonnelManagement.Data;
using TaxPersonnelManagement.Models;
using TaxPersonnelManagement.Helpers;
using Microsoft.Win32;
using ClosedXML.Excel;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

namespace TaxPersonnelManagement.Views
{
    public partial class EvaluationListView : UserControl
    {
        private List<EvaluationRecord> _fullEvaluationList = new List<EvaluationRecord>();
        private int _currentPage = 1;
        private const int PageSize = 20;
        private int _totalPages = 1;
        private System.Windows.Threading.DispatcherTimer? _searchDebounceTimer;
        public List<int> AvailableYears { get; set; } = new List<int>();

        public EvaluationListView()
        {
            InitializeComponent();
            LoadFilterOptions();

            // Set up search debounce timer (300ms)
            _searchDebounceTimer = new System.Windows.Threading.DispatcherTimer();
            _searchDebounceTimer.Interval = TimeSpan.FromMilliseconds(300);
            _searchDebounceTimer.Tick += (s, e) =>
            {
                _searchDebounceTimer.Stop();
                LoadData();
            };

            LoadData();
        }

        public class FilterItem
        {
            public string Label { get; set; } = string.Empty;
            public int Value { get; set; }
            public override string ToString() => Label;
        }

        private void LoadFilterOptions()
        {
            // 1. Department Filter
            var deptOrder = new List<string> {
                "Ban lãnh đạo",
                "Tổ Hành chính, tổng hợp",
                "Tổ Kiểm tra số 1",
                "Tổ Kiểm tra số 2",
                "Tổ Kiểm tra số 3",
                "Tổ Nghiệp vụ, dự toán, pháp chế",
                "Tổ Quản lý các khoản thu khác",
                "Tổ Quản lý, hỗ trợ cá nhân, hộ kinh doanh số 1",
                "Tổ Quản lý, hỗ trợ cá nhân, hộ kinh doanh số 2",
                "Tổ Quản lý, hỗ trợ doanh nghiệp số 1",
                "Tổ Quản lý, hỗ trợ doanh nghiệp số 2"
            };

            try
            {
                using (var context = new AppDbContext())
                {
                    var allDepts = context.Departments
                                          .Select(d => d.Name)
                                          .Where(x => !string.IsNullOrEmpty(x))
                                          .Distinct()
                                          .ToList()
                                          .OrderBy(x =>
                                          {
                                              int idx = deptOrder.FindIndex(d => d.Equals(x, StringComparison.OrdinalIgnoreCase));
                                              return idx == -1 ? 999 : idx;
                                          })
                                          .ThenBy(x => x)
                                          .ToList();

                    allDepts.Insert(0, "-- Tất cả bộ phận --");
                    cbDepartment.ItemsSource = allDepts;
                    cbDepartment.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                var warning = new WarningWindow($"Lỗi tải danh mục bộ phận: {ex.Message}", "Lỗi");
                if (Window.GetWindow(this) is Window p) warning.Owner = p;
                warning.ShowDialog();
            }

            // 3. Year Filter
            var years = new List<FilterItem>();
            years.Add(new FilterItem { Label = "-- Tất cả các năm --", Value = 0 });

            List<int> dbYears = new List<int>();
            try
            {
                using (var context = new AppDbContext())
                {
                    dbYears = context.EvaluationRecords.Select(r => r.Year).Distinct().ToList();
                }
            }
            catch (Exception ex)
            {
                App.DebugLog("Error fetching evaluation years: " + ex.Message);
            }

            int currentYear = DateTime.Now.Year;
            HashSet<int> allYears = new HashSet<int>();
            for (int y = 1980; y <= currentYear; y++)
            {
                allYears.Add(y);
            }
            foreach (var y in dbYears)
            {
                allYears.Add(y);
            }

            AvailableYears = allYears.OrderByDescending(x => x).ToList();

            foreach (var y in AvailableYears)
            {
                years.Add(new FilterItem { Label = $"Năm {y}", Value = y });
            }
            cbYear.ItemsSource = years;
            cbYear.SelectedIndex = 0;
        }

        private void LoadData()
        {
            try
            {
                string search = txtSearch.Text.Trim();
                string? dept = cbDepartment.SelectedItem as string;
                int year = (cbYear.SelectedItem as FilterItem)?.Value ?? 0;

                using (var context = new AppDbContext())
                {
                    var query = context.EvaluationRecords
                                       .Include(e => e.Personnel)
                                       .AsQueryable();

                    var rawList = query.ToList();
                    var filtered = rawList.AsEnumerable();

                    if (!string.IsNullOrEmpty(search))
                    {
                        filtered = filtered.Where(e => e.Personnel != null && (
                            SearchHelper.IsMatch(e.Personnel.FullName, search) ||
                            SearchHelper.IsMatch(e.Personnel.IdentityCardNumber, search)
                        ));
                    }

                    if (!string.IsNullOrEmpty(dept) && dept != "-- Tất cả bộ phận --")
                    {
                        filtered = filtered.Where(e => e.Personnel != null && e.Personnel.Department == dept);
                    }

                    if (year > 0)
                    {
                        filtered = filtered.Where(e => e.Year == year);
                    }

                    // Sort order: Department, then Position, then FullName, then Year desc
                    var deptOrder = new List<string> {
                        "Ban lãnh đạo",
                        "Tổ Hành chính, tổng hợp",
                        "Tổ Kiểm tra số 1",
                        "Tổ Kiểm tra số 2",
                        "Tổ Kiểm tra số 3",
                        "Tổ Nghiệp vụ, dự toán, pháp chế",
                        "Tổ Quản lý các khoản thu khác",
                        "Tổ Quản lý, hỗ trợ cá nhân, hộ kinh doanh số 1",
                        "Tổ Quản lý, hỗ trợ cá nhân, hộ kinh doanh số 2",
                        "Tổ Quản lý, hỗ trợ doanh nghiệp số 1",
                        "Tổ Quản lý, hỗ trợ doanh nghiệp số 2"
                    };

                    _fullEvaluationList = filtered.OrderBy(e => 
                    {
                        string d = (e.Personnel?.Department ?? "").Trim();
                        int idx = deptOrder.FindIndex(o => o.Equals(d, StringComparison.OrdinalIgnoreCase));
                        return idx == -1 ? 999 : idx;
                    })
                    .ThenBy(e => 
                    {
                        string pos = e.Personnel?.Position?.ToLower() ?? "";
                        string d = (e.Personnel?.Department ?? "").ToLower();
                        if (d.Contains("lãnh đạo"))
                        {
                            if (pos.Contains("trưởng") && !pos.Contains("phó") && !pos.Contains("quyền")) return 1;
                            if (pos.Contains("quyền")) return 2;
                            if (pos.Contains("phó")) return 3;
                        }
                        else
                        {
                            if (pos.Contains("tổ trưởng") && !pos.Contains("phó")) return 1;
                            if (pos.Contains("phó")) return 2;
                            if (pos.Contains("công chức")) return 3;
                        }
                        return 99;
                    })
                    .ThenBy(e => e.Personnel?.FullName)
                    .ThenByDescending(e => e.Year)
                    .ToList();

                    _currentPage = 1;
                    ApplyPagination();
                }
            }
            catch (Exception ex)
            {
                var warning = new WarningWindow($"Lỗi tải dữ liệu xếp loại: {ex.Message}", "Lỗi");
                if (Window.GetWindow(this) is Window p) warning.Owner = p;
                warning.ShowDialog();
            }
        }

        private void ApplyPagination()
        {
            int totalItems = _fullEvaluationList.Count;
            _totalPages = Math.Max(1, (int)Math.Ceiling((double)totalItems / PageSize));

            if (_currentPage > _totalPages) _currentPage = _totalPages;
            if (_currentPage < 1) _currentPage = 1;

            var pageItems = _fullEvaluationList
                .Skip((_currentPage - 1) * PageSize)
                .Take(PageSize)
                .ToList();

            int startIndex = (_currentPage - 1) * PageSize;
            for (int i = 0; i < pageItems.Count; i++)
            {
                pageItems[i].STT = startIndex + i + 1;
            }

            dgEvaluation.ItemsSource = pageItems;

            if (totalItems == 0)
            {
                txtPagingInfo.Text = "Không có dữ liệu xếp loại";
                txtPageInfo.Text = "0 / 0";
            }
            else
            {
                int from = startIndex + 1;
                int to = startIndex + pageItems.Count;
                txtPagingInfo.Text = $"Hiển thị {from} - {to} trên {totalItems} bản ghi xếp loại";
                txtPageInfo.Text = $"{_currentPage} / {_totalPages}";
            }

            btnPrevPage.IsEnabled = _currentPage > 1;
            btnNextPage.IsEnabled = _currentPage < _totalPages;

            if (dgEvaluation.Items.Count > 0)
            {
                dgEvaluation.ScrollIntoView(dgEvaluation.Items[0]);
            }
        }

        private void BtnPrevPage_Click(object sender, RoutedEventArgs e)
        {
            if (_currentPage > 1)
            {
                _currentPage--;
                ApplyPagination();
            }
        }

        private void BtnNextPage_Click(object sender, RoutedEventArgs e)
        {
            if (_currentPage < _totalPages)
            {
                _currentPage++;
                ApplyPagination();
            }
        }

        private void txtSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_searchDebounceTimer != null)
            {
                _searchDebounceTimer.Stop();
                _searchDebounceTimer.Start();
            }
            else
            {
                LoadData();
            }
        }

        private void cbFilter_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            LoadData();
        }

        private T? FindVisualChild<T>(DependencyObject obj) where T : DependencyObject
        {
            for (int i = 0; i < System.Windows.Media.VisualTreeHelper.GetChildrenCount(obj); i++)
            {
                DependencyObject child = System.Windows.Media.VisualTreeHelper.GetChild(obj, i);
                if (child != null && child is T t)
                    return t;
                else if (child != null)
                {
                    T? childOfChild = FindVisualChild<T>(child);
                    if (childOfChild != null)
                        return childOfChild;
                }
            }
            return null;
        }

        private T? FindVisualParent<T>(DependencyObject child) where T : DependencyObject
        {
            DependencyObject parentObject = System.Windows.Media.VisualTreeHelper.GetParent(child);
            if (parentObject == null) return null;
            if (parentObject is T parent)
                return parent;
            else
                return FindVisualParent<T>(parentObject);
        }

        private void DataGridCell_PreviewMouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (sender is DataGridCell cell && !cell.IsEditing && !cell.IsReadOnly)
            {
                var grid = FindVisualParent<DataGrid>(cell);
                if (grid != null)
                {
                    cell.Focus();
                    grid.BeginEdit();
                }
            }
        }

        private void ComboBox_Loaded(object sender, RoutedEventArgs e)
        {
            if (sender is ComboBox cb)
            {
                cb.IsDropDownOpen = true;
            }
        }

        private void DatePicker_Loaded(object sender, RoutedEventArgs e)
        {
            if (sender is DatePicker dp)
            {
                var textBox = FindVisualChild<TextBox>(dp);
                if (textBox != null)
                {
                    textBox.Focus();
                    textBox.SelectAll();
                }
            }
        }

        private void dgEvaluation_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.EditAction == DataGridEditAction.Cancel)
                return;

            var record = e.Row.Item as EvaluationRecord;
            if (record == null)
                return;

            string header = e.Column.Header?.ToString() ?? "";
            bool isUpdated = false;

            if (header == "SỐ QĐ")
            {
                var textBox = e.EditingElement as TextBox;
                if (textBox != null)
                {
                    string newValue = textBox.Text.Trim();
                    string? val = string.IsNullOrEmpty(newValue) ? null : newValue;
                    if (record.DecisionNumber != val)
                    {
                        record.DecisionNumber = val;
                        isUpdated = true;
                    }
                }
            }
            else if (header == "NĂM")
            {
                var comboBox = FindVisualChild<ComboBox>(e.EditingElement);
                if (comboBox != null && comboBox.SelectedItem is int newYear)
                {
                    int oldYear = 0;
                    using (var db = new AppDbContext())
                    {
                        var originalRecord = db.EvaluationRecords.AsNoTracking().FirstOrDefault(r => r.Id == record.Id);
                        if (originalRecord != null)
                        {
                            oldYear = originalRecord.Year;
                        }

                        // Check if duplicate year for this personnel
                        bool exists = db.EvaluationRecords.Any(r => r.PersonnelId == record.PersonnelId && r.Year == newYear && r.Id != record.Id);
                        if (exists)
                        {
                            var warning = new WarningWindow($"Nhân sự này đã có bản ghi xếp loại cho năm {newYear}!", "Thông báo");
                            if (Window.GetWindow(this) is Window p) warning.Owner = p;
                            warning.ShowDialog();

                            if (oldYear > 0)
                            {
                                record.Year = oldYear;
                            }

                            // Force refresh display safely after edit transaction is complete
                            Dispatcher.BeginInvoke(new Action(() => 
                            {
                                try
                                {
                                    var view = System.Windows.Data.CollectionViewSource.GetDefaultView(dgEvaluation.ItemsSource) as System.ComponentModel.IEditableCollectionView;
                                    if (view != null)
                                    {
                                        if (view.IsEditingItem)
                                        {
                                            view.CancelEdit();
                                        }
                                        dgEvaluation.Items.Refresh();
                                    }
                                }
                                catch (Exception refreshEx)
                                {
                                    App.DebugLog("Error refreshing DataGrid: " + refreshEx.Message);
                                }
                            }), System.Windows.Threading.DispatcherPriority.ApplicationIdle);
                            return;
                        }
                    }

                    if (oldYear != newYear)
                    {
                        record.Year = newYear;
                        isUpdated = true;
                    }
                }
            }
            else if (header == "NGÀY KÝ QĐ")
            {
                var datePicker = FindVisualChild<DatePicker>(e.EditingElement);
                if (datePicker != null)
                {
                    // Validate if the DatePicker has a validation error
                    if (Validation.GetHasError(datePicker))
                    {
                        var warning = new WarningWindow("Định dạng ngày tháng không hợp lệ (Ví dụ hợp lệ: 15/06/2026)", "Thông báo");
                        if (Window.GetWindow(this) is Window p) warning.Owner = p;
                        warning.ShowDialog();

                        using (var db = new AppDbContext())
                        {
                            var originalRecord = db.EvaluationRecords.AsNoTracking().FirstOrDefault(r => r.Id == record.Id);
                            if (originalRecord != null)
                            {
                                record.DecisionDate = originalRecord.DecisionDate;
                            }
                        }

                        // Force refresh display safely after edit transaction is complete
                        Dispatcher.BeginInvoke(new Action(() => 
                        {
                            try
                            {
                                var view = System.Windows.Data.CollectionViewSource.GetDefaultView(dgEvaluation.ItemsSource) as System.ComponentModel.IEditableCollectionView;
                                if (view != null)
                                {
                                    if (view.IsEditingItem)
                                    {
                                        view.CancelEdit();
                                    }
                                    dgEvaluation.Items.Refresh();
                                }
                            }
                            catch (Exception refreshEx)
                            {
                                App.DebugLog("Error refreshing DataGrid: " + refreshEx.Message);
                            }
                        }), System.Windows.Threading.DispatcherPriority.ApplicationIdle);
                        return;
                    }

                    record.DecisionDate = datePicker.SelectedDate;
                    isUpdated = true;
                }
            }
            else if (header == "ĐƠN VỊ RA QĐ")
            {
                var textBox = e.EditingElement as TextBox;
                if (textBox != null)
                {
                    string newValue = textBox.Text.Trim();
                    string? val = string.IsNullOrEmpty(newValue) ? null : newValue;
                    if (record.DecisionAgency != val)
                    {
                        record.DecisionAgency = val;
                        isUpdated = true;
                    }
                }
            }

            if (isUpdated)
            {
                try
                {
                    using (var db = new AppDbContext())
                    {
                        var dbRecord = db.EvaluationRecords.FirstOrDefault(r => r.Id == record.Id);
                        if (dbRecord != null)
                        {
                            dbRecord.Year = record.Year;
                            dbRecord.DecisionNumber = record.DecisionNumber;
                            dbRecord.DecisionDate = record.DecisionDate;
                            dbRecord.DecisionAgency = record.DecisionAgency;
                            db.SaveChanges();
                        }
                    }
                }
                catch (Exception ex)
                {
                    var warning = new WarningWindow($"Lỗi lưu dữ liệu tự động: {ex.Message}", "Lỗi");
                    if (Window.GetWindow(this) is Window p) warning.Owner = p;
                    warning.ShowDialog();
                }

                // Force refresh display safely after edit transaction is complete
                Dispatcher.BeginInvoke(new Action(() => 
                {
                    try
                    {
                        var view = System.Windows.Data.CollectionViewSource.GetDefaultView(dgEvaluation.ItemsSource) as System.ComponentModel.IEditableCollectionView;
                        if (view != null && !view.IsEditingItem && !view.IsAddingNew)
                        {
                            dgEvaluation.Items.Refresh();
                        }
                    }
                    catch (Exception refreshEx)
                    {
                        App.DebugLog("Error refreshing DataGrid: " + refreshEx.Message);
                    }
                }), System.Windows.Threading.DispatcherPriority.ApplicationIdle);
            }
        }

        private string GenerateExportFileName()
        {
            string yearPart = "";
            if (cbYear.SelectedItem is FilterItem yi && yi.Value > 0)
                yearPart = $"_Nam{yi.Value}";

            return $"DanhSachXepLoai{yearPart}_{DateTime.Now:yyyyMMdd}.xlsx";
        }

        private void btnExport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var data = _fullEvaluationList;
                if (data == null || !data.Any())
                {
                    var warning = new WarningWindow("Không có dữ liệu để xuất!", "Thông báo");
                    if (Window.GetWindow(this) is Window p) warning.Owner = p;
                    warning.ShowDialog();
                    return;
                }

                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel Files|*.xlsx",
                    FileName = GenerateExportFileName()
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("Xếp loại");

                        // Row 1 & 2: Header info
                        var agencyRow1 = worksheet.Range("A1:C1");
                        agencyRow1.Merge();
                        agencyRow1.Value = "THUẾ TỈNH QUẢNG NINH";
                        agencyRow1.Style.Font.FontSize = 11;
                        agencyRow1.Style.Font.Bold = false;
                        
                        var agencyRow2 = worksheet.Range("A2:C2");
                        agencyRow2.Merge();
                        agencyRow2.Value = "THUẾ CƠ SỞ 1 TỈNH QUẢNG NINH";
                        agencyRow2.Style.Font.FontSize = 11;
                        agencyRow2.Style.Font.Bold = true;

                        // Row 3: Merged title
                        int filterYear = (cbYear.SelectedItem as FilterItem)?.Value ?? 0;
                        string title = filterYear > 0 
                            ? $"DANH SÁCH KẾT QUẢ ĐÁNH GIÁ, XẾP LOẠI CHẤT LƯỢNG CÔNG CHỨC NĂM {filterYear}"
                            : "DANH SÁCH KẾT QUẢ ĐÁNH GIÁ, XẾP LOẠI CHẤT LƯỢNG CÔNG CHỨC QUA CÁC NĂM";
                        
                        var titleRange = worksheet.Range("A3:I3");
                        titleRange.Merge();
                        titleRange.Value = title.ToUpper();
                        titleRange.Style.Font.Bold = true;
                        titleRange.Style.Font.FontSize = 14;
                        titleRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        titleRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Row(3).Height = 30;

                        // Row 5 & 6: Headers
                        // Merge TT
                        var ttRange = worksheet.Range("A5:A6");
                        ttRange.Merge();
                        ttRange.Value = "TT";

                        // Merge Họ và tên
                        var nameRange = worksheet.Range("B5:B6");
                        nameRange.Merge();
                        nameRange.Value = "Họ và tên";

                        // Merge Chức vụ
                        var posRange = worksheet.Range("C5:C6");
                        posRange.Merge();
                        posRange.Value = "Chức vụ, chức danh";

                        // Merge CCCD
                        var cccdRange = worksheet.Range("D5:D6");
                        cccdRange.Merge();
                        cccdRange.Value = "CCCD";

                        // Merge Rating Header Group
                        var ratingHeaderRange = worksheet.Range("E5:H5");
                        ratingHeaderRange.Merge();
                        ratingHeaderRange.Value = "Kết quả đánh giá, xếp loại chất lượng";

                        // Rating subheaders
                        worksheet.Cell(6, 5).Value = "Hoàn thành xuất sắc nhiệm vụ";
                        worksheet.Cell(6, 6).Value = "Hoàn thành tốt nhiệm vụ";
                        worksheet.Cell(6, 7).Value = "Hoàn thành nhiệm vụ";
                        worksheet.Cell(6, 8).Value = "Không hoàn thành nhiệm vụ";

                        // Merge Notes
                        var notesHeaderRange = worksheet.Range("I5:I6");
                        notesHeaderRange.Merge();
                        notesHeaderRange.Value = "Ghi chú";

                        // Style all header cells (A5:I6)
                        var headerRange = worksheet.Range("A5:I6");
                        headerRange.Style.Font.Bold = true;
                        headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        headerRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        headerRange.Style.Alignment.WrapText = true;
                        headerRange.Style.Fill.BackgroundColor = XLColor.FromHtml("#F2F2F2");
                        
                        // Apply thin borders to all header cells individually
                        for (int r = 5; r <= 6; r++)
                        {
                            for (int c = 1; c <= 9; c++)
                            {
                                worksheet.Cell(r, c).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            }
                        }
                        worksheet.Row(5).Height = 25;
                        worksheet.Row(6).Height = 55;

                        // Row 7+: Populate Data grouped by Department
                        int currentRow = 7;
                        int stt = 1;

                        var groupedData = data.GroupBy(e => e.Personnel?.Department ?? "Không xác định")
                                              .OrderBy(g => 
                                              {
                                                  var deptOrder = new List<string> {
                                                      "Ban lãnh đạo",
                                                      "Tổ Hành chính, tổng hợp",
                                                      "Tổ Kiểm tra số 1",
                                                      "Tổ Kiểm tra số 2",
                                                      "Tổ Kiểm tra số 3",
                                                      "Tổ Nghiệp vụ, dự toán, pháp chế",
                                                      "Tổ Quản lý các khoản thu khác",
                                                      "Tổ Quản lý, hỗ trợ cá nhân, hộ kinh doanh số 1",
                                                      "Tổ Quản lý, hỗ trợ cá nhân, hộ kinh doanh số 2",
                                                      "Tổ Quản lý, hỗ trợ doanh nghiệp số 1",
                                                      "Tổ Quản lý, hỗ trợ doanh nghiệp số 2"
                                                  };
                                                  int idx = deptOrder.FindIndex(o => o.Equals(g.Key.Trim(), StringComparison.OrdinalIgnoreCase));
                                                  return idx == -1 ? 999 : idx;
                                              })
                                              .ToList();

                        foreach (var group in groupedData)
                        {
                            // Write Department Row
                            var deptCell = worksheet.Cell(currentRow, 2);
                            deptCell.Value = group.Key;
                            deptCell.Style.Font.Bold = true;
                            
                            // Fill borders for the department row (A to I)
                            for (int c = 1; c <= 9; c++)
                            {
                                var cell = worksheet.Cell(currentRow, c);
                                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#FAFAFA");
                            }
                            currentRow++;

                            // Write evaluations in this group
                            foreach (var item in group)
                            {
                                worksheet.Cell(currentRow, 1).Value = stt++;
                                worksheet.Cell(currentRow, 2).Value = item.Personnel?.FullName ?? "";
                                worksheet.Cell(currentRow, 3).Value = item.Personnel?.Position ?? "";
                                worksheet.Cell(currentRow, 4).Value = "'" + (item.Personnel?.IdentityCardNumber ?? "");

                                // Markers under rating subcolumns
                                string r = (item.Rating ?? "").Trim();
                                if (r.Equals("Hoàn thành xuất sắc nhiệm vụ", StringComparison.OrdinalIgnoreCase))
                                {
                                    worksheet.Cell(currentRow, 5).Value = 1;
                                }
                                else if (r.Equals("Hoàn thành tốt nhiệm vụ", StringComparison.OrdinalIgnoreCase))
                                {
                                    worksheet.Cell(currentRow, 6).Value = 1;
                                }
                                else if (r.Equals("Hoàn thành nhiệm vụ", StringComparison.OrdinalIgnoreCase))
                                {
                                    worksheet.Cell(currentRow, 7).Value = 1;
                                }
                                else if (r.Equals("Không hoàn thành nhiệm vụ", StringComparison.OrdinalIgnoreCase))
                                {
                                    worksheet.Cell(currentRow, 8).Value = 1;
                                }


                                // Style and borders
                                for (int c = 1; c <= 9; c++)
                                {
                                    var cell = worksheet.Cell(currentRow, c);
                                    cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                    cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                                    cell.Style.Alignment.WrapText = true;
                                    
                                    if (c == 1 || c == 4 || (c >= 5 && c <= 8))
                                    {
                                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                    }
                                    else
                                    {
                                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                                    }
                                }
                                currentRow++;
                            }
                        }

                        // Columns widths adjustment
                        worksheet.Column(1).Width = 6;   // TT
                        worksheet.Column(2).Width = 25;  // Họ và tên
                        worksheet.Column(3).Width = 22;  // Chức vụ
                        worksheet.Column(4).Width = 18;  // CCCD
                        worksheet.Column(5).Width = 12;  // Xuất sắc
                        worksheet.Column(6).Width = 12;  // Tốt
                        worksheet.Column(7).Width = 12;  // Hoàn thành
                        worksheet.Column(8).Width = 12;  // Không hoàn thành
                        worksheet.Column(9).Width = 35;  // Ghi chú

                        workbook.SaveAs(saveFileDialog.FileName);

                        var success = new SuccessWindow("Xuất danh sách xếp loại thành công!", null, saveFileDialog.FileName, true);
                        if (Window.GetWindow(this) is Window parent) success.Owner = parent;
                        success.ShowDialog();
                    }
                }
            }
            catch (Exception ex)
            {
                var warning = new WarningWindow($"Có lỗi khi xuất Excel: {ex.Message}", "Lỗi");
                if (Window.GetWindow(this) is Window p) warning.Owner = p;
                warning.ShowDialog();
            }
        }

        private async void btnImport_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog();
            dlg.DefaultExt = ".xlsx";
            dlg.Filter = "Excel Documents (.xlsx)|*.xlsx";

            if (dlg.ShowDialog() == true)
            {
                string filePath = dlg.FileName;
                try
                {
                    // Hiện hiệu ứng loading
                    LoadingOverlay.Visibility = Visibility.Visible;

                    // Open and parse the Excel file in a background thread to avoid UI freezing
                    var result = await Task.Run(() => ParseEvaluationExcel(filePath));

                    // Ẩn hiệu ứng loading sau khi đọc xong
                    LoadingOverlay.Visibility = Visibility.Collapsed;

                    if (result.Items == null || !result.Items.Any())
                    {
                        var warning = new WarningWindow("Không tìm thấy dữ liệu xếp loại hợp lệ trong file Excel.", "Thông báo");
                        if (Window.GetWindow(this) is Window p) warning.Owner = p;
                        warning.ShowDialog();
                        return;
                    }

                    // Confirm before saving
                    string confirmMsg = $"Tìm thấy {result.Items.Count} dòng dữ liệu xếp loại.\n" +
                                        $"- Định dạng file: {(result.IsExportedFormat ? "Mẫu chuẩn (Exported)" : "Danh sách tùy chỉnh (List)")}\n" +
                                        $"- Năm đánh giá phát hiện: {result.DetectedYear}\n\n" +
                                        "Bạn có chắc chắn muốn nhập dữ liệu này vào hệ thống? (Nếu đã tồn tại bản ghi cùng nhân sự và năm, hệ thống sẽ cập nhật xếp loại và thông tin quyết định mới nhất)";

                    var confirmDialog = new ConfirmDialog(confirmMsg);
                    if (Window.GetWindow(this) is Window parentWindow) confirmDialog.Owner = parentWindow;

                    if (confirmDialog.ShowDialog() == true)
                    {
                        // Hiện hiệu ứng loading khi lưu vào database
                        LoadingOverlay.Visibility = Visibility.Visible;

                        int successCount = 0;
                        int notFoundCount = 0;
                        int errorCount = 0;
                        var unmatchedNames = new List<string>();

                        await Task.Run(() =>
                        {
                            using (var db = new AppDbContext())
                            {
                                // Load all personnel to memory for fast lookup
                                var allPersonnel = db.Personnel.ToList();

                                foreach (var item in result.Items)
                                {
                                    try
                                    {
                                        // Find personnel by CCCD first
                                        Personnel? person = null;
                                        if (!string.IsNullOrEmpty(item.CCCD))
                                        {
                                            person = allPersonnel.FirstOrDefault(p => p.IdentityCardNumber == item.CCCD);
                                        }

                                        // Fallback to FullName (normalized spacing)
                                        if (person == null && !string.IsNullOrEmpty(item.FullName))
                                        {
                                            string normItemName = Regex.Replace(item.FullName.Trim(), @"\s+", " ");
                                            person = allPersonnel.FirstOrDefault(p => 
                                                Regex.Replace(p.FullName.Trim(), @"\s+", " ").Equals(normItemName, StringComparison.OrdinalIgnoreCase));
                                        }

                                        if (person != null)
                                        {
                                            // Check if EvaluationRecord already exists for this person and year
                                            var record = db.EvaluationRecords.FirstOrDefault(r => r.PersonnelId == person.Id && r.Year == item.Year);
                                            if (record != null)
                                            {
                                                // Update existing
                                                record.Rating = item.Rating;
                                                record.DecisionNumber = item.DecisionNumber;
                                                record.DecisionDate = item.DecisionDate;
                                                record.DecisionAgency = item.DecisionAgency;
                                            }
                                            else
                                            {
                                                // Create new
                                                db.EvaluationRecords.Add(new EvaluationRecord
                                                {
                                                    PersonnelId = person.Id,
                                                    Year = item.Year,
                                                    Rating = item.Rating,
                                                    DecisionNumber = item.DecisionNumber,
                                                    DecisionDate = item.DecisionDate,
                                                    DecisionAgency = item.DecisionAgency
                                                });
                                            }
                                            successCount++;
                                        }
                                        else
                                        {
                                            notFoundCount++;
                                            unmatchedNames.Add(item.FullName);
                                        }
                                    }
                                    catch (Exception)
                                    {
                                        errorCount++;
                                    }
                                }

                                if (successCount > 0)
                                {
                                    db.SaveChanges();
                                }
                            }
                        });

                        // Refresh view
                        LoadFilterOptions();
                        LoadData();

                        // Ẩn hiệu ứng loading sau khi hoàn thành
                        LoadingOverlay.Visibility = Visibility.Collapsed;

                        string subMsg = $"Đã nhập thành công: {successCount} dòng.\n";
                        if (notFoundCount > 0)
                        {
                            var distinctUnmatched = unmatchedNames.Distinct().ToList();
                            string unmatchedListStr = distinctUnmatched.Count <= 5 
                                ? string.Join(", ", distinctUnmatched) 
                                : string.Join(", ", distinctUnmatched.Take(5)) + $" và {distinctUnmatched.Count - 5} người khác";
                            
                            subMsg += $"Không tìm thấy nhân sự (bỏ qua): {notFoundCount} dòng ({unmatchedListStr}).\n";
                        }
                        if (errorCount > 0)
                        {
                            subMsg += $"Bị lỗi dòng dữ liệu: {errorCount} dòng.";
                        }

                        var successWindow = new SuccessWindow("Nhập dữ liệu Excel thành công!", subMsg);
                        if (Window.GetWindow(this) is Window p) successWindow.Owner = p;
                        successWindow.ShowDialog();
                    }
                }
                catch (Exception ex)
                {
                    // Ẩn hiệu ứng loading nếu xảy ra lỗi
                    LoadingOverlay.Visibility = Visibility.Collapsed;
                    var warning = new WarningWindow($"Lỗi khi đọc file Excel: {ex.Message}", "Lỗi");
                    if (Window.GetWindow(this) is Window p) warning.Owner = p;
                    warning.ShowDialog();
                }
            }
        }

        private class ParseResult
        {
            public List<EvaluationImportItem> Items { get; set; } = new List<EvaluationImportItem>();
            public bool IsExportedFormat { get; set; }
            public int DetectedYear { get; set; }
        }

        private class EvaluationImportItem
        {
            public string FullName { get; set; } = string.Empty;
            public string CCCD { get; set; } = string.Empty;
            public string Rating { get; set; } = string.Empty;
            public int Year { get; set; }
            public string? DecisionNumber { get; set; }
            public DateTime? DecisionDate { get; set; }
            public string? DecisionAgency { get; set; }
        }

        private bool HasMark(IXLCell cell)
        {
            if (cell.IsEmpty()) return false;
            string val = cell.Value.ToString().Trim();
            return !string.IsNullOrEmpty(val) && val != "0";
        }

        private ParseResult ParseEvaluationExcel(string filePath)
        {
            var result = new ParseResult();

            using (var stream = new System.IO.FileStream(filePath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite))
            using (var workbook = new XLWorkbook(stream))
            {
                if (workbook.Worksheets.Count == 0) return result;
                var worksheet = workbook.Worksheet(1);
                var range = worksheet.RangeUsed();
                if (range == null) return result;

                // Detect format
                bool isExportedFormat = false;
                int headerRowIndex = -1;

                // We search for headers in the first 15 rows
                int maxRow = Math.Min(15, worksheet.LastRowUsed()?.RowNumber() ?? 1);
                int maxCol = Math.Min(15, worksheet.LastColumnUsed()?.ColumnNumber() ?? 1);

                for (int r = 1; r <= maxRow; r++)
                {
                    for (int c = 1; c <= maxCol; c++)
                    {
                        string val = worksheet.Cell(r, c).Value.ToString().Trim().ToLower();
                        if (val.Contains("hoàn thành xuất sắc nhiệm vụ") || val.Contains("hoàn thành tốt nhiệm vụ"))
                        {
                            isExportedFormat = true;
                            headerRowIndex = r - 1; // row 5 typically
                            break;
                        }
                    }
                    if (isExportedFormat) break;
                }

                int detectedYear = DateTime.Now.Year;
                // Try to find the year in the first 10 rows (e.g., "NĂM 2026")
                for (int r = 1; r <= 10; r++)
                {
                    for (int c = 1; c <= maxCol; c++)
                    {
                        string val = worksheet.Cell(r, c).Value.ToString();
                        var matchYear = Regex.Match(val, @"NĂM\s*(\d{4})", RegexOptions.IgnoreCase);
                        if (matchYear.Success)
                        {
                            detectedYear = int.Parse(matchYear.Groups[1].Value);
                            break;
                        }
                    }
                }

                result.DetectedYear = detectedYear;
                result.IsExportedFormat = isExportedFormat;

                int nameCol = -1;
                int cccdCol = -1;
                int yearCol = -1;
                int ratingCol = -1;
                int decNoCol = -1;
                int decDateCol = -1;
                int decAgencyCol = -1;

                if (!isExportedFormat)
                {
                    // Search for list format headers
                    for (int r = 1; r <= Math.Min(10, worksheet.LastRowUsed()?.RowNumber() ?? 1); r++)
                    {
                        for (int c = 1; c <= (worksheet.LastColumnUsed()?.ColumnNumber() ?? 1); c++)
                        {
                            string val = worksheet.Cell(r, c).Value.ToString().Trim().ToLower();
                            if (val == "họ và tên" || val == "họ tên" || val == "tên" || val == "tên công chức" || val == "họ và tên công chức")
                            {
                                nameCol = c;
                                headerRowIndex = r;
                            }
                            else if (val == "cccd" || val == "số cccd" || val == "số định danh" || val == "mã định danh")
                            {
                                cccdCol = c;
                            }
                            else if (val == "năm" || val == "năm đánh giá" || val == "năm xếp loại")
                            {
                                yearCol = c;
                            }
                            else if (val == "xếp loại" || val == "kết quả xếp loại" || val == "đánh giá" || val == "kết quả")
                            {
                                ratingCol = c;
                            }
                            else if (val == "số quyết định" || val == "số qđ")
                            {
                                decNoCol = c;
                            }
                            else if (val == "ngày quyết định" || val == "ngày qđ" || val == "ngày ký qđ")
                            {
                                decDateCol = c;
                            }
                            else if (val == "cơ quan quyết định" || val == "cơ quan ban hành" || val == "nơi ký" || val == "đơn vị ra qđ")
                            {
                                decAgencyCol = c;
                            }
                        }
                        if (nameCol != -1 && ratingCol != -1)
                        {
                            break; // Found headers
                        }
                    }
                }

                int startRow = isExportedFormat ? headerRowIndex + 2 : headerRowIndex + 1;
                int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;

                for (int r = startRow; r <= lastRow; r++)
                {
                    try
                    {
                        if (isExportedFormat)
                        {
                            string name = worksheet.Cell(r, 2).Value.ToString().Trim();
                            string cccd = worksheet.Cell(r, 4).Value.ToString().Trim();

                            // Skip empty rows
                            if (string.IsNullOrEmpty(name) && string.IsNullOrEmpty(cccd))
                                continue;

                            // Normalize name spacing
                            if (!string.IsNullOrEmpty(name))
                            {
                                name = Regex.Replace(name, @"\s+", " ");
                            }

                            // Clean CCCD formatting
                            if (!string.IsNullOrEmpty(cccd))
                            {
                                cccd = Regex.Replace(cccd, @"\s+", "");
                                if (cccd.EndsWith(".0")) cccd = cccd.Substring(0, cccd.Length - 2);
                            }

                            // Check if this is a Department header row
                            if (string.IsNullOrEmpty(cccd) && (name.Contains("Tổ") || name.Contains("Ban") || name.Contains("Phòng")))
                            {
                                continue; // Skip department group row
                            }

                            string rating = "";
                            if (HasMark(worksheet.Cell(r, 5))) rating = "Hoàn thành xuất sắc nhiệm vụ";
                            else if (HasMark(worksheet.Cell(r, 6))) rating = "Hoàn thành tốt nhiệm vụ";
                            else if (HasMark(worksheet.Cell(r, 7))) rating = "Hoàn thành nhiệm vụ";
                            else if (HasMark(worksheet.Cell(r, 8))) rating = "Không hoàn thành nhiệm vụ";

                            if (string.IsNullOrEmpty(rating))
                                continue; // Skip if no rating is selected

                            string? decNo = null;
                            DateTime? decDate = null;
                            string? decAgency = null;
                            int rowYear = detectedYear;

                            result.Items.Add(new EvaluationImportItem
                            {
                                FullName = name,
                                CCCD = cccd,
                                Rating = rating,
                                Year = rowYear,
                                DecisionNumber = decNo,
                                DecisionDate = decDate,
                                DecisionAgency = decAgency
                            });
                        }
                        else
                        {
                            // List format
                            if (nameCol == -1 || ratingCol == -1) continue;

                            string name = worksheet.Cell(r, nameCol).Value.ToString().Trim();
                            string ratingText = worksheet.Cell(r, ratingCol).Value.ToString().Trim();

                            if (string.IsNullOrEmpty(name) && string.IsNullOrEmpty(ratingText)) continue;

                            // Normalize name spacing
                            if (!string.IsNullOrEmpty(name))
                            {
                                name = Regex.Replace(name, @"\s+", " ");
                            }

                            // Map ratingText to standard ratings
                            string rating = "";
                            if (ratingText.Contains("xuất sắc", StringComparison.OrdinalIgnoreCase)) rating = "Hoàn thành xuất sắc nhiệm vụ";
                            else if (ratingText.Contains("tốt", StringComparison.OrdinalIgnoreCase)) rating = "Hoàn thành tốt nhiệm vụ";
                            else if (ratingText.Contains("không hoàn thành", StringComparison.OrdinalIgnoreCase)) rating = "Không hoàn thành nhiệm vụ";
                            else if (ratingText.Contains("hoàn thành", StringComparison.OrdinalIgnoreCase)) rating = "Hoàn thành nhiệm vụ";
                            else continue;

                            string cccd = cccdCol != -1 ? worksheet.Cell(r, cccdCol).Value.ToString().Trim() : "";
                            if (!string.IsNullOrEmpty(cccd))
                            {
                                cccd = Regex.Replace(cccd, @"\s+", "");
                                if (cccd.EndsWith(".0")) cccd = cccd.Substring(0, cccd.Length - 2);
                            }

                            int rowYear = detectedYear;
                            if (yearCol != -1 && !worksheet.Cell(r, yearCol).IsEmpty())
                            {
                                if (int.TryParse(worksheet.Cell(r, yearCol).Value.ToString().Trim(), out int y))
                                {
                                    rowYear = y;
                                }
                            }

                            string? decNo = decNoCol != -1 ? worksheet.Cell(r, decNoCol).Value.ToString().Trim() : null;
                            DateTime? decDate = null;
                            if (decDateCol != -1 && !worksheet.Cell(r, decDateCol).IsEmpty())
                            {
                                var cell = worksheet.Cell(r, decDateCol);
                                if (cell.DataType == XLDataType.DateTime)
                                {
                                    decDate = cell.GetDateTime();
                                }
                                else
                                {
                                    string dtStr = cell.Value.ToString().Trim();
                                    if (DateTime.TryParse(dtStr, out DateTime dDate))
                                    {
                                        decDate = dDate;
                                    }
                                }
                            }
                            string? decAgency = decAgencyCol != -1 ? worksheet.Cell(r, decAgencyCol).Value.ToString().Trim() : null;

                            result.Items.Add(new EvaluationImportItem
                            {
                                FullName = name,
                                CCCD = cccd,
                                Rating = rating,
                                Year = rowYear,
                                DecisionNumber = decNo,
                                DecisionDate = decDate,
                                DecisionAgency = decAgency
                            });
                        }
                    }
                    catch (Exception)
                    {
                        // Ignore row error
                    }
                }
            }

            return result;
        }
    }
}
