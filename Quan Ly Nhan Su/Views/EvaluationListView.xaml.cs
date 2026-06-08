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

namespace TaxPersonnelManagement.Views
{
    public partial class EvaluationListView : UserControl
    {
        private List<EvaluationRecord> _fullEvaluationList = new List<EvaluationRecord>();
        private int _currentPage = 1;
        private const int PageSize = 20;
        private int _totalPages = 1;
        private System.Windows.Threading.DispatcherTimer? _searchDebounceTimer;

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
                MessageBox.Show($"Lỗi tải danh mục bộ phận: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            // 2. Month Filter
            var months = new List<FilterItem>();
            months.Add(new FilterItem { Label = "-- Tất cả các tháng --", Value = 0 });
            for (int i = 1; i <= 12; i++)
            {
                months.Add(new FilterItem { Label = $"Tháng {i}", Value = i });
            }
            cbMonth.ItemsSource = months;
            cbMonth.SelectedIndex = 0;

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
            for (int y = currentYear - 5; y <= currentYear + 2; y++)
            {
                if (!dbYears.Contains(y))
                    dbYears.Add(y);
            }

            foreach (var y in dbYears.OrderByDescending(x => x))
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
                int month = (cbMonth.SelectedItem as FilterItem)?.Value ?? 0;
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

                    if (month > 0)
                    {
                        filtered = filtered.Where(e => e.DecisionDate.HasValue && e.DecisionDate.Value.Month == month);
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
                MessageBox.Show($"Lỗi tải dữ liệu xếp loại: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
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

        private string GenerateExportFileName()
        {
            string yearPart = "";
            if (cbYear.SelectedItem is FilterItem yi && yi.Value > 0)
                yearPart = $"_Nam{yi.Value}";

            string monthPart = "";
            if (cbMonth.SelectedItem is FilterItem mi && mi.Value > 0)
                monthPart = $"_Thang{mi.Value}";

            return $"DanhSachXepLoai{yearPart}{monthPart}_{DateTime.Now:yyyyMMdd}.xlsx";
        }

        private void btnExport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var data = _fullEvaluationList;
                if (data == null || !data.Any())
                {
                    MessageBox.Show("Không có dữ liệu để xuất!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
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
                        string yearPartText = filterYear > 0 ? filterYear.ToString() : DateTime.Now.Year.ToString();
                        string title = $"DANH SÁCH KẾT QUẢ ĐÁNH GIÁ, XẾP LOẠI CHẤT LƯỢNG CÔNG CHỨC NĂM {yearPartText}";
                        
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

                                // Compile Note
                                var notesList = new List<string>();
                                if (!string.IsNullOrEmpty(item.DecisionNumber)) notesList.Add($"QĐ số {item.DecisionNumber}");
                                if (item.DecisionDate.HasValue) notesList.Add($"ngày {DatePickerHelper.FormatDateForDisplay(item.DecisionDate.Value)}");
                                if (!string.IsNullOrEmpty(item.DecisionAgency)) notesList.Add($"của {item.DecisionAgency}");
                                worksheet.Cell(currentRow, 9).Value = string.Join(" ", notesList);

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
                MessageBox.Show($"Có lỗi khi xuất Excel: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
