using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Microsoft.EntityFrameworkCore;
using TaxPersonnelManagement.Data;
using TaxPersonnelManagement.Models;
using Microsoft.Win32;
using ClosedXML.Excel;
using TaxPersonnelManagement.Helpers;

namespace TaxPersonnelManagement.Views
{
    public partial class PositionDurationView : UserControl
    {
        private List<Personnel> _fullList = new List<Personnel>();
        private List<Personnel> _filteredList = new List<Personnel>();
        private int _currentPage = 1;
        private const int PageSize = 20;
        private int _totalPages = 1;

        public PositionDurationView()
        {
            InitializeComponent();
            dpCalculationDate.SelectedDate = DateTime.Now;
            // The localization fix is now handled via Attached Property in XAML:
            // Helpers:DatePickerHelper.FixCalendarLocale="True"
            LoadData();
        }

        private void dpCalculationDate_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            dgPositionDuration.Items.Refresh();
        }

        private void LoadData()
        {
            try
            {
                using (var context = new AppDbContext())
                {
                    var query = context.Personnel.AsQueryable();

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

                    _fullList = query.AsEnumerable()
                                    .OrderBy(p => {
                                        string dept = (p.Department ?? "").Trim();
                                        int index = deptOrder.FindIndex(d => d.Equals(dept, StringComparison.OrdinalIgnoreCase));
                                        return index == -1 ? 999 : index;
                                    })
                                    .ThenBy(p => {
                                        string pos = p.Position?.ToLower() ?? "";
                                        string dept3 = (p.Department ?? "").ToLower();

                                        if (dept3.Contains("lãnh đạo"))
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
                                    .ThenBy(p => p.FullName)
                                    .ToList();

                    _filteredList = _fullList;
                    _currentPage = 1;
                    ApplyPagination();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi tải dữ liệu: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void txtSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            ApplyFilter();
        }

        private void ApplyFilter()
        {
            string keyword = txtSearch.Text.Trim().ToLower();
            if (string.IsNullOrEmpty(keyword))
            {
                _filteredList = _fullList;
            }
            else
            {
                _filteredList = _fullList.Where(p => 
                    (p.FullName ?? "").ToLower().Contains(keyword) || 
                    (p.Department ?? "").ToLower().Contains(keyword) || 
                    (p.Position ?? "").ToLower().Contains(keyword)
                ).ToList();
            }

            _currentPage = 1;
            ApplyPagination();
        }

        private void ApplyPagination()
        {
            int totalItems = _filteredList.Count;
            _totalPages = Math.Max(1, (int)Math.Ceiling((double)totalItems / PageSize));

            if (_currentPage > _totalPages) _currentPage = _totalPages;
            if (_currentPage < 1) _currentPage = 1;

            var pageItems = _filteredList
                .Skip((_currentPage - 1) * PageSize)
                .Take(PageSize)
                .ToList();

            dgPositionDuration.ItemsSource = pageItems;

            // Update footer
            if (totalItems == 0)
            {
                txtPagingInfo.Text = "Không có dữ liệu";
                txtPageInfo.Text = "0 / 0";
            }
            else
            {
                int startIndex = (_currentPage - 1) * PageSize;
                int from = startIndex + 1;
                int to = startIndex + pageItems.Count;
                txtPagingInfo.Text = $"Hiển thị {from} - {to} trên {totalItems} nhân sự";
                txtPageInfo.Text = $"{_currentPage} / {_totalPages}";
            }

            btnPrevPage.IsEnabled = _currentPage > 1;
            btnNextPage.IsEnabled = _currentPage < _totalPages;
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

        private void dg_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = ((_currentPage - 1) * PageSize + e.Row.GetIndex() + 1).ToString();
        }

        private void btnViewDetail_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.DataContext is Personnel p)
            {
                if (Application.Current.MainWindow is MainWindow mw)
                {
                    // Tab 2 (History Info) is index 1
                    mw.NavigateToPersonnelDetail(p, 1);
                }
            }
        }

        private void btnExport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var data = _filteredList;
                if (data == null || !data.Any())
                {
                    MessageBox.Show("Không có dữ liệu để xuất!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                DateTime referenceDate = dpCalculationDate.SelectedDate ?? DateTime.Now;
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel Files|*.xlsx",
                    FileName = $"ThoiGianGiuViTri_{referenceDate:yyyyMMdd}.xlsx"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("Thoi gian giu vi tri");

                        // Title
                        var titleRange = worksheet.Range("A1:E1");
                        titleRange.Merge();
                        titleRange.Value = "DANH SÁCH THỜI GIAN GIỮ VỊ TRÍ CÔNG TÁC";
                        titleRange.Style.Font.Bold = true;
                        titleRange.Style.Font.FontSize = 14;
                        titleRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        titleRange.Style.Font.FontColor = XLColor.DarkBlue;

                        // Subtitle with Reference Date
                        var subtitleRange = worksheet.Range("A2:E2");
                        subtitleRange.Merge();
                        subtitleRange.Value = $"(Tính đến ngày: {referenceDate:dd/MM/yyyy})";
                        subtitleRange.Style.Font.Italic = true;
                        subtitleRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                        // Headers
                        string[] headers = { "STT", "Họ và tên", "Bộ phận", "Ngày quyết định", "Thời gian giữ vị trí" };
                        for (int i = 0; i < headers.Length; i++)
                        {
                            var cell = worksheet.Cell(4, i + 1);
                            cell.Value = headers[i];
                            cell.Style.Font.Bold = true;
                            cell.Style.Fill.BackgroundColor = XLColor.LightGray;
                            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        }

                        // Data
                        int row = 5;
                        int stt = 1;

                        foreach (var item in data)
                        {
                            worksheet.Cell(row, 1).Value = stt++;
                            worksheet.Cell(row, 2).Value = item.FullName;
                            worksheet.Cell(row, 3).Value = item.Department;
                            
                            string durationText = "";
                            if (item.PositionDecisionDate.HasValue)
                            {
                                worksheet.Cell(row, 4).Value = item.PositionDecisionDate.Value;
                                
                                // Calculate duration
                                DateTime start = item.PositionDecisionDate.Value;
                                if (referenceDate >= start)
                                {
                                    int years = referenceDate.Year - start.Year;
                                    if (start.Date > referenceDate.AddYears(-years)) years--;

                                    DateTime tmpDate = start.AddYears(years);
                                    int months = 0;
                                    while (tmpDate.AddMonths(1) <= referenceDate)
                                    {
                                        months++;
                                        tmpDate = tmpDate.AddMonths(1);
                                    }
                                    durationText = $"{years} năm {months} tháng";
                                }
                                else
                                {
                                    durationText = "0 năm 0 tháng";
                                }
                            }
                            
                            worksheet.Cell(row, 5).Value = durationText;

                            for (int col = 1; col <= 5; col++)
                                worksheet.Cell(row, col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                            row++;
                        }

                        worksheet.Columns().AdjustToContents();
                        workbook.SaveAs(saveFileDialog.FileName);
                        
                        var successWindow = new SuccessWindow("Xuất Excel thành công!", "File đã được lưu tại:", saveFileDialog.FileName, true);
                        successWindow.Owner = Window.GetWindow(this);
                        successWindow.ShowDialog();
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
