using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using TaxPersonnelManagement.Models;
using TaxPersonnelManagement.Data;
using System.Collections.ObjectModel;
using Microsoft.EntityFrameworkCore;
using ClosedXML.Excel;
using Microsoft.Win32;

namespace TaxPersonnelManagement.Views
{
    public partial class LeaveDetailView : Page
    {
        private List<LeaveSummaryItem> _allSummaries = new List<LeaveSummaryItem>();
        private List<LeaveSummaryItem> _filteredSummaries = new List<LeaveSummaryItem>();
        private bool _isUpdating = false;

        // Pagination
        private int _currentPage = 1;
        private const int PageSize = 20;
        private int _totalPages = 1;

        public LeaveDetailView()
        {
            InitializeComponent();
            LoadAllLeaveSummaries();
        }
        
        private void Grid_PreviewMouseWheel(object sender, System.Windows.Input.MouseWheelEventArgs e)
        {
            if (!e.Handled)
            {
                e.Handled = true;
                var eventArg = new System.Windows.Input.MouseWheelEventArgs(e.MouseDevice, e.Timestamp, e.Delta);
                eventArg.RoutedEvent = UIElement.MouseWheelEvent;
                eventArg.Source = sender;
                var parent = ((Control)sender).Parent as UIElement;
                parent?.RaiseEvent(eventArg);
            }
        }



        private void TxtSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (btnClearSearch != null)
            {
                btnClearSearch.Visibility = string.IsNullOrEmpty(txtSearch.Text) ? Visibility.Collapsed : Visibility.Visible;
            }
            ApplyFilter();
        }

        private void BtnClearSearch_Click(object sender, RoutedEventArgs e)
        {
            txtSearch.Clear();
            txtSearch.Focus();
        }

        private void ApplyFilter()
        {
            if (_allSummaries == null) return;

            string keyword = txtSearch.Text.Trim().ToLower();
            _filteredSummaries = _allSummaries.Where(s => 
                (s.FullName != null && s.FullName.ToLower().Contains(keyword)) ||
                (s.StaffId != null && s.StaffId.ToLower().Contains(keyword)) ||
                (s.IdentityCardNumber != null && s.IdentityCardNumber.ToLower().Contains(keyword))
            ).ToList();

            // Re-assign STT for the entire filtered list
            for (int i = 0; i < _filteredSummaries.Count; i++)
            {
                _filteredSummaries[i].STT = i + 1;
            }

            _currentPage = 1;
            ApplyPagination();
        }

        private void ApplyPagination()
        {
            int totalItems = _filteredSummaries.Count;
            _totalPages = Math.Max(1, (int)Math.Ceiling((double)totalItems / PageSize));

            if (_currentPage > _totalPages) _currentPage = _totalPages;
            if (_currentPage < 1) _currentPage = 1;

            var pageItems = _filteredSummaries
                .Skip((_currentPage - 1) * PageSize)
                .Take(PageSize)
                .ToList();

            // Re-number STT across pages
            int startIndex = (_currentPage - 1) * PageSize;
            for (int i = 0; i < pageItems.Count; i++)
            {
                pageItems[i].STT = startIndex + i + 1;
            }

            dgLeaveSummary.ItemsSource = pageItems;

            // Update footer
            if (totalItems == 0)
            {
                txtPagingInfo.Text = "Không có dữ liệu";
                txtPageInfo.Text = "0 / 0";
            }
            else
            {
                int from = startIndex + 1;
                int to = startIndex + pageItems.Count;
                txtPagingInfo.Text = $"Hiển thị {from} - {to} trên {totalItems} nhân viên";
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



        private void LoadAllLeaveSummaries()
        {
            if (_isUpdating) return;
            _isUpdating = true;

            try
            {
                int selectedYear = DateTime.Now.Year;
                txtHeaderTitle.Text = $"BẢNG TỔNG HỢP SỐ NGÀY NGHỈ PHÉP THỰC TẾ NĂM {selectedYear}";
                
                // Update DataGrid dynamic headers
                colOldYear.Header = $"Số ngày phép nghỉ theo chế độ năm {selectedYear - 1}";
                colCurrentYear.Header = $"Số ngày phép nghỉ theo chế độ năm {selectedYear}";

                using (var db = new AppDbContext())
                {
                    var personnelList = db.Personnel.Include(p => p.LeaveHistories)
                                          .OrderBy(p => p.FullName)
                                          .ToList();
                    
                    var results = new List<LeaveSummaryItem>();
                    int index = 1;

                    foreach (var p in personnelList)
                    {
                        int totalAnnual = CalculateTotalAnnualLeave(p);
                        int prevYear = selectedYear - 1;
                        double takenFromOldYear = p.LeaveHistories
                            .Where(h => h.LeaveType == "Phép năm" && h.StartDate.Year == selectedYear && h.LeaveYear == prevYear)
                            .Sum(h => h.DurationDays);

                        double takenFromCurrentYear = p.LeaveHistories
                            .Where(h => h.LeaveType == "Phép năm" && h.StartDate.Year == selectedYear && (h.LeaveYear == selectedYear || h.LeaveYear == null))
                            .Sum(h => h.DurationDays);

                        double annualTaken = takenFromOldYear + takenFromCurrentYear;

                        var histories = p.LeaveHistories
                            .Where(h => h.LeaveType == "Phép năm" && h.StartDate.Year == selectedYear)
                            .OrderBy(h => h.StartDate)
                            .ToList();
                        
                        var detailLines = new List<LeaveDetailLine>();
                        for (int i = 0; i < histories.Count; i++)
                        {
                            var h = histories[i];
                            bool isOldYear = (h.LeaveYear == prevYear);
                            string reason = h.UserReasonDisplay;
                            string note = string.IsNullOrWhiteSpace(reason) ? "" : $": {reason}";
                            
                            detailLines.Add(new LeaveDetailLine
                            {
                                MainContent = $"Lần {i + 1}: {h.DurationDays:0.#} ngày từ ngày {h.StartDate:dd/MM/yyyy} đến ngày {h.EndDate:dd/MM/yyyy}{note}",
                                Suffix = isOldYear ? " - [NGHỈ PHÉP NĂM CŨ]" : ""
                            });
                        }
                        string detailedContent = string.Join("\n", detailLines.Select(d => d.MainContent + d.Suffix));

                        results.Add(new LeaveSummaryItem
                        {
                            STT = index++,
                            StaffId = p.StaffId ?? string.Empty,
                            IdentityCardNumber = p.IdentityCardNumber ?? string.Empty,
                            FullName = p.FullName ?? string.Empty,
                            TotalTarget = totalAnnual,
                            TakenFromOldYear = takenFromOldYear,
                            TakenFromCurrentYear = takenFromCurrentYear,
                            ActualTaken = annualTaken,
                            DetailedContent = detailedContent,
                            DetailLines = detailLines,
                            Remaining = totalAnnual - takenFromCurrentYear,
                            AvatarBase64 = p.AvatarBase64 ?? string.Empty
                        });
                    }

                    _allSummaries = results;
                    ApplyFilter();
                }
            }
            catch (Exception ex)
            {
                App.DebugLog("Error loading all leave summaries: " + ex.Message);
            }
            finally
            {
                _isUpdating = false;
            }
        }

        private int CalculateTotalAnnualLeave(Personnel p)
        {
            int baseDays = 12;
            if (p.TaxAuthorityStartDate.HasValue || p.StartDate.HasValue)
            {
                DateTime start = p.TaxAuthorityStartDate ?? p.StartDate!.Value;
                DateTime now = DateTime.Now;
                int years = now.Year - start.Year;
                if (now < start.AddYears(years)) years--;
                
                if (years < 0) years = 0;
                baseDays += (years / 5);
            }
            return baseDays;
        }

        private void BtnExportExcel_Click(object sender, RoutedEventArgs e)
        {
            if (_filteredSummaries == null || !_filteredSummaries.Any())
            {
                MessageBox.Show("Không có dữ liệu để xuất!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            int selectedYear = DateTime.Now.Year;
            var saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = $"BaoCao_PhepNam_{selectedYear}_{DateTime.Now:yyyyMMdd}.xlsx"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("Chi tiết phép năm");

                        // 1. Tiêu đề báo cáo
                        string title = $"BẢNG TỔNG HỢP SỐ NGÀY NGHỈ PHÉP THỰC TẾ NĂM {selectedYear}";
                        var titleRange = worksheet.Range("A1:I1");
                        titleRange.Merge();
                        titleRange.Value = title;
                        titleRange.Style.Font.Bold = true;
                        titleRange.Style.Font.FontSize = 16;
                        titleRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        titleRange.Style.Font.FontColor = XLColor.FromHtml("#1565C0");
                        titleRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Row(1).Height = 40;

                        // 2. Tiêu đề cột
                        // Row 2: Merged headers
                        var entRange = worksheet.Range("D2:D3");
                        entRange.Merge();
                        entRange.Value = "Số ngày được nghỉ phép";
                        entRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        entRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        entRange.Style.Font.Bold = true;
                        entRange.Style.Fill.BackgroundColor = XLColor.FromHtml("#1565C0");
                        entRange.Style.Font.FontColor = XLColor.White;
                        entRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        var takenRange = worksheet.Range("E2:G2");
                        takenRange.Merge();
                        takenRange.Value = $"Số ngày đã nghỉ phép năm {selectedYear}";
                        takenRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        takenRange.Style.Font.Bold = true;
                        takenRange.Style.Fill.BackgroundColor = XLColor.FromHtml("#1565C0");
                        takenRange.Style.Font.FontColor = XLColor.White;
                        takenRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        string[] row2Headers = { "STT", "Họ và tên", "Số CCCD" };
                        for (int i = 0; i < row2Headers.Length; i++)
                        {
                            var cell = worksheet.Range(worksheet.Cell(2, i + 1), worksheet.Cell(3, i + 1));
                            cell.Merge();
                            cell.Value = row2Headers[i];
                            cell.Style.Font.Bold = true;
                            cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#1565C0");
                            cell.Style.Font.FontColor = XLColor.White;
                            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        }

                        // Row 3: Sub-headers
                        worksheet.Cell(3, 5).Value = $"Số ngày phép nghỉ theo chế độ năm {selectedYear - 1}";
                        worksheet.Cell(3, 6).Value = $"Số ngày phép nghỉ theo chế độ năm {selectedYear}";
                        worksheet.Cell(3, 7).Value = "Tổng";
                        
                        var remRange = worksheet.Range("H2:H3");
                        remRange.Merge();
                        remRange.Value = "Phép còn lại";
                        remRange.Style.Font.Bold = true;
                        remRange.Style.Fill.BackgroundColor = XLColor.FromHtml("#1565C0");
                        remRange.Style.Font.FontColor = XLColor.White;
                        remRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        remRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        remRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        var detailRange = worksheet.Range("I2:I3");
                        detailRange.Merge();
                        detailRange.Value = "Nội dung chi tiết";
                        detailRange.Style.Font.Bold = true;
                        detailRange.Style.Fill.BackgroundColor = XLColor.FromHtml("#1565C0");
                        detailRange.Style.Font.FontColor = XLColor.White;
                        detailRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        detailRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        detailRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        for (int i = 5; i <= 7; i++)
                        {
                            var cell = worksheet.Cell(3, i);
                            cell.Style.Font.Bold = true;
                            cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#1565C0");
                            cell.Style.Font.FontColor = XLColor.White;
                            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            cell.Style.Alignment.WrapText = true;
                        }

                        worksheet.Row(2).Height = 25;
                        worksheet.Row(3).Height = 40;

                        // 3. Dữ liệu
                        int currentRow = 4;
                        foreach (var item in _filteredSummaries)
                        {
                            worksheet.Cell(currentRow, 1).Value = item.STT;
                            worksheet.Cell(currentRow, 2).Value = item.FullName;
                            worksheet.Cell(currentRow, 3).Value = "'" + item.IdentityCardNumber;
                            worksheet.Cell(currentRow, 4).Value = item.TotalTarget;
                            worksheet.Cell(currentRow, 5).Value = item.TakenFromOldYear;
                            worksheet.Cell(currentRow, 6).Value = item.TakenFromCurrentYear;
                            worksheet.Cell(currentRow, 7).Value = item.ActualTaken;
                            worksheet.Cell(currentRow, 8).Value = item.Remaining;
                            // RichText cho nội dung chi tiết
                            var richText = worksheet.Cell(currentRow, 9).GetRichText();
                            for (int i = 0; i < item.DetailLines.Count; i++)
                            {
                                var line = item.DetailLines[i];
                                
                                // Nội dung chính (không in đậm)
                                richText.AddText(line.MainContent);
                                
                                // Hậu tố (in đậm nếu là năm cũ)
                                if (!string.IsNullOrEmpty(line.Suffix))
                                {
                                    richText.AddText(line.Suffix).SetBold(true);
                                }
                                
                                if (i < item.DetailLines.Count - 1)
                                {
                                    richText.AddNewLine();
                                }
                            }

                            // Định dạng dòng
                            for (int i = 1; i <= 9; i++)
                            {
                                var cell = worksheet.Cell(currentRow, i);
                                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                                
                                if (i != 2 && i != 9) // Căn giữa trừ cột Tên và Nội dung chi tiết
                                {
                                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                }
                                
                                if (i == 9) // Warp text cho nội dung chi tiết
                                {
                                    cell.Style.Alignment.WrapText = true;
                                }
                            }

                            // Màu sắc cho các cột số liệu
                            worksheet.Cell(currentRow, 5).Style.Font.FontColor = XLColor.FromHtml("#2E7D32"); // Green
                            worksheet.Cell(currentRow, 6).Style.Font.FontColor = XLColor.FromHtml("#7B1FA2"); // Purple
                            worksheet.Cell(currentRow, 7).Style.Font.FontColor = XLColor.FromHtml("#0288D1"); // Blue
                            worksheet.Cell(currentRow, 8).Style.Font.FontColor = XLColor.FromHtml("#E65100"); // Orange
                            worksheet.Cell(currentRow, 9).Style.Font.FontColor = XLColor.FromHtml("#1B5E20"); // Dark Green
                            worksheet.Cell(currentRow, 9).Style.Font.Bold = true;

                            currentRow++;
                        }

                        // 4. Định dạng cột
                        worksheet.Column(1).Width = 10;  // STT
                        worksheet.Column(2).Width = 35;  // Họ và tên
                        worksheet.Column(3).Width = 25;  // Số CCCD
                        worksheet.Column(4).Width = 25;  // Số ngày được nghỉ phép
                        worksheet.Column(5).Width = 25;  // Số ngày phép năm 2025
                        worksheet.Column(6).Width = 25;  // Số ngày phép năm 2026
                        worksheet.Column(7).Width = 15;  // Tổng
                        worksheet.Column(8).Width = 15;  // Phép còn lại
                        worksheet.Column(9).Width = 85;  // Nội dung chi tiết

                        // Row heights & Wrap text for Headers
                        worksheet.Row(2).Height = 35;
                        worksheet.Row(2).Style.Alignment.WrapText = true;
                        worksheet.Row(3).Height = 55;
                        worksheet.Row(3).Style.Alignment.WrapText = true;

                        workbook.SaveAs(saveFileDialog.FileName);
                        var success = new SuccessWindow("Xuất báo cáo Excel thành công!", saveFileDialog.FileName);
                        if (Window.GetWindow(this) is Window parent) success.Owner = parent;
                        success.ShowDialog();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Lỗi khi xuất Excel: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
    }

    public class LeaveDetailLine
    {
        public string MainContent { get; set; } = string.Empty;
        public string Suffix { get; set; } = string.Empty; // "[NGHỈ PHÉP NĂM CŨ]"
        public bool IsOldYear => !string.IsNullOrEmpty(Suffix);
    }

    public class LeaveSummaryItem
    {
        public int STT { get; set; }
        public string StaffId { get; set; } = string.Empty;
        public string IdentityCardNumber { get; set; } = string.Empty;
        public string FullName { get; set; } = string.Empty;
        public int TotalTarget { get; set; }
        public double TakenFromOldYear { get; set; }
        public double TakenFromCurrentYear { get; set; }
        public double ActualTaken { get; set; }
        public string DetailedContent { get; set; } = string.Empty;
        public List<LeaveDetailLine> DetailLines { get; set; } = new List<LeaveDetailLine>();
        public double Remaining { get; set; }
        public string? AvatarBase64 { get; set; }
    }
}
