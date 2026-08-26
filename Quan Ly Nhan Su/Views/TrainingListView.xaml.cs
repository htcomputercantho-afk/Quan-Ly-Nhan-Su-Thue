using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Microsoft.EntityFrameworkCore;
using TaxPersonnelManagement.Data;
using TaxPersonnelManagement.Models;
using ClosedXML.Excel;
using Microsoft.Win32;

namespace TaxPersonnelManagement.Views
{
    public partial class TrainingListView : Page
    {
        private List<TrainingClassItem> _allClasses = new();
        private List<TrainingClassItem> _displayedClasses = new();
        private bool _isInitializingFilter = false;

        public TrainingListView()
        {
            InitializeComponent();
            LoadData();
        }

        public void LoadData()
        {
            try
            {
                using (var db = new AppDbContext())
                {
                    var classes = db.TrainingClasses
                                    .Include(tc => tc.PersonnelTrainings)
                                    .ToList();

                    int stt = 1;
                    _allClasses = classes.Select(tc => new TrainingClassItem
                    {
                        Id = tc.Id,
                        STT = stt++, // Will re-assign during filter
                        ClassName = tc.ClassName,
                        ParticipationDate = tc.ParticipationDate,
                        DecisionNumber = tc.DecisionNumber,
                        DecisionDate = tc.DecisionDate,
                        DecisionUnit = tc.DecisionUnit,
                        ParticipantCount = tc.PersonnelTrainings.Count
                    })
                    .OrderByDescending(tc => tc.DecisionDate ?? tc.ParticipationDate ?? DateTime.MinValue)
                    .ToList();

                    // Load Available Years for Filter
                    var distinctYears = _allClasses
                        .Select(tc => tc.Year)
                        .Where(y => y.HasValue)
                        .Select(y => y!.Value)
                        .Distinct()
                        .ToList();

                    int currentYear = DateTime.Now.Year;
                    if (!distinctYears.Contains(currentYear))
                    {
                        distinctYears.Add(currentYear);
                    }
                    distinctYears = distinctYears.OrderByDescending(y => y).ToList();

                    var yearItems = new List<TrainingYearFilterItem>
                    {
                        new TrainingYearFilterItem { Label = "-- Tất cả --", Value = null }
                    };
                    foreach (var y in distinctYears)
                    {
                        yearItems.Add(new TrainingYearFilterItem { Label = y.ToString(), Value = y });
                    }

                    int? prevYear = (cbYear.SelectedItem as TrainingYearFilterItem)?.Value;

                    _isInitializingFilter = true;
                    cbYear.ItemsSource = yearItems;
                    if (prevYear.HasValue && yearItems.Any(i => i.Value == prevYear.Value))
                    {
                        cbYear.SelectedItem = yearItems.First(i => i.Value == prevYear.Value);
                    }
                    else
                    {
                        var currentYearItem = yearItems.FirstOrDefault(i => i.Value == currentYear);
                        if (currentYearItem != null)
                        {
                            cbYear.SelectedItem = currentYearItem;
                        }
                        else
                        {
                            cbYear.SelectedIndex = 0; // "-- Tất cả --"
                        }
                    }
                    _isInitializingFilter = false;

                    ApplyFilter();
                }
            }
            catch (Exception ex)
            {
                App.DebugLog("Error loading TrainingClasses: " + ex.Message);
                var warning = new WarningWindow($"Lỗi tải danh sách lớp học: {ex.Message}", "Lỗi");
                if (Window.GetWindow(this) is Window parent) warning.Owner = parent;
                warning.ShowDialog();
            }
        }

        private void txtSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            ApplyFilter();
        }

        private void cbFilter_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_isInitializingFilter)
            {
                ApplyFilter();
            }
        }

        private void ApplyFilter()
        {
            string keyword = txtSearch.Text.Trim();
            int? selectedYear = (cbYear.SelectedItem as TrainingYearFilterItem)?.Value;

            var query = _allClasses.AsEnumerable();

            if (selectedYear.HasValue)
            {
                query = query.Where(tc => tc.Year == selectedYear.Value);
            }

            if (!string.IsNullOrEmpty(keyword))
            {
                query = query.Where(tc =>
                    TaxPersonnelManagement.Helpers.SearchHelper.IsMatch(tc.ClassName, keyword) ||
                    TaxPersonnelManagement.Helpers.SearchHelper.IsMatch(tc.DecisionNumber, keyword) ||
                    TaxPersonnelManagement.Helpers.SearchHelper.IsMatch(tc.DecisionUnit, keyword) ||
                    (tc.Year.HasValue && tc.Year.Value.ToString().Contains(keyword))
                );
            }

            var filtered = query.ToList();

            // Re-assign STT for visual correctness
            int stt = 1;
            foreach (var tc in filtered)
            {
                tc.STT = stt++;
            }

            _displayedClasses = filtered;
            dgTrainingClasses.ItemsSource = null;
            dgTrainingClasses.ItemsSource = filtered;
            txtTotalCount.Text = $"Hiển thị {filtered.Count} lớp học / hội nghị" + (selectedYear.HasValue ? $" (Năm {selectedYear.Value})" : "");
        }

        private void btnAddClass_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new AddTrainingClassDialog();
            if (Window.GetWindow(this) is Window parent)
            {
                dialog.Owner = parent;
            }
            if (dialog.ShowDialog() == true)
            {
                LoadData();
            }
        }

        private void btnDetail_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is int classId)
            {
                var dialog = new TrainingClassDetailDialog(classId);
                if (Window.GetWindow(this) is Window parent)
                {
                    dialog.Owner = parent;
                }
                if (dialog.ShowDialog() == true || dialog.IsDataChanged)
                {
                    LoadData();
                }
            }
        }

        private void btnExport_Click(object sender, RoutedEventArgs e)
        {
            if (!_displayedClasses.Any())
            {
                var warning = new WarningWindow("Không có dữ liệu để xuất!", "Thông báo");
                if (Window.GetWindow(this) is Window parent) warning.Owner = parent;
                warning.ShowDialog();
                return;
            }

            int? selectedYear = (cbYear.SelectedItem as TrainingYearFilterItem)?.Value;
            string fileNameSuffix = selectedYear.HasValue ? $"Nam{selectedYear.Value}_" : "";

            var saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = $"DanhSachLopDaoTao_HoiNghi_{fileNameSuffix}{DateTime.Now:yyyyMMdd}.xlsx"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("Lớp Đào Tạo - Hội Nghị");

                        // Title
                        string title = selectedYear.HasValue
                            ? $"DANH SÁCH TỔNG HỢP CÁC LỚP ĐÀO TẠO, BỒI DƯỠNG & HỘI NGHỊ NĂM {selectedYear.Value}"
                            : "DANH SÁCH TỔNG HỢP CÁC LỚP ĐÀO TẠO, BỒI DƯỠNG & HỘI NGHỊ";

                        worksheet.Cell(1, 1).Value = title;
                        worksheet.Cell(1, 1).Style.Font.Bold = true;
                        worksheet.Cell(1, 1).Style.Font.FontSize = 16;
                        worksheet.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml("#1565C0");
                        worksheet.Range("A1:H1").Merge().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Row(1).Height = 35;

                        // Headers
                        string[] headers = { "STT", "Năm", "Tên các lớp, hội nghị", "Ngày tham gia", "Số QĐ", "Ngày ra QĐ", "Đơn vị ra QĐ", "Số lượng học viên" };
                        for (int i = 0; i < headers.Length; i++)
                        {
                            var cell = worksheet.Cell(2, i + 1);
                            cell.Value = headers[i];
                            cell.Style.Font.Bold = true;
                            cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#1565C0");
                            cell.Style.Font.FontColor = XLColor.White;
                            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        }
                        worksheet.Row(2).Height = 25;

                        // Data
                        int currentRow = 3;
                        int stt = 1;
                        foreach (var tc in _displayedClasses)
                        {
                            worksheet.Cell(currentRow, 1).Value = stt++;
                            worksheet.Cell(currentRow, 2).Value = tc.YearDisplay;
                            worksheet.Cell(currentRow, 3).Value = tc.ClassName;
                            worksheet.Cell(currentRow, 4).Value = tc.ParticipationDate?.ToString("dd/MM/yyyy") ?? "";
                            worksheet.Cell(currentRow, 5).Value = tc.DecisionNumber;
                            worksheet.Cell(currentRow, 6).Value = tc.DecisionDate?.ToString("dd/MM/yyyy") ?? "";
                            worksheet.Cell(currentRow, 7).Value = tc.DecisionUnit;
                            worksheet.Cell(currentRow, 8).Value = tc.ParticipantCount;

                            // Formats
                            for (int i = 1; i <= 8; i++)
                            {
                                var cell = worksheet.Cell(currentRow, i);
                                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                                if (i == 1 || i == 2 || i == 4 || i == 6 || i == 8) cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            }
                            currentRow++;
                        }

                        // Column widths
                        worksheet.Column(1).Width = 8;
                        worksheet.Column(2).Width = 10;
                        worksheet.Column(3).Width = 40;
                        worksheet.Column(4).Width = 18;
                        worksheet.Column(5).Width = 18;
                        worksheet.Column(6).Width = 18;
                        worksheet.Column(7).Width = 25;
                        worksheet.Column(8).Width = 20;

                        workbook.SaveAs(saveFileDialog.FileName);

                        var success = new SuccessWindow("Xuất danh sách đào tạo và bồi dưỡng thành công!", null, saveFileDialog.FileName, true);
                        if (Window.GetWindow(this) is Window parent) success.Owner = parent;
                        success.ShowDialog();
                    }
                }
                catch (Exception ex)
                {
                    var warning = new WarningWindow($"Lỗi xuất Excel: {ex.Message}", "Lỗi");
                    if (Window.GetWindow(this) is Window parent) warning.Owner = parent;
                    warning.ShowDialog();
                }
            }
        }
    }

    public class TrainingClassItem
    {
        public int Id { get; set; }
        public int STT { get; set; }
        public string ClassName { get; set; } = string.Empty;
        public DateTime? ParticipationDate { get; set; }
        public string? DecisionNumber { get; set; }
        public DateTime? DecisionDate { get; set; }
        public string? DecisionUnit { get; set; }
        public int ParticipantCount { get; set; }

        public int? Year => DecisionDate?.Year ?? ParticipationDate?.Year;
        public string YearDisplay => Year.HasValue ? Year.Value.ToString() : "---";
    }

    public class TrainingYearFilterItem
    {
        public string Label { get; set; } = string.Empty;
        public int? Value { get; set; }
        public override string ToString() => Label;
    }
}
