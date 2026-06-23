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
                    .OrderByDescending(tc => tc.ParticipationDate ?? DateTime.MinValue)
                    .ToList();

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

        private void ApplyFilter()
        {
            string keyword = txtSearch.Text.Trim();
            List<TrainingClassItem> filtered;

            if (string.IsNullOrEmpty(keyword))
            {
                filtered = _allClasses;
            }
            else
            {
                filtered = _allClasses.Where(tc =>
                    TaxPersonnelManagement.Helpers.SearchHelper.IsMatch(tc.ClassName, keyword) ||
                    TaxPersonnelManagement.Helpers.SearchHelper.IsMatch(tc.DecisionNumber, keyword) ||
                    TaxPersonnelManagement.Helpers.SearchHelper.IsMatch(tc.DecisionUnit, keyword)
                ).ToList();
            }

            // Re-assign STT for visual correctness
            int stt = 1;
            foreach (var tc in filtered)
            {
                tc.STT = stt++;
            }

            dgTrainingClasses.ItemsSource = null;
            dgTrainingClasses.ItemsSource = filtered;
            txtTotalCount.Text = $"Hiển thị {filtered.Count} lớp học / hội nghị";
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
            if (!_allClasses.Any())
            {
                var warning = new WarningWindow("Không có dữ liệu để xuất!", "Thông báo");
                if (Window.GetWindow(this) is Window parent) warning.Owner = parent;
                warning.ShowDialog();
                return;
            }

            var saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = $"DanhSachLopDaoTao_HoiNghi_{DateTime.Now:yyyyMMdd}.xlsx"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("Lớp Đào Tạo - Hội Nghị");

                        // Title
                        worksheet.Cell(1, 1).Value = "DANH SÁCH TỔNG HỢP CÁC LỚP ĐÀO TẠO, BỒI DƯỠNG & HỘI NGHỊ";
                        worksheet.Cell(1, 1).Style.Font.Bold = true;
                        worksheet.Cell(1, 1).Style.Font.FontSize = 16;
                        worksheet.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml("#1565C0");
                        worksheet.Range("A1:G1").Merge().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Row(1).Height = 35;

                        // Headers
                        string[] headers = { "STT", "Tên các lớp, hội nghị", "Ngày tham gia", "Số QĐ", "Ngày ra QĐ", "Đơn vị ra QĐ", "Số lượng học viên" };
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
                        foreach (var tc in _allClasses)
                        {
                            worksheet.Cell(currentRow, 1).Value = stt++;
                            worksheet.Cell(currentRow, 2).Value = tc.ClassName;
                            worksheet.Cell(currentRow, 3).Value = tc.ParticipationDate?.ToString("dd/MM/yyyy") ?? "";
                            worksheet.Cell(currentRow, 4).Value = tc.DecisionNumber;
                            worksheet.Cell(currentRow, 5).Value = tc.DecisionDate?.ToString("dd/MM/yyyy") ?? "";
                            worksheet.Cell(currentRow, 6).Value = tc.DecisionUnit;
                            worksheet.Cell(currentRow, 7).Value = tc.ParticipantCount;

                            // Formats
                            for (int i = 1; i <= 7; i++)
                            {
                                var cell = worksheet.Cell(currentRow, i);
                                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                                if (i == 1 || i == 3 || i == 5 || i == 7) cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            }
                            currentRow++;
                        }

                        // Column widths
                        worksheet.Column(1).Width = 8;
                        worksheet.Column(2).Width = 40;
                        worksheet.Column(3).Width = 18;
                        worksheet.Column(4).Width = 18;
                        worksheet.Column(5).Width = 18;
                        worksheet.Column(6).Width = 25;
                        worksheet.Column(7).Width = 20;

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
    }
}
