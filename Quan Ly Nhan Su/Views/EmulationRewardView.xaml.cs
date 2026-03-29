using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using TaxPersonnelManagement.Data;
using TaxPersonnelManagement.Models;
using ClosedXML.Excel;
using Microsoft.Win32;

namespace TaxPersonnelManagement.Views
{
    public partial class EmulationRewardView : Page
    {
        private List<Personnel> _allPersonnel = new();
        private Personnel? _selectedPersonnel;

        public EmulationRewardView()
        {
            InitializeComponent();
            LoadData();
        }

        private void LoadData()
        {
            try
            {
                using (var db = new AppDbContext())
                {
                    _allPersonnel = db.Personnel.OrderBy(p => p.FullName).ToList();
                    ApplyFilter();
                }
            }
            catch (Exception ex)
            {
                App.DebugLog("Error loading Personnel in EmulationRewardView: " + ex.Message);
            }
        }

        private void TxtSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            ApplyFilter();
        }

        private void ApplyFilter()
        {
            if (_allPersonnel == null) return;
            
            string keyword = txtSearch.Text.Trim().ToLower();
            if (string.IsNullOrEmpty(keyword))
            {
                lvPersonnel.ItemsSource = _allPersonnel;
            }
            else
            {
                lvPersonnel.ItemsSource = _allPersonnel.Where(p => 
                    (p.FullName != null && p.FullName.ToLower().Contains(keyword)) ||
                    (p.IdentityCardNumber != null && p.IdentityCardNumber.ToLower().Contains(keyword)) ||
                    (p.StaffId != null && p.StaffId.ToLower().Contains(keyword))
                ).ToList();
            }
        }

        private void LvPersonnel_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            _selectedPersonnel = lvPersonnel.SelectedItem as Personnel;

            if (_selectedPersonnel != null)
            {
                pnlDetails.IsEnabled = true;
                pnlDetails.Opacity = 1.0;
                txtPrompt.Visibility = Visibility.Collapsed;
                
                txtEmulationTitles.Text = _selectedPersonnel.EmulationTitles;
                txtRewardForms.Text = _selectedPersonnel.RewardForms;
            }
            else
            {
                pnlDetails.IsEnabled = false;
                pnlDetails.Opacity = 0.5;
                txtPrompt.Visibility = Visibility.Visible;
                
                txtEmulationTitles.Text = string.Empty;
                txtRewardForms.Text = string.Empty;
            }
        }

        private void BtnExportExcel_Click(object sender, RoutedEventArgs e)
        {
            if (_allPersonnel == null || !_allPersonnel.Any())
            {
                MessageBox.Show("Không có dữ liệu để xuất!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = $"DSHieuThiDua_KhenThuong_{DateTime.Now:yyyyMMdd}.xlsx"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("Thi đua khen thưởng");

                        // 1. Tiêu đề
                        string title = "DANH SÁCH TỔNG HỢP DANH HIỆU THI ĐUA & KHEN THƯỞNG CÔNG CHỨC";
                        var titleRange = worksheet.Range("A1:E1");
                        titleRange.Merge().Value = title;
                        titleRange.Style.Font.Bold = true;
                        titleRange.Style.Font.FontSize = 16;
                        titleRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        titleRange.Style.Font.FontColor = XLColor.FromHtml("#1565C0");
                        worksheet.Row(1).Height = 35;

                        // 2. Tiêu đề cột
                        string[] headers = { "STT", "Họ và tên", "Số CCCD", "Danh hiệu thi đua", "Hình thức khen thưởng" };
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

                        // 3. Dữ liệu
                        int currentRow = 3;
                        int stt = 1;
                        foreach (var p in _allPersonnel)
                        {
                            worksheet.Cell(currentRow, 1).Value = stt++;
                            worksheet.Cell(currentRow, 2).Value = p.FullName;
                            worksheet.Cell(currentRow, 3).Value = "'" + p.IdentityCardNumber;
                            worksheet.Cell(currentRow, 4).Value = p.EmulationTitles;
                            worksheet.Cell(currentRow, 5).Value = p.RewardForms;

                            // Kẻ bảng & Căn lề
                            for (int i = 1; i <= 5; i++)
                            {
                                var cell = worksheet.Cell(currentRow, i);
                                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                                cell.Style.Alignment.WrapText = true;
                                if (i == 1 || i == 3) cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            }
                            currentRow++;
                        }

                        // 4. Định dạng cột
                        worksheet.Column(1).Width = 8;   // STT
                        worksheet.Column(2).Width = 30;  // Họ tên
                        worksheet.Column(3).Width = 20;  // CCCD
                        worksheet.Column(4).Width = 50;  // Danh hiệu
                        worksheet.Column(5).Width = 50;  // Khen thưởng

                        workbook.SaveAs(saveFileDialog.FileName);
                        
                        // Thông báo thành công giống các menu khác
                        var success = new SuccessWindow("Xuất danh sách thi đua khen thưởng thành công!", saveFileDialog.FileName);
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
}
