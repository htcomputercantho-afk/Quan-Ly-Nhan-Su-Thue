using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using TaxPersonnelManagement.Models;
using TaxPersonnelManagement.Data;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Collections.ObjectModel;
using System.IO;
using Microsoft.Win32;
using ExcelDataReader;
using System.Text.RegularExpressions;
using System.Data;
using ClosedXML.Excel;
namespace TaxPersonnelManagement.Views
{
    public partial class AnnualIncomeView : Page
    {
        private List<Personnel> _allPersonnel = new List<Personnel>();
        private ObservableCollection<AnnualIncomeRowViewModel> _matrixData = new ObservableCollection<AnnualIncomeRowViewModel>();
        private bool _isUpdating = false;

        public AnnualIncomeView()
        {
            InitializeComponent();
            LoadData();
            LoadYears();
            dgMonthlyIncome.ItemsSource = _matrixData;
            ApplyAuthorization();
        }

        private void ApplyAuthorization()
        {
            if (App.CurrentUser?.Role == UserRole.Staff)
            {
                // Hide action buttons
                btnBulkAdd.Visibility = Visibility.Collapsed;
                btnClearData.Visibility = Visibility.Collapsed;
            }
        }

        private void AdminOnly_Loaded(object sender, RoutedEventArgs e)
        {
            if (App.CurrentUser?.Role == UserRole.Staff && sender is FrameworkElement element)
            {
                element.Visibility = Visibility.Collapsed;
            }
        }

        private void LoadData()
        {
            try
            {
                using var db = new AppDbContext();
                // Load personnel list
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

                _allPersonnel = db.Personnel.AsEnumerable()
                                  .OrderBy(p =>
                                  {
                                      string dept = (p.Department ?? "").Trim();
                                      int index = deptOrder.FindIndex(d => d.Equals(dept, StringComparison.OrdinalIgnoreCase));
                                      return index == -1 ? 999 : index;
                                  })
                                  .ThenBy(p =>
                                  {
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

                lvPersonnel.ItemsSource = _allPersonnel;
            }
            catch (Exception ex)
            {
                App.DebugLog("Error loading personnel for income view: " + ex.Message);
            }
        }

        private void LoadYears()
        {
            int currentYear = DateTime.Now.Year;
            var years = new List<int>();

            try
            {
                using (var db = new AppDbContext())
                {
                    years = db.IncomeRecords.Select(r => r.Year).Distinct().ToList();
                }
            }
            catch (Exception ex)
            {
                App.DebugLog("Error loading years: " + ex.Message);
            }

            if (!years.Contains(currentYear))
            {
                years.Add(currentYear);
            }

            years.Sort((a, b) => b.CompareTo(a)); // Sort descending

            int? selected = cboYear.SelectedItem as int?;

            cboYear.ItemsSource = years;

            if (selected.HasValue && years.Contains(selected.Value))
            {
                cboYear.SelectedItem = selected.Value;
            }
            else
            {
                cboYear.SelectedItem = currentYear;
            }
        }

        private void TxtSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            string keyword = txtSearch.Text.Trim();
            lvPersonnel.ItemsSource = _allPersonnel.Where(p =>
                TaxPersonnelManagement.Helpers.SearchHelper.IsMatch(p.FullName, keyword) ||
                TaxPersonnelManagement.Helpers.SearchHelper.IsMatch(p.StaffId, keyword)).ToList();
        }

        private void LvPersonnel_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            GenerateDummyData();
        }

        private void CboYear_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            GenerateDummyData();
        }

        private void GenerateDummyData()
        {
            if (_isUpdating) return;
            _isUpdating = true;

            if (lvPersonnel.SelectedItem is Personnel selectedPerson)
            {
                _matrixData.Clear();
                int selectedYear = (int)(cboYear.SelectedItem ?? DateTime.Now.Year);

                var salaryRow = new AnnualIncomeRowViewModel { IncomeType = "Lương" };
                var overtimeRow = new AnnualIncomeRowViewModel { IncomeType = "Làm thêm giờ" };
                var otherRow = new AnnualIncomeRowViewModel { IncomeType = "Thu nhập khác" };
                var totalRow = new AnnualIncomeRowViewModel { IncomeType = "Tổng cộng (Tháng)", IsTotalRow = true };

                var rows = new[] { salaryRow, overtimeRow, otherRow };
                foreach (var row in rows)
                {
                    row.PropertyChanged += Row_PropertyChanged;
                }

                // Load from DB
                using (var db = new AppDbContext())
                {
                    var records = db.IncomeRecords
                        .Where(r => r.PersonnelId == selectedPerson.Id && r.Year == selectedYear)
                        .ToList();

                    foreach (var record in records)
                    {
                        AnnualIncomeRowViewModel? targetRow = null;
                        if (record.IncomeType == "Lương") targetRow = salaryRow;
                        else if (record.IncomeType == "Làm thêm giờ") targetRow = overtimeRow;
                        else if (record.IncomeType == "Thu nhập khác") targetRow = otherRow;

                        if (targetRow != null)
                        {
                            if (record.IncomeType == "Thu nhập khác" || record.IncomeType == "Làm thêm giờ")
                            {
                                string formattedNote = record.Note ?? "";
                                if (!string.IsNullOrWhiteSpace(formattedNote) && 
                                    !Regex.IsMatch(formattedNote, @"^\s*\d+([\.,]\d+)*\s*đ?\s*-") && 
                                    record.Amount > 0)
                                {
                                    formattedNote = $"{record.Amount:N0} đ - {formattedNote}";
                                }
                                targetRow.SetNote(record.Month, formattedNote);
                                targetRow.SetAmount(record.Month, record.Amount);
                            }
                            else
                            {
                                targetRow.SetAmount(record.Month, record.Amount);
                                targetRow.SetNote(record.Month, record.Note ?? "");
                            }
                        }
                    }
                }

                _matrixData.Add(salaryRow);
                _matrixData.Add(overtimeRow);
                _matrixData.Add(otherRow);
                _matrixData.Add(totalRow);

                UpdateTotals();
            }
            else
            {
                _matrixData.Clear();
                txtTotalSalary.Text = "0 đ";
                txtTotalOvertime.Text = "0 đ";
                txtTotalOther.Text = "0 đ";
                txtGrandTotal.Text = "0 đ";
            }

            _isUpdating = false;
        }

        private void Row_PropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            UpdateTotals();
        }

        private void DgMonthlyIncome_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            Dispatcher.BeginInvoke(new Action(() => UpdateTotals()), System.Windows.Threading.DispatcherPriority.Background);
        }

        private void DgMonthlyIncome_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
            if (e.Row.Item is AnnualIncomeRowViewModel row)
            {
                if (row.IsTotalRow)
                {
                    e.Cancel = true; // Prevent editing the total row
                    return;
                }

                if (row.IncomeType == "Làm thêm giờ" || row.IncomeType == "Thu nhập khác")
                {
                    e.Cancel = true; // Cancel default inline editing
                    
                    int month = 0;
                    var headerStr = e.Column.Header?.ToString() ?? "";
                    if (headerStr.StartsWith("Tháng "))
                    {
                        int.TryParse(headerStr.Substring(6), out month);
                    }

                    if (month >= 1 && month <= 12)
                    {
                        string currentNote = row.GetNote(month);
                        var dialog = new EditMonthlyIncomeDialog(row.IncomeType, month, currentNote);
                        if (Window.GetWindow(this) is Window parent) dialog.Owner = parent;
                        if (dialog.ShowDialog() == true)
                        {
                            string newNote = dialog.ResultNote;
                            row.SetNote(month, newNote);
                            
                            // Trigger property changes to update UI
                            row.TriggerMonthChanged(month);
                            UpdateTotals();
                        }
                    }
                }
            }
        }

        private void UpdateTotals()
        {
            if (_matrixData.Count != 4) return;

            var salaryRow = _matrixData[0];
            var overtimeRow = _matrixData[1];
            var otherRow = _matrixData[2];
            var totalRow = _matrixData[3];

            for (int m = 1; m <= 12; m++)
            {
                decimal totalForMonth = salaryRow.GetAmount(m) + overtimeRow.GetAmount(m) + otherRow.GetAmount(m);
                totalRow.SetAmount(m, totalForMonth);
            }

            decimal totalSalary = salaryRow.Total;
            decimal totalOvertime = overtimeRow.Total;
            decimal totalOther = otherRow.Total;

            txtTotalSalary.Text = totalSalary.ToString("N0") + " đ";
            txtTotalOvertime.Text = totalOvertime.ToString("N0") + " đ";
            txtTotalOther.Text = totalOther.ToString("N0") + " đ";
            txtGrandTotal.Text = (totalSalary + totalOvertime + totalOther).ToString("N0") + " đ";
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            if (lvPersonnel.SelectedItem is Personnel selectedPerson)
            {
                int selectedYear = (int)(cboYear.SelectedItem ?? DateTime.Now.Year);
                try
                {
                    using (var db = new AppDbContext())
                    {
                        var existingRecords = db.IncomeRecords
                            .Where(r => r.PersonnelId == selectedPerson.Id && r.Year == selectedYear)
                            .ToList();

                        db.IncomeRecords.RemoveRange(existingRecords);

                        var salaryRow = _matrixData[0];
                        var overtimeRow = _matrixData[1];
                        var otherRow = _matrixData[2];

                        for (int m = 1; m <= 12; m++)
                        {
                            if (salaryRow.GetAmount(m) > 0 || !string.IsNullOrWhiteSpace(salaryRow.GetNote(m)))
                                db.IncomeRecords.Add(new IncomeRecord { PersonnelId = selectedPerson.Id, Year = selectedYear, Month = m, IncomeType = "Lương", Amount = salaryRow.GetAmount(m), Note = salaryRow.GetNote(m) });

                            if (overtimeRow.GetAmount(m) > 0 || !string.IsNullOrWhiteSpace(overtimeRow.GetNote(m)))
                            {
                                string rawNote = overtimeRow.GetNote(m);
                                string stdNote = AnnualIncomeRowViewModel.StandardizeOtherIncomeNote(rawNote);
                                decimal recalculatedAmount = AnnualIncomeRowViewModel.ParseAmountFromNote(rawNote);
                                
                                db.IncomeRecords.Add(new IncomeRecord 
                                { 
                                    PersonnelId = selectedPerson.Id, 
                                    Year = selectedYear, 
                                    Month = m, 
                                    IncomeType = "Làm thêm giờ", 
                                    Amount = recalculatedAmount, 
                                    Note = stdNote 
                                });
                                
                                overtimeRow.SetNote(m, stdNote);
                                overtimeRow.SetAmount(m, recalculatedAmount);
                            }

                            if (otherRow.GetAmount(m) > 0 || !string.IsNullOrWhiteSpace(otherRow.GetNote(m)))
                            {
                                string rawNote = otherRow.GetNote(m);
                                string stdNote = AnnualIncomeRowViewModel.StandardizeOtherIncomeNote(rawNote);
                                decimal recalculatedAmount = AnnualIncomeRowViewModel.ParseAmountFromNote(rawNote);
                                
                                db.IncomeRecords.Add(new IncomeRecord 
                                { 
                                    PersonnelId = selectedPerson.Id, 
                                    Year = selectedYear, 
                                    Month = m, 
                                    IncomeType = "Thu nhập khác", 
                                    Amount = recalculatedAmount, 
                                    Note = stdNote 
                                });
                                
                                // Sync back to the row model to update UI immediately
                                otherRow.SetNote(m, stdNote);
                                otherRow.SetAmount(m, recalculatedAmount);
                            }
                        }

                        db.SaveChanges();

                        var successDialog = new SuccessWindow("Đã lưu dữ liệu thu nhập thành công!");
                        successDialog.Owner = Window.GetWindow(this);
                        successDialog.ShowDialog();
                    }
                }
                catch (Exception ex)
                {
                    var warningDialog = new WarningWindow("Lỗi khi lưu dữ liệu", ex.Message);
                    warningDialog.Owner = Window.GetWindow(this);
                    warningDialog.ShowDialog();
                }
            }
        }

        private void BtnClearData_Click(object sender, RoutedEventArgs e)
        {
            if (lvPersonnel.SelectedItem is Personnel selectedPerson)
            {
                int selectedYear = (int)(cboYear.SelectedItem ?? DateTime.Now.Year);
                var confirm = new ConfirmWindow($"Bạn có chắc chắn muốn xóa toàn bộ dữ liệu thu nhập năm {selectedYear} của công chức '{selectedPerson.FullName}' trên bảng không?\n(Dữ liệu dưới cơ sở dữ liệu sẽ chỉ bị xóa khi bạn nhấn 'Lưu dữ liệu')", "Xác nhận xóa dữ liệu năm");
                confirm.Owner = Window.GetWindow(this);
                if (confirm.ShowDialog() == true)
                {
                    foreach (var row in _matrixData)
                    {
                        if (row.IsTotalRow) continue;

                        for (int m = 1; m <= 12; m++)
                        {
                            row.SetAmount(m, 0);
                            row.SetNote(m, "");
                        }
                    }

                    UpdateTotals();
                    dgMonthlyIncome.Items.Refresh();
                }
            }
            else
            {
                var warning = new WarningWindow("Thông báo", "Vui lòng chọn một công chức từ danh sách trước.");
                warning.Owner = Window.GetWindow(this);
                warning.ShowDialog();
            }
        }

        private void AmountTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (sender is TextBox textBox)
            {
                textBox.TextChanged -= AmountTextBox_TextChanged;
                try
                {
                    string clean = Regex.Replace(textBox.Text, @"[^0-9]", "");
                    if (decimal.TryParse(clean, out decimal parsed))
                    {
                        string formatted = parsed.ToString("N0", System.Globalization.CultureInfo.GetCultureInfo("vi-VN"));
                        
                        int oldSelectionStart = textBox.SelectionStart;
                        int oldLength = textBox.Text.Length;

                        textBox.Text = formatted;

                        int newLength = formatted.Length;
                        int newSelectionStart = oldSelectionStart + (newLength - oldLength);
                        textBox.SelectionStart = Math.Max(0, Math.Min(newSelectionStart, newLength));
                    }
                    else
                    {
                        textBox.Text = "";
                    }
                }
                catch { }
                finally
                {
                    textBox.TextChanged += AmountTextBox_TextChanged;
                }
            }
        }

        private void BtnBulkAdd_Click(object sender, RoutedEventArgs e)
        {
            int selectedYear = (int)(cboYear.SelectedItem ?? DateTime.Now.Year);
            var dialog = new BulkIncomeDialog(selectedYear) { Owner = Window.GetWindow(this) };

            if (dialog.ShowDialog() == true)
            {
                // Refresh data
                LoadYears();
                GenerateDummyData();
            }
        }

        private async void BtnExportExcel_Click(object sender, RoutedEventArgs e)
        {
            int selectedYear = (int)(cboYear.SelectedItem ?? DateTime.Now.Year);

            var saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = $"Bao_Cao_Thu_Nhap_{selectedYear}.xlsx",
                Title = "Lưu báo cáo thu nhập năm"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                string filePath = saveFileDialog.FileName;
                LoadingOverlay.Visibility = Visibility.Visible;

                try
                {
                    await Task.Run(() =>
                    {
                        using (var db = new AppDbContext())
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

                            var personnelList = db.Personnel.AsEnumerable()
                                                  .OrderBy(p =>
                                                  {
                                                      string dept = (p.Department ?? "").Trim();
                                                      int index = deptOrder.FindIndex(d => d.Equals(dept, StringComparison.OrdinalIgnoreCase));
                                                      return index == -1 ? 999 : index;
                                                  })
                                                  .ThenBy(p =>
                                                  {
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
                            var allRecords = db.IncomeRecords.Where(r => r.Year == selectedYear).ToList();

                            using (var workbook = new XLWorkbook())
                            {
                                var worksheet = workbook.Worksheets.Add("Thu Nhập Năm " + selectedYear);

                                // Header row
                                worksheet.Cell(1, 1).Value = "BÁO CÁO TỔNG HỢP THU NHẬP NĂM " + selectedYear;
                                worksheet.Range(1, 1, 1, 6).Merge().Style.Font.SetBold().Font.SetFontSize(16).Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

                                var headers = new[] { "Họ và tên", "CCCD", "Lương (VNĐ)", "Làm thêm giờ (VNĐ)", "Thu nhập khác (VNĐ)", "Tổng cộng (VNĐ)" };
                                for (int i = 0; i < headers.Length; i++)
                                {
                                    var cell = worksheet.Cell(3, i + 1);
                                    cell.Value = headers[i];
                                    cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#1976D2");
                                    cell.Style.Font.FontColor = XLColor.White;
                                    cell.Style.Font.Bold = true;
                                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                    cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                }

                                int row = 4;
                                foreach (var person in personnelList)
                                {
                                    var personRecords = allRecords.Where(r => r.PersonnelId == person.Id).ToList();
                                    decimal salary = personRecords.Where(r => r.IncomeType == "Lương").Sum(r => r.Amount);
                                    decimal overtime = personRecords.Where(r => r.IncomeType == "Làm thêm giờ").Sum(r => r.Amount);
                                    decimal other = personRecords.Where(r => r.IncomeType == "Thu nhập khác").Sum(r => r.Amount);
                                    decimal total = salary + overtime + other;

                                    worksheet.Cell(row, 1).Value = person.FullName;
                                    worksheet.Cell(row, 2).Value = "'" + person.IdentityCardNumber; // Force string to avoid scientific notation
                                    worksheet.Cell(row, 3).Value = salary;
                                    worksheet.Cell(row, 4).Value = overtime;
                                    worksheet.Cell(row, 5).Value = other;
                                    worksheet.Cell(row, 6).Value = total;

                                    // Formatting numbers
                                    worksheet.Range(row, 3, row, 6).Style.NumberFormat.Format = "#,##0";
                                    worksheet.Range(row, 1, row, 6).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                                    row++;
                                }

                                // Footer: Grand Total
                                worksheet.Cell(row, 1).Value = "TỔNG CỘNG";
                                worksheet.Range(row, 1, row, 2).Merge().Style.Font.Bold = true;
                                worksheet.Cell(row, 3).FormulaA1 = $"=SUM(C4:C{row - 1})";
                                worksheet.Cell(row, 4).FormulaA1 = $"=SUM(D4:D{row - 1})";
                                worksheet.Cell(row, 5).FormulaA1 = $"=SUM(E4:E{row - 1})";
                                worksheet.Cell(row, 6).FormulaA1 = $"=SUM(F4:F{row - 1})";
                                worksheet.Range(row, 1, row, 6).Style.Font.Bold = true;
                                worksheet.Range(row, 3, row, 6).Style.NumberFormat.Format = "#,##0";
                                worksheet.Range(row, 1, row, 6).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                                // Auto-fit columns
                                worksheet.Columns().AdjustToContents();
                                worksheet.Column(2).Width = 20; // Fix CCCD column width

                                workbook.SaveAs(filePath);
                            }
                        }
                    });

                    LoadingOverlay.Visibility = Visibility.Collapsed;
                    var successDialog = new SuccessWindow("Xuất báo cáo Excel thành công!", null, filePath, true);
                    if (Window.GetWindow(this) is Window parent) successDialog.Owner = parent;
                    successDialog.ShowDialog();
                }
                catch (Exception ex)
                {
                    LoadingOverlay.Visibility = Visibility.Collapsed;
                    var warning = new WarningWindow("Lỗi khi xuất Excel", ex.Message);
                    warning.ShowDialog();
                }
            }
        }
    }

    public class IncomeBreakdownItem
    {
        public decimal Amount { get; set; }
        public string Reason { get; set; } = "";
        public string FormattedAmount => Amount > 0 ? Amount.ToString("N0") + " đ" : "";
    }

    public class AnnualIncomeRowViewModel : INotifyPropertyChanged
    {
        public string IncomeType { get; set; } = "";
        public bool IsTotalRow { get; set; } = false;

        public static decimal ParseAmountFromNote(string note)
        {
            if (string.IsNullOrWhiteSpace(note)) return 0;
            decimal total = 0;
            var lines = note.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var line in lines)
            {
                if (string.IsNullOrWhiteSpace(line)) continue;
                int dashIndex = line.IndexOf('-');
                string amtStr = dashIndex == -1 ? line : line.Substring(0, dashIndex);
                string cleanAmt = Regex.Replace(amtStr, @"[^0-9]", "");
                if (decimal.TryParse(cleanAmt, out decimal amount))
                {
                    total += amount;
                }
            }
            return total;
        }

        public static List<IncomeBreakdownItem> ParseBreakdown(string note)
        {
            var list = new List<IncomeBreakdownItem>();
            if (string.IsNullOrWhiteSpace(note)) return list;
            
            var lines = note.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var line in lines)
            {
                if (string.IsNullOrWhiteSpace(line)) continue;
                int dashIndex = line.IndexOf('-');
                if (dashIndex == -1)
                {
                    string cleanAmt = Regex.Replace(line, @"[^0-9]", "");
                    if (decimal.TryParse(cleanAmt, out decimal amount) && amount > 0)
                    {
                        list.Add(new IncomeBreakdownItem { Amount = amount, Reason = "" });
                    }
                    else
                    {
                        list.Add(new IncomeBreakdownItem { Amount = 0, Reason = line.Trim() });
                    }
                }
                else
                {
                    string amtStr = line.Substring(0, dashIndex);
                    string reasonStr = line.Substring(dashIndex + 1).Trim();
                    string cleanAmt = Regex.Replace(amtStr, @"[^0-9]", "");
                    decimal.TryParse(cleanAmt, out decimal amount);
                    list.Add(new IncomeBreakdownItem { Amount = amount, Reason = reasonStr });
                }
            }
            return list;
        }

        public static string StandardizeOtherIncomeNote(string note)
        {
            if (string.IsNullOrWhiteSpace(note)) return "";
            var lines = note.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            var standardizedLines = new List<string>();
            foreach (var line in lines)
            {
                if (string.IsNullOrWhiteSpace(line)) continue;
                int dashIndex = line.IndexOf('-');
                if (dashIndex == -1)
                {
                    string cleanAmt = Regex.Replace(line, @"[^0-9]", "");
                    if (decimal.TryParse(cleanAmt, out decimal amount) && amount > 0)
                    {
                        standardizedLines.Add($"{amount:N0} đ");
                    }
                    else
                    {
                        standardizedLines.Add(line.Trim());
                    }
                }
                else
                {
                    string amtStr = line.Substring(0, dashIndex);
                    string reasonStr = line.Substring(dashIndex + 1).Trim();
                    string cleanAmt = Regex.Replace(amtStr, @"[^0-9]", "");
                    if (decimal.TryParse(cleanAmt, out decimal amount))
                    {
                        standardizedLines.Add($"{amount:N0} đ - {reasonStr}");
                    }
                    else
                    {
                        standardizedLines.Add(line.Trim());
                    }
                }
            }
            return string.Join("\n", standardizedLines);
        }

        private decimal _m1Amount; public decimal M1Amount { get => _m1Amount; set { _m1Amount = value; OnPropertyChanged(); OnPropertyChanged("Total"); } }
        private string _m1Note = ""; public string M1Note { get => _m1Note; set { _m1Note = value; OnPropertyChanged(); if (IncomeType == "Thu nhập khác" || IncomeType == "Làm thêm giờ") { M1Amount = ParseAmountFromNote(value); OnPropertyChanged("M1Breakdown"); } } }
        public List<IncomeBreakdownItem> M1Breakdown => ParseBreakdown(M1Note);

        private decimal _m2Amount; public decimal M2Amount { get => _m2Amount; set { _m2Amount = value; OnPropertyChanged(); OnPropertyChanged("Total"); } }
        private string _m2Note = ""; public string M2Note { get => _m2Note; set { _m2Note = value; OnPropertyChanged(); if (IncomeType == "Thu nhập khác" || IncomeType == "Làm thêm giờ") { M2Amount = ParseAmountFromNote(value); OnPropertyChanged("M2Breakdown"); } } }
        public List<IncomeBreakdownItem> M2Breakdown => ParseBreakdown(M2Note);

        private decimal _m3Amount; public decimal M3Amount { get => _m3Amount; set { _m3Amount = value; OnPropertyChanged(); OnPropertyChanged("Total"); } }
        private string _m3Note = ""; public string M3Note { get => _m3Note; set { _m3Note = value; OnPropertyChanged(); if (IncomeType == "Thu nhập khác" || IncomeType == "Làm thêm giờ") { M3Amount = ParseAmountFromNote(value); OnPropertyChanged("M3Breakdown"); } } }
        public List<IncomeBreakdownItem> M3Breakdown => ParseBreakdown(M3Note);

        private decimal _m4Amount; public decimal M4Amount { get => _m4Amount; set { _m4Amount = value; OnPropertyChanged(); OnPropertyChanged("Total"); } }
        private string _m4Note = ""; public string M4Note { get => _m4Note; set { _m4Note = value; OnPropertyChanged(); if (IncomeType == "Thu nhập khác" || IncomeType == "Làm thêm giờ") { M4Amount = ParseAmountFromNote(value); OnPropertyChanged("M4Breakdown"); } } }
        public List<IncomeBreakdownItem> M4Breakdown => ParseBreakdown(M4Note);

        private decimal _m5Amount; public decimal M5Amount { get => _m5Amount; set { _m5Amount = value; OnPropertyChanged(); OnPropertyChanged("Total"); } }
        private string _m5Note = ""; public string M5Note { get => _m5Note; set { _m5Note = value; OnPropertyChanged(); if (IncomeType == "Thu nhập khác" || IncomeType == "Làm thêm giờ") { M5Amount = ParseAmountFromNote(value); OnPropertyChanged("M5Breakdown"); } } }
        public List<IncomeBreakdownItem> M5Breakdown => ParseBreakdown(M5Note);

        private decimal _m6Amount; public decimal M6Amount { get => _m6Amount; set { _m6Amount = value; OnPropertyChanged(); OnPropertyChanged("Total"); } }
        private string _m6Note = ""; public string M6Note { get => _m6Note; set { _m6Note = value; OnPropertyChanged(); if (IncomeType == "Thu nhập khác" || IncomeType == "Làm thêm giờ") { M6Amount = ParseAmountFromNote(value); OnPropertyChanged("M6Breakdown"); } } }
        public List<IncomeBreakdownItem> M6Breakdown => ParseBreakdown(M6Note);

        private decimal _m7Amount; public decimal M7Amount { get => _m7Amount; set { _m7Amount = value; OnPropertyChanged(); OnPropertyChanged("Total"); } }
        private string _m7Note = ""; public string M7Note { get => _m7Note; set { _m7Note = value; OnPropertyChanged(); if (IncomeType == "Thu nhập khác" || IncomeType == "Làm thêm giờ") { M7Amount = ParseAmountFromNote(value); OnPropertyChanged("M7Breakdown"); } } }
        public List<IncomeBreakdownItem> M7Breakdown => ParseBreakdown(M7Note);

        private decimal _m8Amount; public decimal M8Amount { get => _m8Amount; set { _m8Amount = value; OnPropertyChanged(); OnPropertyChanged("Total"); } }
        private string _m8Note = ""; public string M8Note { get => _m8Note; set { _m8Note = value; OnPropertyChanged(); if (IncomeType == "Thu nhập khác" || IncomeType == "Làm thêm giờ") { M8Amount = ParseAmountFromNote(value); OnPropertyChanged("M8Breakdown"); } } }
        public List<IncomeBreakdownItem> M8Breakdown => ParseBreakdown(M8Note);

        private decimal _m9Amount; public decimal M9Amount { get => _m9Amount; set { _m9Amount = value; OnPropertyChanged(); OnPropertyChanged("Total"); } }
        private string _m9Note = ""; public string M9Note { get => _m9Note; set { _m9Note = value; OnPropertyChanged(); if (IncomeType == "Thu nhập khác" || IncomeType == "Làm thêm giờ") { M9Amount = ParseAmountFromNote(value); OnPropertyChanged("M9Breakdown"); } } }
        public List<IncomeBreakdownItem> M9Breakdown => ParseBreakdown(M9Note);

        private decimal _m10Amount; public decimal M10Amount { get => _m10Amount; set { _m10Amount = value; OnPropertyChanged(); OnPropertyChanged("Total"); } }
        private string _m10Note = ""; public string M10Note { get => _m10Note; set { _m10Note = value; OnPropertyChanged(); if (IncomeType == "Thu nhập khác" || IncomeType == "Làm thêm giờ") { M10Amount = ParseAmountFromNote(value); OnPropertyChanged("M10Breakdown"); } } }
        public List<IncomeBreakdownItem> M10Breakdown => ParseBreakdown(M10Note);

        private decimal _m11Amount; public decimal M11Amount { get => _m11Amount; set { _m11Amount = value; OnPropertyChanged(); OnPropertyChanged("Total"); } }
        private string _m11Note = ""; public string M11Note { get => _m11Note; set { _m11Note = value; OnPropertyChanged(); if (IncomeType == "Thu nhập khác" || IncomeType == "Làm thêm giờ") { M11Amount = ParseAmountFromNote(value); OnPropertyChanged("M11Breakdown"); } } }
        public List<IncomeBreakdownItem> M11Breakdown => ParseBreakdown(M11Note);

        private decimal _m12Amount; public decimal M12Amount { get => _m12Amount; set { _m12Amount = value; OnPropertyChanged(); OnPropertyChanged("Total"); } }
        private string _m12Note = ""; public string M12Note { get => _m12Note; set { _m12Note = value; OnPropertyChanged(); if (IncomeType == "Thu nhập khác" || IncomeType == "Làm thêm giờ") { M12Amount = ParseAmountFromNote(value); OnPropertyChanged("M12Breakdown"); } } }
        public List<IncomeBreakdownItem> M12Breakdown => ParseBreakdown(M12Note);

        public decimal Total =>
            M1Amount + M2Amount + M3Amount + M4Amount + M5Amount + M6Amount +
            M7Amount + M8Amount + M9Amount + M10Amount + M11Amount + M12Amount;

        public void SetAmount(int month, decimal amount)
        {
            switch (month)
            {
                case 1: M1Amount = amount; break;
                case 2: M2Amount = amount; break;
                case 3: M3Amount = amount; break;
                case 4: M4Amount = amount; break;
                case 5: M5Amount = amount; break;
                case 6: M6Amount = amount; break;
                case 7: M7Amount = amount; break;
                case 8: M8Amount = amount; break;
                case 9: M9Amount = amount; break;
                case 10: M10Amount = amount; break;
                case 11: M11Amount = amount; break;
                case 12: M12Amount = amount; break;
            }
        }

        public decimal GetAmount(int month)
        {
            return month switch
            {
                1 => M1Amount,
                2 => M2Amount,
                3 => M3Amount,
                4 => M4Amount,
                5 => M5Amount,
                6 => M6Amount,
                7 => M7Amount,
                8 => M8Amount,
                9 => M9Amount,
                10 => M10Amount,
                11 => M11Amount,
                12 => M12Amount,
                _ => 0
            };
        }

        public void SetNote(int month, string note)
        {
            switch (month)
            {
                case 1: M1Note = note; break;
                case 2: M2Note = note; break;
                case 3: M3Note = note; break;
                case 4: M4Note = note; break;
                case 5: M5Note = note; break;
                case 6: M6Note = note; break;
                case 7: M7Note = note; break;
                case 8: M8Note = note; break;
                case 9: M9Note = note; break;
                case 10: M10Note = note; break;
                case 11: M11Note = note; break;
                case 12: M12Note = note; break;
            }
        }

        public string GetNote(int month)
        {
            return month switch
            {
                1 => M1Note,
                2 => M2Note,
                3 => M3Note,
                4 => M4Note,
                5 => M5Note,
                6 => M6Note,
                7 => M7Note,
                8 => M8Note,
                9 => M9Note,
                10 => M10Note,
                11 => M11Note,
                12 => M12Note,
                _ => ""
            };
        }


        public void TriggerMonthChanged(int month)
        {
            OnPropertyChanged($"M{month}Amount");
            OnPropertyChanged($"M{month}Note");
            OnPropertyChanged($"M{month}Breakdown");
            OnPropertyChanged("Total");
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string? name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
