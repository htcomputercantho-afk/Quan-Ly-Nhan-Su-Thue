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
                btnImportExcel.Visibility = Visibility.Collapsed;
                btnBulkAdd.Visibility = Visibility.Collapsed;
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
                _allPersonnel = db.Personnel.OrderBy(p => p.FullName).ToList();
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
            string keyword = txtSearch.Text.ToLower();
            lvPersonnel.ItemsSource = _allPersonnel.Where(p => 
                (p.FullName != null && p.FullName.ToLower().Contains(keyword)) ||
                (p.StaffId != null && p.StaffId.ToLower().Contains(keyword))).ToList();
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
                            targetRow.SetAmount(record.Month, record.Amount);
                            targetRow.SetNote(record.Month, record.Note ?? "");
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
            if (e.Row.Item is AnnualIncomeRowViewModel row && row.IsTotalRow)
            {
                e.Cancel = true; // Prevent editing the total row
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
                                db.IncomeRecords.Add(new IncomeRecord { PersonnelId = selectedPerson.Id, Year = selectedYear, Month = m, IncomeType = "Làm thêm giờ", Amount = overtimeRow.GetAmount(m), Note = overtimeRow.GetNote(m) });

                            if (otherRow.GetAmount(m) > 0 || !string.IsNullOrWhiteSpace(otherRow.GetNote(m)))
                                db.IncomeRecords.Add(new IncomeRecord { PersonnelId = selectedPerson.Id, Year = selectedYear, Month = m, IncomeType = "Thu nhập khác", Amount = otherRow.GetAmount(m), Note = otherRow.GetNote(m) });
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
        
        private void BtnImportExcel_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xls;*.xlsx|All Files|*.*",
                Title = "Chọn file Bảng lương Excel"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    // Configure text encoding for ExcelDataReader (required for .NET Core)
                    System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                    int detectedMonth = 0;
                    int detectedYear = 0;
                    string incomeType = "";
                    int cccdIndex = -1;
                    int amountIndex = -1;
                    var importedData = new List<(string cccd, decimal amount)>();

                    using (var stream = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.Read))
                    {
                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                                {
                                    UseHeaderRow = false // We read raw rows because headers are spread out
                                }
                            });

                            if (dataSet.Tables.Count == 0) return;
                            
                            // We need to find the correct table in case it's not the first one (e.g. TH vs T01)
                            DataTable? targetTable = null;
                            
                            foreach (DataTable table in dataSet.Tables)
                            {
                                // 1. Find the header row containing our key identifying strings
                                for (int i = 0; i < Math.Min(10, table.Rows.Count); i++)
                                {
                                    string rowString = string.Join(" ", table.Rows[i].ItemArray).ToUpper();
                                    
                                    if (rowString.Contains("BẢNG THANH TOÁN TIỀN LƯƠNG"))
                                    {
                                        incomeType = "Lương";
                                        cccdIndex = 4; // Column E
                                        amountIndex = 55; // Column BD
                                    }
                                    else if (rowString.Contains("BẢNG THANH TOÁN TIỀN LÀM THÊM GIỜ") || rowString.Contains("TIỀN CHI LÀM THÊM GIỜ"))
                                    {
                                        incomeType = "Làm thêm giờ";
                                        cccdIndex = 2;   // Column C
                                        amountIndex = 28; // Column AC
                                    }

                                    // Try to match "THÁNG 01/2026" or similar anywhere in the top 10 rows
                                    var match = Regex.Match(rowString, @"THÁNG\s*(\d{1,2})\s*/\s*(\d{4})");
                                    if (match.Success)
                                    {
                                        detectedMonth = int.Parse(match.Groups[1].Value);
                                        detectedYear = int.Parse(match.Groups[2].Value);
                                    }

                                    if (!string.IsNullOrEmpty(incomeType) && detectedMonth > 0 && detectedYear > 0)
                                    {
                                        targetTable = table;
                                        break;
                                    }
                                }
                                if (targetTable != null) break;
                            }

                            if (targetTable == null || detectedMonth == 0 || detectedYear == 0)
                            {
                                var warning = new WarningWindow("Không thể tự động nhận diện loại file hoặc Tháng/Năm từ tiêu đề. Vui lòng kiểm tra lại định dạng file.", "Lỗi Import");
                                warning.ShowDialog();
                                return;
                            }

                            // Confirm with user
                            var confirmDialog = new ConfirmDialog($"Phần mềm nhận diện bảng thu nhập này là [{incomeType}] thuộc Tháng {detectedMonth} năm {detectedYear}.\n\nBạn có muốn Import dữ liệu vào tháng này không?");
                            if (confirmDialog.ShowDialog() != true) return;

                            // 2. Iterate through rows to find data
                            for (int r = 0; r < targetTable.Rows.Count; r++)
                            {
                                var row = targetTable.Rows[r];
                                // Ensure row has enough columns
                                if (row.ItemArray.Length > Math.Max(cccdIndex, amountIndex))
                                {
                                    string cccdStr = row[cccdIndex]?.ToString()?.Trim() ?? "";
                                    string amountStr = row[amountIndex]?.ToString()?.Trim() ?? ""; 

                                    // Simple heuristic: if CCCD column is a long string of digits, it's a CCCD
                                    if (!string.IsNullOrEmpty(cccdStr) && Regex.IsMatch(cccdStr, @"^\d{9,12}$"))
                                    {
                                        if (decimal.TryParse(amountStr, out decimal amount) && amount > 0)
                                        {
                                            importedData.Add((cccdStr, amount));
                                        }
                                    }
                                }
                            }
                        }
                    }

                    if (importedData.Count == 0)
                    {
                        var warning = new WarningWindow("Không tìm thấy dữ liệu hợp lệ trong file Excel. Vui lòng kiểm tra cấu trúc cột.", "Lỗi Import");
                        warning.ShowDialog();
                        return;
                    }

                    // 3. Save to database
                    using (var db = new AppDbContext())
                    {
                        int successCount = 0;
                        int notFoundCount = 0;

                        foreach (var data in importedData)
                        {
                            var person = db.Personnel.FirstOrDefault(p => p.IdentityCardNumber == data.cccd);
                            if (person != null)
                            {
                                // Remove existing record for this month/year/type
                                var existing = db.IncomeRecords.FirstOrDefault(r => 
                                    r.PersonnelId == person.Id && 
                                    r.Year == detectedYear && 
                                    r.Month == detectedMonth &&
                                    r.IncomeType == incomeType);

                                if (existing != null)
                                {
                                    db.IncomeRecords.Remove(existing);
                                }

                                // Add new record
                                db.IncomeRecords.Add(new IncomeRecord
                                {
                                    PersonnelId = person.Id,
                                    Year = detectedYear,
                                    Month = detectedMonth,
                                    IncomeType = incomeType,
                                    Amount = data.amount,
                                    Note = incomeType == "Lương" ? "Import từ file Bảng lương" : "Import từ file Làm thêm giờ"
                                });
                                successCount++;
                            }
                            else
                            {
                                notFoundCount++;
                            }
                        }

                        db.SaveChanges();

                        // Change selected year in combo box to imported year if different, to show results instantly
                        if ((int)(cboYear.SelectedItem ?? 0) != detectedYear)
                        {
                            LoadYears();
                            cboYear.SelectedItem = detectedYear;
                        }
                        
                        // Force refresh grid if a user is selected
                        if (lvPersonnel.SelectedItem != null)
                        {
                            var temp = lvPersonnel.SelectedItem;
                            lvPersonnel.SelectedItem = null;
                            lvPersonnel.SelectedItem = temp;
                        }

                        var successDialog = new SuccessWindow($"Import thành công!\n\n- Đã cập nhật lương cho {successCount} công chức.\n- Không tìm thấy người nhận cho {notFoundCount} CCCD trong CSDL.", openFileDialog.FileName);
                        successDialog.ShowDialog();
                    }
                }
                catch (Exception ex)
                {
                    var errorDialog = new WarningWindow("Lỗi khi đọc file Excel: " + ex.Message, "Lỗi");
                    errorDialog.ShowDialog();
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
    }

    public class AnnualIncomeRowViewModel : INotifyPropertyChanged
    {
        public string IncomeType { get; set; } = "";
        public bool IsTotalRow { get; set; } = false;

        private decimal _m1Amount; public decimal M1Amount { get => _m1Amount; set { _m1Amount = value; OnPropertyChanged(); OnPropertyChanged("Total"); } }
        private string _m1Note = ""; public string M1Note { get => _m1Note; set { _m1Note = value; OnPropertyChanged(); } }

        private decimal _m2Amount; public decimal M2Amount { get => _m2Amount; set { _m2Amount = value; OnPropertyChanged(); OnPropertyChanged("Total"); } }
        private string _m2Note = ""; public string M2Note { get => _m2Note; set { _m2Note = value; OnPropertyChanged(); } }

        private decimal _m3Amount; public decimal M3Amount { get => _m3Amount; set { _m3Amount = value; OnPropertyChanged(); OnPropertyChanged("Total"); } }
        private string _m3Note = ""; public string M3Note { get => _m3Note; set { _m3Note = value; OnPropertyChanged(); } }

        private decimal _m4Amount; public decimal M4Amount { get => _m4Amount; set { _m4Amount = value; OnPropertyChanged(); OnPropertyChanged("Total"); } }
        private string _m4Note = ""; public string M4Note { get => _m4Note; set { _m4Note = value; OnPropertyChanged(); } }

        private decimal _m5Amount; public decimal M5Amount { get => _m5Amount; set { _m5Amount = value; OnPropertyChanged(); OnPropertyChanged("Total"); } }
        private string _m5Note = ""; public string M5Note { get => _m5Note; set { _m5Note = value; OnPropertyChanged(); } }

        private decimal _m6Amount; public decimal M6Amount { get => _m6Amount; set { _m6Amount = value; OnPropertyChanged(); OnPropertyChanged("Total"); } }
        private string _m6Note = ""; public string M6Note { get => _m6Note; set { _m6Note = value; OnPropertyChanged(); } }

        private decimal _m7Amount; public decimal M7Amount { get => _m7Amount; set { _m7Amount = value; OnPropertyChanged(); OnPropertyChanged("Total"); } }
        private string _m7Note = ""; public string M7Note { get => _m7Note; set { _m7Note = value; OnPropertyChanged(); } }

        private decimal _m8Amount; public decimal M8Amount { get => _m8Amount; set { _m8Amount = value; OnPropertyChanged(); OnPropertyChanged("Total"); } }
        private string _m8Note = ""; public string M8Note { get => _m8Note; set { _m8Note = value; OnPropertyChanged(); } }

        private decimal _m9Amount; public decimal M9Amount { get => _m9Amount; set { _m9Amount = value; OnPropertyChanged(); OnPropertyChanged("Total"); } }
        private string _m9Note = ""; public string M9Note { get => _m9Note; set { _m9Note = value; OnPropertyChanged(); } }

        private decimal _m10Amount; public decimal M10Amount { get => _m10Amount; set { _m10Amount = value; OnPropertyChanged(); OnPropertyChanged("Total"); } }
        private string _m10Note = ""; public string M10Note { get => _m10Note; set { _m10Note = value; OnPropertyChanged(); } }

        private decimal _m11Amount; public decimal M11Amount { get => _m11Amount; set { _m11Amount = value; OnPropertyChanged(); OnPropertyChanged("Total"); } }
        private string _m11Note = ""; public string M11Note { get => _m11Note; set { _m11Note = value; OnPropertyChanged(); } }

        private decimal _m12Amount; public decimal M12Amount { get => _m12Amount; set { _m12Amount = value; OnPropertyChanged(); OnPropertyChanged("Total"); } }
        private string _m12Note = ""; public string M12Note { get => _m12Note; set { _m12Note = value; OnPropertyChanged(); } }

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


        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string? name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
