using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using Microsoft.Win32;
using ExcelDataReader;
using System.Data;
using TaxPersonnelManagement.Models;
using TaxPersonnelManagement.Data;

namespace TaxPersonnelManagement.Views
{
    public class ExcelSheetConfig : System.ComponentModel.INotifyPropertyChanged
    {
        public string SheetName { get; set; } = "";
        public string IncomeType { get; set; } = "";
        public int OriginalMonth { get; set; }
        public int OriginalYear { get; set; }
        public string OriginalMonthName => $"Tháng {OriginalMonth:00}";
        
        private bool _isSelected = true;
        public bool IsSelected
        {
            get => _isSelected;
            set { _isSelected = value; OnPropertyChanged(); }
        }

        private int _targetMonth;
        public int TargetMonth
        {
            get => _targetMonth;
            set
            {
                if (_targetMonth != value)
                {
                    _targetMonth = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(TargetMonthIndex));
                    UpdateDefaultNote();
                }
            }
        }

        public int TargetMonthIndex
        {
            get => TargetMonth - 1;
            set => TargetMonth = value + 1;
        }

        private string _note = "";
        public string Note
        {
            get => _note;
            set { _note = value; OnPropertyChanged(); }
        }

        public List<(string CCCD, decimal Amount)> Records { get; set; } = new List<(string CCCD, decimal Amount)>();
        public int RecordsCount => Records?.Count ?? 0;

        public string DefaultNotePrefix { get; set; } = "";

        public void UpdateDefaultNote()
        {
            if (!string.IsNullOrEmpty(DefaultNotePrefix))
            {
                Note = $"{DefaultNotePrefix} tháng {OriginalMonth:00}";
            }
            else if (IncomeType == "Lương")
            {
                Note = $"Lương tháng {OriginalMonth:00}";
            }
            else if (IncomeType == "Làm thêm giờ")
            {
                Note = $"Làm thêm giờ tháng {OriginalMonth:00}";
            }
            else
            {
                Note = $"Thu nhập khác tháng {OriginalMonth:00}";
            }
        }

        public event System.ComponentModel.PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([System.Runtime.CompilerServices.CallerMemberName] string? name = null)
        {
            PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(name));
        }
    }

    public partial class BulkIncomeDialog : Window
    {
        private ObservableCollection<ExcelSheetConfig> _sheetsList = new ObservableCollection<ExcelSheetConfig>();

        public BulkIncomeDialog(int defaultYear)
        {
            InitializeComponent();
            LoadYears(defaultYear);
            LoadDepartments();

            cboMonth.SelectedIndex = DateTime.Now.Month - 1;
            dgSheets.ItemsSource = _sheetsList;
        }

        private void LoadYears(int defaultYear)
        {
            for (int i = 2025; i <= DateTime.Now.Year + 5; i++)
            {
                cboYear.Items.Add(i);
            }
            cboYear.SelectedItem = defaultYear;
        }

        private void LoadDepartments()
        {
            try
            {
                using var db = new AppDbContext();
                var depts = db.Departments.Select(d => d.Name).OrderBy(n => n).ToList();
                cboDepartment.ItemsSource = depts;
                if (depts.Any())
                    cboDepartment.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                App.DebugLog("Error loading departments for BulkIncome: " + ex.Message);
            }
        }

        private void RadDepartment_Checked(object sender, RoutedEventArgs e)
        {
            if (cboDepartment != null)
                cboDepartment.IsEnabled = true;
        }

        private void RadDepartment_Unchecked(object sender, RoutedEventArgs e)
        {
            if (cboDepartment != null)
                cboDepartment.IsEnabled = false;
        }

        private void TxtAmount_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            // Chỉ cho phép nhập số (0-9)
            Regex regex = new Regex("[^0-9]+");
            e.Handled = regex.IsMatch(e.Text);
        }

        private void TxtAmount_TextChanged(object sender, TextChangedEventArgs e)
        {
            // Simple number formatting
            string value = Regex.Replace(txtAmount.Text, "[^0-9]", "");

            txtAmount.TextChanged -= TxtAmount_TextChanged;

            if (decimal.TryParse(value, out decimal amount))
            {
                int caretIndex = txtAmount.CaretIndex;
                int originalLength = txtAmount.Text.Length;

                txtAmount.Text = amount.ToString("N0");

                int newLength = txtAmount.Text.Length;
                int newCaretIndex = caretIndex + (newLength - originalLength);
                txtAmount.CaretIndex = Math.Max(0, newCaretIndex);
            }
            else
            {
                txtAmount.Text = "";
            }

            txtAmount.TextChanged += TxtAmount_TextChanged;
        }

        private void BtnApply_Click(object sender, RoutedEventArgs e)
        {
            // Validation
            if (cboMonth.SelectedIndex == -1 || cboYear.SelectedItem == null)
            {
                new WarningWindow("Vui lòng chọn Tháng và Năm.", "Lỗi nhập liệu") { Owner = this }.ShowDialog();
                return;
            }

            string reason = txtReason.Text.Trim();
            int selectedMonth = cboMonth.SelectedIndex + 1;
            int selectedYear = (int)cboYear.SelectedItem;

            // Handle Excel File target selection
            if (radExcelFile.IsChecked == true)
            {
                if (string.IsNullOrEmpty(_manualExcelFilePath) || !File.Exists(_manualExcelFilePath))
                {
                    new WarningWindow("Vui lòng chọn file Excel hợp lệ.", "Lỗi nhập liệu") { Owner = this }.ShowDialog();
                    return;
                }

                ImportFromManualExcelFile(_manualExcelFilePath, selectedMonth, selectedYear, reason);
                return;
            }

            // Standard manual input validation (for amount)
            string amountStr = txtAmount.Text.Replace(",", "").Replace(".", "");
            if (!decimal.TryParse(amountStr, out decimal amount) || amount <= 0)
            {
                new WarningWindow("Vui lòng nhập số tiền hợp lệ lớn hơn 0.", "Lỗi nhập liệu") { Owner = this }.ShowDialog();
                return;
            }

            if (string.IsNullOrEmpty(reason))
            {
                reason = "Thu nhập khác";
            }

            try
            {
                using var db = new AppDbContext();
                var personnelQuery = db.Personnel.AsQueryable();

                string targetDesc = "";

                if (radFemale.IsChecked == true)
                {
                    personnelQuery = personnelQuery.Where(p => p.Gender == "Nữ" || p.Gender == "nữ");
                    targetDesc = "Công chức Nữ";
                }
                else if (radMale.IsChecked == true)
                {
                    personnelQuery = personnelQuery.Where(p => p.Gender == "Nam" || p.Gender == "nam");
                    targetDesc = "Công chức Nam";
                }
                else if (radDepartment.IsChecked == true)
                {
                    string dept = cboDepartment.SelectedItem?.ToString() ?? "";
                    if (string.IsNullOrEmpty(dept))
                    {
                        new WarningWindow("Vui lòng chọn phòng ban.", "Lỗi nhập liệu") { Owner = this }.ShowDialog();
                        return;
                    }
                    personnelQuery = personnelQuery.Where(p => p.Department == dept);
                    targetDesc = $"Phòng ban: {dept}";
                }
                else
                {
                    targetDesc = "Tất cả công chức";
                }

                var targetPersonnel = personnelQuery.ToList();

                if (targetPersonnel.Count == 0)
                {
                    new WarningWindow("Không tìm thấy công chức nào thỏa mãn điều kiện đã chọn.", "Thông báo") { Owner = this }.ShowDialog();
                    return;
                }

                var confirm = new ConfirmDialog($"Bạn chuẩn bị thêm {amount:N0} đ vào [Thu nhập khác] cho {targetPersonnel.Count} người.\nĐối tượng: {targetDesc}\nTháng: {selectedMonth}/{selectedYear}\nNội dung: {reason}\n\nBạn có chắc chắn muốn thực hiện?");
                confirm.Owner = this;

                if (confirm.ShowDialog() == true)
                {
                    int updateCount = 0;

                    foreach (var person in targetPersonnel)
                    {
                        var record = db.IncomeRecords.FirstOrDefault(r =>
                            r.PersonnelId == person.Id &&
                            r.Year == selectedYear &&
                            r.Month == selectedMonth &&
                            r.IncomeType == "Thu nhập khác");

                        string newEntry = $"{amount:N0} đ - {reason}";

                        if (record != null)
                        {
                            string legacyNote = record.Note ?? "";
                            if (!string.IsNullOrWhiteSpace(legacyNote) && 
                                !Regex.IsMatch(legacyNote, @"^\s*\d+([\.,]\d+)*\s*đ?\s*-") && 
                                record.Amount > 0)
                            {
                                legacyNote = $"{record.Amount:N0} đ - {legacyNote}";
                            }

                            string combinedNote = string.IsNullOrEmpty(legacyNote) 
                                ? newEntry 
                                : legacyNote + "\n" + newEntry;

                            record.Note = AnnualIncomeRowViewModel.StandardizeOtherIncomeNote(combinedNote);
                            record.Amount = AnnualIncomeRowViewModel.ParseAmountFromNote(record.Note);
                        }
                        else
                        {
                            db.IncomeRecords.Add(new IncomeRecord
                            {
                                PersonnelId = person.Id,
                                Year = selectedYear,
                                Month = selectedMonth,
                                IncomeType = "Thu nhập khác",
                                Amount = amount,
                                Note = AnnualIncomeRowViewModel.StandardizeOtherIncomeNote(newEntry)
                            });
                        }
                        updateCount++;
                    }

                    db.SaveChanges();

                    var success = new SuccessWindow($"Đã nhập thu nhập thành công cho {updateCount} công chức!");
                    success.Owner = this;
                    success.ShowDialog();

                    this.DialogResult = true;
                }
            }
            catch (Exception ex)
            {
                new WarningWindow("Lỗi khi cập nhật dữ liệu: " + ex.Message, "Lỗi hệ thống") { Owner = this }.ShowDialog();
            }
        }

        private async void BtnBrowse_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xls;*.xlsx|All Files|*.*",
                Title = "Chọn file Bảng lương hoặc Làm thêm giờ Excel"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                string filePath = openFileDialog.FileName;
                txtFileName.Text = Path.GetFileName(filePath);
                txtFilePath.Text = filePath;

                try
                {
                    excelLoadingOverlay.Visibility = Visibility.Visible;
                    btnConfirmExcelImport.IsEnabled = false;
                    foreach (var cfg in _sheetsList)
                    {
                        cfg.PropertyChanged -= SheetConfig_PropertyChanged;
                    }
                    _sheetsList.Clear();

                    System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                    var detectedConfigs = new List<ExcelSheetConfig>();

                    await Task.Run(() =>
                    {
                        using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                        {
                            using (var reader = ExcelReaderFactory.CreateReader(stream))
                            {
                                var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
                                {
                                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                                    {
                                        UseHeaderRow = false
                                    }
                                });

                                if (dataSet.Tables.Count == 0) return;

                                foreach (DataTable table in dataSet.Tables)
                                {
                                    int detectedMonth = 0;
                                    int detectedYear = 0;
                                    string incomeType = "";
                                    int cccdIndex = -1;
                                    int amountIndex = -1;

                                    string defaultNotePrefix = "";

                                    // Find header row in first 10 rows
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
                                        else if (rowString.Contains("CHI THƯỞNG THEO NGHỊ ĐỊNH"))
                                        {
                                            incomeType = "Thu nhập khác";
                                            cccdIndex = 5; // Column F
                                            amountIndex = 30; // Column AE
                                            
                                            var ndMatch = Regex.Match(rowString, @"NGHỊ\s*ĐỊNH\s*(\d+(?:/\d+/[A-ZĐ\-]+)?)");
                                            if (ndMatch.Success)
                                            {
                                                defaultNotePrefix = $"Thưởng theo NĐ {ndMatch.Groups[1].Value}";
                                            }
                                            else
                                            {
                                                defaultNotePrefix = "Thưởng theo Nghị định";
                                            }
                                        }
                                        else if (rowString.Contains("CHI THƯỞNG HTNV CÔNG TÁC THUẾ"))
                                        {
                                            incomeType = "Thu nhập khác";
                                            cccdIndex = 5; // Column F
                                            amountIndex = 35; // Column AJ
                                            defaultNotePrefix = "Thưởng HTNV công tác thuế";
                                        }
                                        else if (rowString.Contains("TRUY LĨNH NÂNG LƯƠNG"))
                                        {
                                            incomeType = "Thu nhập khác";
                                            cccdIndex = 3; // Column D
                                            amountIndex = 31; // Column AF
                                            defaultNotePrefix = "Truy lĩnh nâng lương";
                                        }

                                        var match = Regex.Match(rowString, @"THÁNG\s*(\d{1,2})\s*/\s*(\d{4})");
                                        if (match.Success)
                                        {
                                            detectedMonth = int.Parse(match.Groups[1].Value);
                                            detectedYear = int.Parse(match.Groups[2].Value);
                                        }
                                        else
                                        {
                                            // Check for month and year in headers
                                            var monthMatch = Regex.Match(rowString, @"THÁNG\s*(\d{1,2})");
                                            if (monthMatch.Success && detectedMonth == 0)
                                            {
                                                detectedMonth = int.Parse(monthMatch.Groups[1].Value);
                                            }
                                            var yearMatch = Regex.Match(rowString, @"\b(202\d)\b");
                                            if (yearMatch.Success && detectedYear == 0)
                                            {
                                                detectedYear = int.Parse(yearMatch.Groups[1].Value);
                                            }
                                        }

                                        if (!string.IsNullOrEmpty(incomeType) && detectedMonth > 0 && detectedYear > 0)
                                        {
                                            break;
                                        }
                                    }

                                    // Fallback: Check sheet name if month isn't in headers
                                    if (detectedMonth == 0)
                                    {
                                        var sheetMatch = Regex.Match(table.TableName, @"^[tT](\d{1,2})$");
                                        if (sheetMatch.Success)
                                        {
                                            detectedMonth = int.Parse(sheetMatch.Groups[1].Value);
                                        }
                                    }

                                    // Final fallbacks for month and year
                                    if (detectedMonth == 0)
                                    {
                                        detectedMonth = DateTime.Now.Month;
                                    }
                                    if (detectedYear == 0)
                                    {
                                        detectedYear = DateTime.Now.Year;
                                    }

                                    if (!string.IsNullOrEmpty(incomeType))
                                    {
                                        var sheetRecords = new List<(string cccd, decimal amount)>();

                                        for (int r = 0; r < table.Rows.Count; r++)
                                        {
                                            var row = table.Rows[r];
                                            if (row.ItemArray.Length > Math.Max(cccdIndex, amountIndex))
                                            {
                                                string cccdStr = row[cccdIndex]?.ToString()?.Trim() ?? "";
                                                var amountObj = row[amountIndex];

                                                if (!string.IsNullOrEmpty(cccdStr) && Regex.IsMatch(cccdStr, @"^\d{9,15}$"))
                                                {
                                                    if (TryParseExcelValue(amountObj, out decimal amount) && amount > 0)
                                                    {
                                                        sheetRecords.Add((cccdStr, amount));
                                                    }
                                                }
                                            }
                                        }

                                        if (sheetRecords.Count > 0)
                                        {
                                            var config = new ExcelSheetConfig
                                            {
                                                SheetName = table.TableName,
                                                IncomeType = incomeType,
                                                OriginalMonth = detectedMonth,
                                                OriginalYear = detectedYear,
                                                TargetMonth = detectedMonth,
                                                DefaultNotePrefix = defaultNotePrefix,
                                                Records = sheetRecords
                                            };
                                            config.UpdateDefaultNote();
                                            detectedConfigs.Add(config);
                                        }
                                    }
                                }
                            }
                        }
                    });

                    excelLoadingOverlay.Visibility = Visibility.Collapsed;

                    if (detectedConfigs.Count == 0)
                    {
                        new WarningWindow("Không tìm thấy dữ liệu hợp lệ trong file Excel hoặc không nhận diện được cấu trúc cột.", "Lỗi Import") { Owner = this }.ShowDialog();
                        txtFileName.Text = "Chưa chọn file Excel";
                        txtFilePath.Text = "Chọn file Bảng lương hoặc Làm thêm giờ để phân tích...";
                        return;
                    }

                    if (detectedConfigs.Any(cfg => cfg.IncomeType == "Thu nhập khác"))
                    {
                        var confirm = new ConfirmDialog("File Excel này chứa dữ liệu Thưởng hoặc Truy lĩnh (Thu nhập khác).\nĐể nhập dữ liệu này chính xác, bạn nên dùng chức năng 'Theo file Excel' ở tab 'Nhập thủ công' để có thể tùy chỉnh Tháng, Năm và Nội dung ghi chú.\n\nBạn có muốn tự động chuyển sang tab 'Nhập thủ công' và tải file này không?");
                        confirm.Owner = this;
                        if (confirm.ShowDialog() == true)
                        {
                            tcBulkIncome.SelectedIndex = 0;
                            radExcelFile.IsChecked = true;
                            _manualExcelFilePath = filePath;
                            txtManualExcelPath.Text = Path.GetFileName(filePath);
                        }
                        
                        txtFileName.Text = "Chưa chọn file Excel";
                        txtFilePath.Text = "Chọn file Bảng lương hoặc Làm thêm giờ để phân tích...";
                        return;
                    }

                    foreach (var cfg in detectedConfigs)
                    {
                        cfg.PropertyChanged += SheetConfig_PropertyChanged;
                        _sheetsList.Add(cfg);
                    }
                    if (chkSelectAll != null)
                    {
                        chkSelectAll.IsChecked = true;
                    }
                    btnConfirmExcelImport.IsEnabled = true;
                }
                catch (Exception ex)
                {
                    excelLoadingOverlay.Visibility = Visibility.Collapsed;
                    new WarningWindow("Lỗi khi đọc file Excel: " + ex.Message, "Lỗi") { Owner = this }.ShowDialog();
                    txtFileName.Text = "Chưa chọn file Excel";
                    txtFilePath.Text = "Chọn file Bảng lương hoặc Làm thêm giờ để phân tích...";
                }
            }
        }

        private async void BtnConfirmExcelImport_Click(object sender, RoutedEventArgs e)
        {
            var selectedSheets = _sheetsList.Where(s => s.IsSelected).ToList();
            if (selectedSheets.Count == 0)
            {
                new WarningWindow("Vui lòng chọn ít nhất một sheet để import.", "Thông báo") { Owner = this }.ShowDialog();
                return;
            }

            string details = string.Join("\n", selectedSheets.Select(s => $"- Sheet {s.SheetName} ({s.IncomeType}) -> Áp dụng: Tháng {s.TargetMonth}/{s.OriginalYear} | Ghi chú: {s.Note}"));
            var confirm = new ConfirmDialog($"Bạn chuẩn bị import các sheet dữ liệu sau:\n{details}\n\nBạn có chắc chắn muốn thực hiện?");
            confirm.Owner = this;

            if (confirm.ShowDialog() == true)
            {
                try
                {
                    excelLoadingOverlay.Visibility = Visibility.Visible;
                    btnConfirmExcelImport.IsEnabled = false;

                    int totalSuccessCount = 0;
                    int totalNotFoundCount = 0;

                    await Task.Run(() =>
                    {
                        using (var db = new AppDbContext())
                        {
                            var clearedRecords = new HashSet<(int PersonnelId, int Year, int Month, string IncomeType)>();

                            foreach (var sheet in selectedSheets)
                            {
                                foreach (var data in sheet.Records)
                                {
                                    var person = db.Personnel.FirstOrDefault(p => p.IdentityCardNumber == data.CCCD);
                                    if (person != null)
                                    {
                                        var personKey = (person.Id, sheet.OriginalYear, sheet.TargetMonth, sheet.IncomeType);
                                        bool isFirstTimeInSession = !clearedRecords.Contains(personKey);
                                        if (isFirstTimeInSession)
                                        {
                                            clearedRecords.Add(personKey);
                                        }

                                        // Find existing record, either loaded locally or in DB
                                        var record = db.IncomeRecords.Local.FirstOrDefault(r =>
                                            r.PersonnelId == person.Id &&
                                            r.Year == sheet.OriginalYear &&
                                            r.Month == sheet.TargetMonth &&
                                            r.IncomeType == sheet.IncomeType);

                                        if (record == null)
                                        {
                                            record = db.IncomeRecords.FirstOrDefault(r =>
                                                r.PersonnelId == person.Id &&
                                                r.Year == sheet.OriginalYear &&
                                                r.Month == sheet.TargetMonth &&
                                                r.IncomeType == sheet.IncomeType);
                                        }

                                        if (record != null)
                                        {
                                            if (sheet.IncomeType == "Thu nhập khác" || sheet.IncomeType == "Làm thêm giờ")
                                            {
                                                string newEntry = $"{data.Amount:N0} đ - {sheet.Note}";
                                                string legacyNote = record.Note ?? "";
                                                if (!string.IsNullOrWhiteSpace(legacyNote) && 
                                                    !Regex.IsMatch(legacyNote, @"^\s*\d+([\.,]\d+)*\s*đ?\s*-") && 
                                                    record.Amount > 0)
                                                {
                                                    legacyNote = $"{record.Amount:N0} đ - {legacyNote}";
                                                }

                                                string combinedNote = string.IsNullOrEmpty(legacyNote) 
                                                    ? newEntry 
                                                    : legacyNote + "\n" + newEntry;

                                                record.Note = AnnualIncomeRowViewModel.StandardizeOtherIncomeNote(combinedNote);
                                                record.Amount = AnnualIncomeRowViewModel.ParseAmountFromNote(record.Note);
                                            }
                                            else
                                            {
                                                // Overwrite for "Lương"
                                                record.Amount = data.Amount;
                                                record.Note = sheet.Note;
                                            }
                                        }
                                        else
                                        {
                                            // Create new record
                                            string note = sheet.Note;
                                            decimal amount = data.Amount;

                                            if (sheet.IncomeType == "Thu nhập khác" || sheet.IncomeType == "Làm thêm giờ")
                                            {
                                                string newEntry = $"{data.Amount:N0} đ - {sheet.Note}";
                                                note = AnnualIncomeRowViewModel.StandardizeOtherIncomeNote(newEntry);
                                                amount = AnnualIncomeRowViewModel.ParseAmountFromNote(note);
                                            }

                                            db.IncomeRecords.Add(new IncomeRecord
                                            {
                                                PersonnelId = person.Id,
                                                Year = sheet.OriginalYear,
                                                Month = sheet.TargetMonth,
                                                IncomeType = sheet.IncomeType,
                                                Amount = amount,
                                                Note = note
                                            });
                                        }
                                        totalSuccessCount++;
                                    }
                                    else
                                    {
                                        totalNotFoundCount++;
                                    }
                                }
                            }

                            db.SaveChanges();
                        }
                    });

                    excelLoadingOverlay.Visibility = Visibility.Collapsed;

                    var success = new SuccessWindow($"Import thành công!", $"Đã cập nhật dữ liệu cho {totalSuccessCount} lượt công chức.\nKhông tìm thấy người nhận cho {totalNotFoundCount} lượt CCCD trong CSDL.");
                    success.Owner = this;
                    success.ShowDialog();

                    this.DialogResult = true;
                }
                catch (Exception ex)
                {
                    excelLoadingOverlay.Visibility = Visibility.Collapsed;
                    btnConfirmExcelImport.IsEnabled = true;
                    new WarningWindow("Lỗi khi import dữ liệu vào CSDL: " + ex.Message, "Lỗi") { Owner = this }.ShowDialog();
                }
            }
        }

        private bool _isUpdatingSelectAll = false;

        private void SheetConfig_PropertyChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(ExcelSheetConfig.IsSelected))
            {
                UpdateSelectAllHeaderState();
            }
        }

        private void UpdateSelectAllHeaderState()
        {
            if (_isUpdatingSelectAll || chkSelectAll == null) return;

            _isUpdatingSelectAll = true;
            
            bool allSelected = _sheetsList.All(s => s.IsSelected);
            bool allDeselected = _sheetsList.All(s => !s.IsSelected);

            if (allSelected)
            {
                chkSelectAll.IsChecked = true;
            }
            else if (allDeselected)
            {
                chkSelectAll.IsChecked = false;
            }
            else
            {
                chkSelectAll.IsChecked = null; // Indeterminate
            }

            _isUpdatingSelectAll = false;
        }

        private void ChkSelectAll_Checked(object sender, RoutedEventArgs e)
        {
            if (_isUpdatingSelectAll) return;
            _isUpdatingSelectAll = true;
            foreach (var sheet in _sheetsList)
            {
                sheet.IsSelected = true;
            }
            _isUpdatingSelectAll = false;
        }

        private void ChkSelectAll_Unchecked(object sender, RoutedEventArgs e)
        {
            if (_isUpdatingSelectAll) return;
            _isUpdatingSelectAll = true;
            foreach (var sheet in _sheetsList)
            {
                sheet.IsSelected = false;
            }
            _isUpdatingSelectAll = false;
        }

        private string _originalNote = "";
        private ExcelSheetConfig? _editingConfig = null;

        private Popup? _currentOpenPopup = null;

        private void TxtNote_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is TextBox tb && tb.DataContext is ExcelSheetConfig cfg)
            {
                if (!cfg.IsSelected) return; // Don't edit if deselected

                _editingConfig = cfg;
                _originalNote = cfg.Note;

                var grid = FindParent<Grid>(tb);
                if (grid != null)
                {
                    var popup = grid.Children.OfType<Popup>().FirstOrDefault();
                    if (popup != null)
                    {
                        // Close any other open popup first
                        if (_currentOpenPopup != null && _currentOpenPopup != popup)
                        {
                            _currentOpenPopup.IsOpen = false;
                        }

                        popup.IsOpen = true;
                        _currentOpenPopup = popup;

                        // Auto-focus the textbox inside the popup
                        var border = popup.Child as Border;
                        var stack = border?.Child as StackPanel;
                        var popupTb = stack?.Children.OfType<TextBox>().FirstOrDefault();
                        if (popupTb != null)
                        {
                            popupTb.Focus();
                            popupTb.CaretIndex = popupTb.Text.Length;
                        }
                    }
                }
                e.Handled = true; // Mark as handled to prevent further event routing
            }
        }

        protected override void OnPreviewMouseDown(MouseButtonEventArgs e)
        {
            base.OnPreviewMouseDown(e);

            if (_currentOpenPopup != null && _currentOpenPopup.IsOpen)
            {
                // Check if the mouse is over the popup or the target text box
                bool isMouseOverPopup = _currentOpenPopup.IsMouseOver;
                bool isMouseOverTarget = false;
                if (_currentOpenPopup.PlacementTarget is FrameworkElement target)
                {
                    isMouseOverTarget = target.IsMouseOver;
                }

                if (!isMouseOverPopup && !isMouseOverTarget)
                {
                    _currentOpenPopup.IsOpen = false;
                    _currentOpenPopup = null;
                }
            }
        }

        private void BtnPopupCancel_Click(object sender, RoutedEventArgs e)
        {
            if (_editingConfig != null)
            {
                _editingConfig.Note = _originalNote;
            }
            if (_currentOpenPopup != null)
            {
                _currentOpenPopup.IsOpen = false;
                _currentOpenPopup = null;
            }
        }

        private void BtnPopupSave_Click(object sender, RoutedEventArgs e)
        {
            if (_currentOpenPopup != null)
            {
                _currentOpenPopup.IsOpen = false;
                _currentOpenPopup = null;
            }
        }

        private string _manualExcelFilePath = "";

        private void RadExcelFile_Checked(object sender, RoutedEventArgs e)
        {
            if (gridExcelFile != null) gridExcelFile.IsEnabled = true;
            if (txtAmount != null)
            {
                txtAmount.Text = "";
                txtAmount.IsEnabled = false;
                txtAmount.Background = new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#F0F0F0"));
                txtAmount.Opacity = 0.6;
            }
        }

        private void RadExcelFile_Unchecked(object sender, RoutedEventArgs e)
        {
            if (gridExcelFile != null) gridExcelFile.IsEnabled = false;
            if (txtAmount != null)
            {
                txtAmount.IsEnabled = true;
                txtAmount.ClearValue(TextBox.BackgroundProperty);
                txtAmount.ClearValue(TextBox.OpacityProperty);
            }
        }

        private void BtnBrowseManualExcel_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xls;*.xlsx|All Files|*.*",
                Title = "Chọn file Excel Bảng thưởng/Truy lĩnh"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                _manualExcelFilePath = openFileDialog.FileName;
                txtManualExcelPath.Text = Path.GetFileName(_manualExcelFilePath);
            }
        }

        private async void ImportFromManualExcelFile(string filePath, int selectedMonth, int selectedYear, string reason)
        {
            try
            {
                excelLoadingOverlay.Visibility = Visibility.Visible;
                
                var detectedConfigs = new List<ExcelSheetConfig>();

                await Task.Run(() =>
                {
                    System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                    using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                    {
                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                                {
                                    UseHeaderRow = false
                                }
                            });

                            if (dataSet.Tables.Count == 0) return;

                            foreach (DataTable table in dataSet.Tables)
                            {
                                string incomeType = "";
                                int cccdIndex = -1;
                                int amountIndex = -1;
                                string defaultNotePrefix = "";

                                // Find header row in first 10 rows
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
                                    else if (rowString.Contains("CHI THƯỞNG THEO NGHỊ ĐỊNH"))
                                    {
                                        incomeType = "Thu nhập khác";
                                        cccdIndex = 5; // Column F
                                        amountIndex = 30; // Column AE
                                        
                                        var ndMatch = Regex.Match(rowString, @"NGHỊ\s*ĐỊNH\s*(\d+(?:/\d+/[A-ZĐ\-]+)?)");
                                        if (ndMatch.Success)
                                        {
                                            defaultNotePrefix = $"Thưởng theo NĐ {ndMatch.Groups[1].Value}";
                                        }
                                        else
                                        {
                                            defaultNotePrefix = "Thưởng theo Nghị định";
                                        }
                                    }
                                    else if (rowString.Contains("CHI THƯỞNG HTNV CÔNG TÁC THUẾ"))
                                    {
                                        incomeType = "Thu nhập khác";
                                        cccdIndex = 5; // Column F
                                        amountIndex = 35; // Column AJ
                                        defaultNotePrefix = "Thưởng HTNV công tác thuế";
                                    }
                                    else if (rowString.Contains("TRUY LĨNH NÂNG LƯƠNG"))
                                    {
                                        incomeType = "Thu nhập khác";
                                        cccdIndex = 3; // Column D
                                        amountIndex = 31; // Column AF
                                        defaultNotePrefix = "Truy lĩnh nâng lương";
                                    }

                                    if (!string.IsNullOrEmpty(incomeType))
                                    {
                                        break;
                                    }
                                }

                                if (!string.IsNullOrEmpty(incomeType))
                                {
                                    var sheetRecords = new List<(string cccd, decimal amount)>();

                                    for (int r = 0; r < table.Rows.Count; r++)
                                    {
                                        var row = table.Rows[r];
                                        if (row.ItemArray.Length > Math.Max(cccdIndex, amountIndex))
                                        {
                                            string cccdStr = row[cccdIndex]?.ToString()?.Trim() ?? "";
                                            var amountObj = row[amountIndex];

                                            if (!string.IsNullOrEmpty(cccdStr) && Regex.IsMatch(cccdStr, @"^\d{9,15}$"))
                                            {
                                                if (TryParseExcelValue(amountObj, out decimal amount) && amount > 0)
                                                {
                                                    sheetRecords.Add((cccdStr, amount));
                                                }
                                            }
                                        }
                                    }

                                    if (sheetRecords.Count > 0)
                                    {
                                        var config = new ExcelSheetConfig
                                        {
                                            SheetName = table.TableName,
                                            IncomeType = incomeType,
                                            OriginalMonth = selectedMonth,
                                            OriginalYear = selectedYear,
                                            TargetMonth = selectedMonth,
                                            DefaultNotePrefix = defaultNotePrefix,
                                            Records = sheetRecords
                                        };
                                        
                                        // If user typed a custom reason, override the note prefix
                                        if (!string.IsNullOrWhiteSpace(reason))
                                        {
                                            config.Note = reason;
                                        }
                                        else
                                        {
                                            config.UpdateDefaultNote();
                                        }

                                        detectedConfigs.Add(config);
                                    }
                                }
                            }
                        }
                    }
                });

                excelLoadingOverlay.Visibility = Visibility.Collapsed;

                if (detectedConfigs.Count == 0)
                {
                    new WarningWindow("Không tìm thấy dữ liệu hợp lệ trong file Excel hoặc không nhận diện được cấu trúc cột.", "Lỗi Import") { Owner = this }.ShowDialog();
                    return;
                }

                // Show confirmation before saving
                string details = string.Join("\n", detectedConfigs.Select(s => $"- Sheet {s.SheetName} ({s.IncomeType}) -> Áp dụng: Tháng {selectedMonth}/{selectedYear} | Ghi chú: {s.Note}"));
                var confirm = new ConfirmDialog($"Bạn chuẩn bị import dữ liệu từ file Excel sau:\n{details}\n\nBạn có chắc chắn muốn thực hiện?");
                confirm.Owner = this;

                if (confirm.ShowDialog() == true)
                {
                    excelLoadingOverlay.Visibility = Visibility.Visible;
                    
                    int totalSuccessCount = 0;
                    int totalNotFoundCount = 0;

                    await Task.Run(() =>
                    {
                        using (var db = new AppDbContext())
                        {
                            var clearedRecords = new HashSet<(int PersonnelId, int Year, int Month, string IncomeType)>();

                            foreach (var sheet in detectedConfigs)
                            {
                                foreach (var data in sheet.Records)
                                {
                                    var person = db.Personnel.FirstOrDefault(p => p.IdentityCardNumber == data.CCCD);
                                    if (person != null)
                                    {
                                        var personKey = (person.Id, selectedYear, selectedMonth, sheet.IncomeType);
                                        bool isFirstTimeInSession = !clearedRecords.Contains(personKey);
                                        if (isFirstTimeInSession)
                                        {
                                            clearedRecords.Add(personKey);
                                        }

                                        // Find existing record, either loaded locally or in DB
                                        var record = db.IncomeRecords.Local.FirstOrDefault(r =>
                                            r.PersonnelId == person.Id &&
                                            r.Year == selectedYear &&
                                            r.Month == selectedMonth &&
                                            r.IncomeType == sheet.IncomeType);

                                        if (record == null)
                                        {
                                            record = db.IncomeRecords.FirstOrDefault(r =>
                                                r.PersonnelId == person.Id &&
                                                r.Year == selectedYear &&
                                                r.Month == selectedMonth &&
                                                r.IncomeType == sheet.IncomeType);
                                        }

                                        if (record != null)
                                        {
                                            if (sheet.IncomeType == "Thu nhập khác" || sheet.IncomeType == "Làm thêm giờ")
                                            {
                                                string newEntry = $"{data.Amount:N0} đ - {sheet.Note}";
                                                string legacyNote = record.Note ?? "";
                                                if (!string.IsNullOrWhiteSpace(legacyNote) && 
                                                    !Regex.IsMatch(legacyNote, @"^\s*\d+([\.,]\d+)*\s*đ?\s*-") && 
                                                    record.Amount > 0)
                                                {
                                                    legacyNote = $"{record.Amount:N0} đ - {legacyNote}";
                                                }

                                                string combinedNote = string.IsNullOrEmpty(legacyNote) 
                                                    ? newEntry 
                                                    : legacyNote + "\n" + newEntry;

                                                record.Note = AnnualIncomeRowViewModel.StandardizeOtherIncomeNote(combinedNote);
                                                record.Amount = AnnualIncomeRowViewModel.ParseAmountFromNote(record.Note);
                                            }
                                            else
                                            {
                                                // Overwrite for "Lương"
                                                record.Amount = data.Amount;
                                                record.Note = sheet.Note;
                                            }
                                        }
                                        else
                                        {
                                            // Create new record
                                            string note = sheet.Note;
                                            decimal amount = data.Amount;

                                            if (sheet.IncomeType == "Thu nhập khác" || sheet.IncomeType == "Làm thêm giờ")
                                            {
                                                string newEntry = $"{data.Amount:N0} đ - {sheet.Note}";
                                                note = AnnualIncomeRowViewModel.StandardizeOtherIncomeNote(newEntry);
                                                amount = AnnualIncomeRowViewModel.ParseAmountFromNote(note);
                                            }

                                            db.IncomeRecords.Add(new IncomeRecord
                                            {
                                                PersonnelId = person.Id,
                                                Year = selectedYear,
                                                Month = selectedMonth,
                                                IncomeType = sheet.IncomeType,
                                                Amount = amount,
                                                Note = note
                                            });
                                        }
                                        totalSuccessCount++;
                                    }
                                    else
                                    {
                                        totalNotFoundCount++;
                                    }
                                }
                            }

                            db.SaveChanges();
                        }
                    });

                    excelLoadingOverlay.Visibility = Visibility.Collapsed;

                    var success = new SuccessWindow($"Import thành công!", $"Đã cập nhật dữ liệu cho {totalSuccessCount} lượt công chức.\nKhông tìm thấy người nhận cho {totalNotFoundCount} lượt CCCD trong CSDL.");
                    success.Owner = this;
                    success.ShowDialog();

                    this.DialogResult = true;
                }
            }
            catch (Exception ex)
            {
                excelLoadingOverlay.Visibility = Visibility.Collapsed;
                new WarningWindow("Lỗi khi import dữ liệu từ file Excel: " + ex.Message, "Lỗi") { Owner = this }.ShowDialog();
            }
        }

        private static T? FindParent<T>(DependencyObject child) where T : DependencyObject
        {
            DependencyObject parentObject = VisualTreeHelper.GetParent(child);
            if (parentObject == null) return null;
            if (parentObject is T parent) return parent;
            return FindParent<T>(parentObject);
        }

        private static bool TryParseExcelValue(object? val, out decimal amount)
        {
            amount = 0;
            if (val == null) return false;

            if (val is double d)
            {
                amount = (decimal)d;
                amount = Math.Round(amount, 0);
                return true;
            }
            if (val is float f)
            {
                amount = (decimal)f;
                amount = Math.Round(amount, 0);
                return true;
            }
            if (val is int i)
            {
                amount = i;
                return true;
            }
            if (val is long l)
            {
                amount = l;
                return true;
            }
            if (val is decimal dec)
            {
                amount = Math.Round(dec, 0);
                return true;
            }

            string amountStr = val.ToString()?.Trim() ?? "";
            return TryParseExcelAmount(amountStr, out amount);
        }

        private static bool TryParseExcelAmount(string amountStr, out decimal amount)
        {
            amount = 0;
            if (string.IsNullOrWhiteSpace(amountStr)) return false;

            // Strip any currency symbols (đ, vnđ, etc.) or spaces
            string clean = Regex.Replace(amountStr, @"[^\d\.,\-]", "").Trim();

            if (string.IsNullOrEmpty(clean)) return false;

            int dotCount = clean.Count(c => c == '.');
            int commaCount = clean.Count(c => c == ',');

            if (dotCount > 1)
            {
                clean = clean.Replace(".", "");
                dotCount = 0;
            }
            if (commaCount > 1)
            {
                clean = clean.Replace(",", "");
                commaCount = 0;
            }

            dotCount = clean.Count(c => c == '.');
            commaCount = clean.Count(c => c == ',');

            if (dotCount == 1 && commaCount == 1)
            {
                int dotIndex = clean.IndexOf('.');
                int commaIndex = clean.IndexOf(',');
                if (dotIndex < commaIndex)
                {
                    clean = clean.Replace(".", "").Replace(",", ".");
                }
                else
                {
                    clean = clean.Replace(",", "");
                }
            }
            else if (dotCount == 1 && commaCount == 0)
            {
                int dotIndex = clean.IndexOf('.');
                int charsAfterDot = clean.Length - 1 - dotIndex;
                if (charsAfterDot == 3)
                {
                    var viCulture = System.Globalization.CultureInfo.GetCultureInfo("vi-VN");
                    if (decimal.TryParse(clean, System.Globalization.NumberStyles.Any, viCulture, out amount))
                    {
                        amount = Math.Round(amount, 0);
                        return true;
                    }
                }
            }
            else if (commaCount == 1 && dotCount == 0)
            {
                int commaIndex = clean.IndexOf(',');
                int charsAfterComma = clean.Length - 1 - commaIndex;
                if (charsAfterComma != 3)
                {
                    clean = clean.Replace(",", ".");
                }
                else
                {
                    clean = clean.Replace(",", "");
                }
            }

            if (decimal.TryParse(clean, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out amount))
            {
                amount = Math.Round(amount, 0);
                return true;
            }

            return false;
        }
    }
}
