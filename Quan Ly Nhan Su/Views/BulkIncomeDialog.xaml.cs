using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using TaxPersonnelManagement.Models;
using TaxPersonnelManagement.Data;
using System.Text.RegularExpressions;

namespace TaxPersonnelManagement.Views
{
    public partial class BulkIncomeDialog : Window
    {
        public BulkIncomeDialog(int defaultYear)
        {
            InitializeComponent();
            LoadYears(defaultYear);
            LoadDepartments();
            
            cboMonth.SelectedIndex = DateTime.Now.Month - 1;
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
            // Loại bỏ các ký tự không phải số (đề phòng trường hợp paste dữ liệu)
            string value = Regex.Replace(txtAmount.Text, "[^0-9]", "");
            
            // Tạm thời gỡ bỏ sự kiện TextChanged để tránh lặp vô hạn khi gán lại Text
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
                // Nếu xóa hết thì để trống
                txtAmount.Text = "";
            }

            // Gán lại sự kiện
            txtAmount.TextChanged += TxtAmount_TextChanged;
        }

        private void BtnApply_Click(object sender, RoutedEventArgs e)
        {
            // Validation
            if (cboMonth.SelectedIndex == -1 || cboYear.SelectedItem == null)
            {
                new WarningWindow("Vui lòng chọn Tháng và Năm.", "Lỗi nhập liệu"){ Owner = this }.ShowDialog();
                return;
            }

            string amountStr = txtAmount.Text.Replace(",", "").Replace(".", "");
            if (!decimal.TryParse(amountStr, out decimal amount) || amount <= 0)
            {
                new WarningWindow("Vui lòng nhập số tiền hợp lệ lớn hơn 0.", "Lỗi nhập liệu"){ Owner = this }.ShowDialog();
                return;
            }

            string reason = txtReason.Text.Trim();
            if (string.IsNullOrEmpty(reason))
            {
                reason = "Thu nhập khác";
            }

            int selectedMonth = cboMonth.SelectedIndex + 1;
            int selectedYear = (int)cboYear.SelectedItem;
            
            // Build query
            try
            {
                using var db = new AppDbContext();
                var personnelQuery = db.Personnel.AsQueryable();

                string targetDesc = "";

                if (radFemale.IsChecked == true)
                {
                    // Assuming "Nữ" is the standard value in DB
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
                        new WarningWindow("Vui lòng chọn phòng ban.", "Lỗi nhập liệu"){ Owner = this }.ShowDialog();
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
                    new WarningWindow("Không tìm thấy công chức nào thỏa mãn điều kiện đã chọn.", "Thông báo"){ Owner = this }.ShowDialog();
                    return;
                }

                // Confirm dialog
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

                        if (record != null)
                        {
                            // Append if exists
                            record.Amount += amount;
                            if (string.IsNullOrEmpty(record.Note))
                                record.Note = reason;
                            else if (!record.Note.Contains(reason))
                                record.Note += "; " + reason;
                        }
                        else
                        {
                            // Create new
                            db.IncomeRecords.Add(new IncomeRecord
                            {
                                PersonnelId = person.Id,
                                Year = selectedYear,
                                Month = selectedMonth,
                                IncomeType = "Thu nhập khác",
                                Amount = amount,
                                Note = reason
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
                new WarningWindow("Lỗi khi cập nhật dữ liệu: " + ex.Message, "Lỗi hệ thống"){ Owner = this }.ShowDialog();
            }
        }
    }
}
