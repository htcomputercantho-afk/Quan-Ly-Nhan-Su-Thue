using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using TaxPersonnelManagement.Data;
using TaxPersonnelManagement.Models;

namespace TaxPersonnelManagement.Views
{
    public partial class BulkDecisionDialog : Window
    {
        public BulkDecisionDialog()
        {
            InitializeComponent();
            LoadYears();
        }

        private void LoadYears()
        {
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
                App.DebugLog("Error fetching evaluation years for bulk dialog: " + ex.Message);
            }

            int currentYear = DateTime.Now.Year;
            HashSet<int> allYears = new HashSet<int>();
            for (int y = 1980; y <= currentYear; y++)
            {
                allYears.Add(y);
            }
            foreach (var y in dbYears)
            {
                allYears.Add(y);
            }

            var sortedYears = allYears.OrderByDescending(x => x).ToList();
            cboYear.ItemsSource = sortedYears;
            cboYear.SelectedItem = currentYear;
        }

        private void Header_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                this.DragMove();
            }
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (cboYear.SelectedItem == null)
            {
                var warning = new WarningWindow("Vui lòng chọn năm áp dụng!", "Thông báo");
                warning.Owner = this;
                warning.ShowDialog();
                return;
            }

            int selectedYear = (int)cboYear.SelectedItem;
            int recordCount = 0;

            try
            {
                using (var db = new AppDbContext())
                {
                    recordCount = db.EvaluationRecords.Count(r => r.Year == selectedYear);
                }
            }
            catch (Exception ex)
            {
                var warning = new WarningWindow($"Lỗi khi kiểm tra dữ liệu: {ex.Message}", "Lỗi");
                warning.Owner = this;
                warning.ShowDialog();
                return;
            }

            if (recordCount == 0)
            {
                var warning = new WarningWindow($"Không tìm thấy nhân sự nào được xếp loại trong năm {selectedYear}!", "Thông báo");
                warning.Owner = this;
                warning.ShowDialog();
                return;
            }

            var confirm = new ConfirmWindow($"Bạn có chắc chắn muốn cập nhật thông tin quyết định cho {recordCount} nhân sự thuộc năm {selectedYear}?", "Xác nhận cập nhật");
            confirm.Owner = this;
            if (confirm.ShowDialog() == true)
            {
                string? decNo = string.IsNullOrWhiteSpace(txtDecisionNumber.Text) ? null : txtDecisionNumber.Text.Trim();
                DateTime? decDate = dpDecisionDate.SelectedDate;
                string? decAgency = string.IsNullOrWhiteSpace(txtDecisionAgency.Text) ? null : txtDecisionAgency.Text.Trim();

                try
                {
                    using (var db = new AppDbContext())
                    {
                        var records = db.EvaluationRecords.Where(r => r.Year == selectedYear).ToList();
                        foreach (var r in records)
                        {
                            r.DecisionNumber = decNo;
                            r.DecisionDate = decDate;
                            r.DecisionAgency = decAgency;
                        }
                        db.SaveChanges();
                    }

                    var success = new SuccessWindow($"Đã cập nhật thông tin quyết định thành công cho {recordCount} nhân sự năm {selectedYear}!", "Thành công");
                    success.Owner = this;
                    success.ShowDialog();

                    this.DialogResult = true;
                    this.Close();
                }
                catch (Exception ex)
                {
                    var warning = new WarningWindow($"Cập nhật thất bại: {ex.Message}", "Lỗi");
                    warning.Owner = this;
                    warning.ShowDialog();
                }
            }
        }
    }
}
