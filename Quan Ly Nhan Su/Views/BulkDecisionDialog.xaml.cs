using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using Microsoft.EntityFrameworkCore;
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
            string selectedTarget = (cboTargetAudience.SelectedItem as System.Windows.Controls.ComboBoxItem)?.Content?.ToString() ?? "Toàn bộ công chức, người lao động";
            int recordCount = 0;

            try
            {
                using (var db = new AppDbContext())
                {
                    var query = db.EvaluationRecords.Include(r => r.Personnel).Where(r => r.Year == selectedYear);
                    if (selectedTarget != "Toàn bộ công chức, người lao động")
                    {
                        if (selectedTarget == "Trưởng Thuế cơ sở")
                        {
                            query = query.Where(r => r.Personnel != null && 
                                                    (r.Personnel.Position == "Trưởng Thuế cơ sở" || 
                                                     r.Personnel.Position == "Quyền Trưởng Thuế cơ sở"));
                        }
                        else
                        {
                            query = query.Where(r => r.Personnel != null && r.Personnel.Position == selectedTarget);
                        }
                    }
                    recordCount = query.Count();
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
                string targetMsg = selectedTarget == "Toàn bộ công chức, người lao động" ? "" : $" có chức vụ '{selectedTarget}'";
                var warning = new WarningWindow($"Không tìm thấy nhân sự nào{targetMsg} trong năm {selectedYear}!", "Thông báo");
                warning.Owner = this;
                warning.ShowDialog();
                return;
            }

            string? decNo = string.IsNullOrWhiteSpace(txtDecisionNumber.Text) ? null : txtDecisionNumber.Text.Trim();
            DateTime? decDate = dpDecisionDate.SelectedDate;
            string? decAgency = string.IsNullOrWhiteSpace(txtDecisionAgency.Text) ? null : txtDecisionAgency.Text.Trim();

            if (decNo == null && decDate == null && decAgency == null)
            {
                var warning = new WarningWindow("Vui lòng nhập ít nhất một thông tin quyết định (Số quyết định, Ngày ký, hoặc Đơn vị ra quyết định) để cập nhật!", "Thông báo");
                warning.Owner = this;
                warning.ShowDialog();
                return;
            }

            string targetDesc = selectedTarget == "Toàn bộ công chức, người lao động" ? "toàn bộ công chức, người lao động" : $"chức vụ '{selectedTarget}'";
            var confirm = new ConfirmWindow($"Bạn có chắc chắn muốn cập nhật thông tin quyết định cho {recordCount} nhân sự thuộc năm {selectedYear} có {targetDesc}?", "Xác nhận cập nhật");
            confirm.Owner = this;
            if (confirm.ShowDialog() == true)
            {
                try
                {
                    using (var db = new AppDbContext())
                    {
                        var query = db.EvaluationRecords.Include(r => r.Personnel).Where(r => r.Year == selectedYear);
                        if (selectedTarget != "Toàn bộ công chức, người lao động")
                        {
                            if (selectedTarget == "Trưởng Thuế cơ sở")
                            {
                                query = query.Where(r => r.Personnel != null && 
                                                        (r.Personnel.Position == "Trưởng Thuế cơ sở" || 
                                                         r.Personnel.Position == "Quyền Trưởng Thuế cơ sở"));
                            }
                            else
                            {
                                query = query.Where(r => r.Personnel != null && r.Personnel.Position == selectedTarget);
                            }
                        }
                        var records = query.ToList();
                        foreach (var r in records)
                        {
                            if (decNo != null)
                            {
                                r.DecisionNumber = decNo;
                            }
                            if (decDate.HasValue)
                            {
                                r.DecisionDate = decDate;
                            }
                            if (decAgency != null)
                            {
                                r.DecisionAgency = decAgency;
                            }
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
