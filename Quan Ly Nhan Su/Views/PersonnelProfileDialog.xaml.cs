using System;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using TaxPersonnelManagement.Models;

namespace TaxPersonnelManagement.Views
{
    public partial class PersonnelProfileDialog : Window
    {
        public PersonnelProfileDialog(Personnel personnel)
        {
            InitializeComponent();
            this.DataContext = personnel;

            CalculateStatistics(personnel);
        }

        private void CalculateStatistics(Personnel p)
        {
            // 1. Leave Logic (Simple Approximation for Demo)
            int used = 0;
            if (p.LeaveHistories != null)
            {
                var now = DateTime.Now;
                // Calculate used leave for the current year
                used = (int)p.LeaveHistories
                        .Where(h => h.StartDate.Year == now.Year)
                        .Sum(h => h.DurationDays);

            }
            txtUsedLeave.Text = used.ToString();

            int total = p.TotalAnnualLeaveDays;
            int remaining = total - used;
            if (remaining < 0) remaining = 0;
            txtRemainingLeave.Text = remaining.ToString();


            // Helper for precise duration
            string CalcDuration(DateTime start, DateTime end) 
            {
                if (start > end) return "0 năm 0 tháng 0 ngày";
                DateTime temp = start;
                int y = 0;
                while (temp.AddYears(1) <= end) { y++; temp = temp.AddYears(1); }
                int m = 0;
                while (temp.AddMonths(1) <= end) { m++; temp = temp.AddMonths(1); }
                int d = (end - temp).Days;
                return $"{y} năm {m} tháng {d} ngày";
            }
            DateTime today = DateTime.Now.Date;

            // 2. Work Years Logic (Tax Authority)
            if (p.TaxAuthorityStartDate.HasValue)
            {
                txtYearsWork.Text = CalcDuration(p.TaxAuthorityStartDate.Value, today);
            }

            // 3. Retirement Remaining Logic
            if (p.RetirementDate.HasValue)
            {
                if (p.RetirementDate.Value > today)
                    txtYearsRemaining.Text = CalcDuration(today, p.RetirementDate.Value);
                else
                    txtYearsRemaining.Text = "Đã nghỉ hưu";
            }

            // 4. Discipline Visibility
            if (!string.IsNullOrEmpty(p.DisciplineType) && p.DisciplineType != "-- Không có --" && p.DisciplineType != "---")
            {
                bdrNoDiscipline.Visibility = Visibility.Collapsed;
                bdrHasDiscipline.Visibility = Visibility.Visible;
            }
            else
            {
                bdrNoDiscipline.Visibility = Visibility.Visible;
                bdrHasDiscipline.Visibility = Visibility.Collapsed;
            }

            // 5. Leave History Visibility
             // 6. Calculated Position Stats
             if (p.PositionDecisionDate.HasValue)
             {
                 DateTime startDate = p.PositionDecisionDate.Value;
                 DateTime endDate = p.PositionCalculationDate ?? DateTime.Now.Date;
                 
                 if (endDate >= startDate)
                 {
                     int wYears = endDate.Year - startDate.Year;
                     if (startDate.Date > endDate.AddYears(-wYears)) wYears--;

                     int wMonths = 0;
                     DateTime tmpDate = startDate.AddYears(wYears);
                     while (tmpDate.AddMonths(1) <= endDate)
                     {
                         wMonths++;
                         tmpDate = tmpDate.AddMonths(1);
                     }
                     
                     txtCalculatedYears.Text = wYears.ToString();
                     txtCalculatedMonths.Text = wMonths.ToString();
                 }
             }

             if (p.LeaveHistories != null && p.LeaveHistories.Count > 0)
             {
                 txtNoLeaveHistory.Visibility = Visibility.Collapsed;
                 dgLeaveHistory.Visibility = Visibility.Visible;
             }
             else
             {
                 txtNoLeaveHistory.Visibility = Visibility.Visible;
                 dgLeaveHistory.Visibility = Visibility.Collapsed;
             }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Header_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                this.DragMove();
        }
        private void btnExportPdf_Click(object sender, RoutedEventArgs e)
        {
            var p = this.DataContext as Personnel;
            if (p == null) return;

            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
            dlg.FileName = $"HoSo_{p.StaffId}_{p.FullName}.pdf"; // Safe formatting needed but ok for now
            dlg.DefaultExt = ".pdf";
            dlg.Filter = "PDF Documents (.pdf)|*.pdf";

            if (dlg.ShowDialog() == true)
            {
                try
                {
                    TaxPersonnelManagement.Services.PdfExporter.Export(p, dlg.FileName);
                    
                    var success = new SuccessWindow("Xuất file PDF thành công!");
                    success.Owner = this;
                    success.ShowDialog();
                    
                    try 
                    {
                        var process = new System.Diagnostics.Process();
                        process.StartInfo = new System.Diagnostics.ProcessStartInfo(dlg.FileName) { UseShellExecute = true };
                        process.Start();
                    } catch {}
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Lỗi khi xuất file: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
    }
}
