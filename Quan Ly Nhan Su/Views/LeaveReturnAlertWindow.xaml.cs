using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;

namespace TaxPersonnelManagement.Views
{
    /// <summary>
    /// Hiển thị danh sách công chức sắp kết thúc nghỉ (thai sản / ốm / phép)
    /// và sắp đi làm lại trong vòng 7 ngày tới với giao diện hiện đại, trực quan.
    /// </summary>
    public partial class LeaveReturnAlertWindow : Window
    {
        public LeaveReturnAlertWindow(List<LeaveReturnInfo> alerts)
        {
            InitializeComponent();

            int count = alerts.Count;
            txtSubtitle.Text = $"Có {count} công chức sẽ hoàn thành đợt nghỉ và đi làm lại trong 7 ngày tới";

            // Gán STT
            for (int i = 0; i < alerts.Count; i++)
                alerts[i].STT = i + 1;

            dgAlerts.ItemsSource = alerts;
        }

        private void Header_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                this.DragMove();
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }

    /// <summary>
    /// DTO lưu thông tin một công chức sắp đi làm lại.
    /// </summary>
    public class LeaveReturnInfo
    {
        public int STT { get; set; }
        public string FullName { get; set; } = string.Empty;
        public string? Department { get; set; }
        public string LeaveType { get; set; } = string.Empty;

        /// <summary>Ngày đi làm lại = EndDate + 1</summary>
        public DateTime ReturnDate { get; set; }

        /// <summary>Số ngày còn lại cho đến ngày đi làm lại (tính từ hôm nay).</summary>
        public int DaysLeft { get; set; }

        // ── Computed display properties ─────────────────────────────────────

        public string ReturnDateDisplay => ReturnDate.ToString("dd/MM/yyyy");

        public string DepartmentDisplay => string.IsNullOrWhiteSpace(Department) ? "---" : Department;

        public string DaysLeftDisplay
        {
            get
            {
                if (DaysLeft <= 0) return "Hôm nay";
                if (DaysLeft == 1) return "Ngày mai";
                return $"Còn {DaysLeft} ngày";
            }
        }

        /// <summary>Màu badge "Còn lại": đỏ khi <= 1 ngày, vàng hổ phách khi <= 3 ngày, xanh dương nhẹ khi còn nhiều hơn.</summary>
        public Brush DaysLeftBackground
        {
            get
            {
                if (DaysLeft <= 1) return new SolidColorBrush(Color.FromRgb(254, 226, 226)); // Red 100
                if (DaysLeft <= 3) return new SolidColorBrush(Color.FromRgb(254, 243, 199)); // Amber 100
                return new SolidColorBrush(Color.FromRgb(224, 242, 254));                     // Sky 100
            }
        }

        public Brush DaysLeftForeground
        {
            get
            {
                if (DaysLeft <= 1) return new SolidColorBrush(Color.FromRgb(220, 38, 38));  // Red 600
                if (DaysLeft <= 3) return new SolidColorBrush(Color.FromRgb(217, 119, 6));  // Amber 600
                return new SolidColorBrush(Color.FromRgb(2, 132, 199));                      // Sky 600
            }
        }

        public Brush DaysLeftBorder
        {
            get
            {
                if (DaysLeft <= 1) return new SolidColorBrush(Color.FromRgb(254, 202, 202)); // Red 200
                if (DaysLeft <= 3) return new SolidColorBrush(Color.FromRgb(253, 230, 138)); // Amber 200
                return new SolidColorBrush(Color.FromRgb(186, 230, 253));                     // Sky 200
            }
        }

        /// <summary>Màu badge loại nghỉ</summary>
        public Brush LeaveTypeBackground
        {
            get
            {
                string lt = (LeaveType ?? "").ToLower();
                if (lt.Contains("thai sản")) return new SolidColorBrush(Color.FromRgb(253, 242, 248)); // Pink 50
                if (lt.Contains("ốm")) return new SolidColorBrush(Color.FromRgb(255, 247, 237));       // Orange 50
                return new SolidColorBrush(Color.FromRgb(240, 253, 244));                               // Green 50
            }
        }

        public Brush LeaveTypeForeground
        {
            get
            {
                string lt = (LeaveType ?? "").ToLower();
                if (lt.Contains("thai sản")) return new SolidColorBrush(Color.FromRgb(190, 24, 93));  // Pink 700
                if (lt.Contains("ốm")) return new SolidColorBrush(Color.FromRgb(194, 65, 12));        // Orange 700
                return new SolidColorBrush(Color.FromRgb(21, 128, 61));                                // Green 700
            }
        }
    }
}

