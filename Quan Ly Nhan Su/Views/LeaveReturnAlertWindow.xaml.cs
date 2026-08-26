using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;

namespace TaxPersonnelManagement.Views
{
    public partial class LeaveReturnAlertWindow : Window
    {
        public LeaveReturnAlertWindow(List<LeaveReturnInfo> alerts)
        {
            InitializeComponent();

            int upcoming = alerts.Count(a => a.AlertType == "upcoming");
            int open     = alerts.Count(a => a.AlertType == "open");

            var parts = new List<string>();
            if (upcoming > 0) parts.Add($"{upcoming} sắp đi làm lại (7 ngày tới)");
            if (open     > 0) parts.Add($"{open} đang nghỉ chưa có ngày kết thúc");

            txtSubtitle.Text = string.Join("  •  ", parts);

            alerts = alerts
                .OrderBy(a => a.AlertType == "upcoming" ? 0 : 1)
                .ThenBy(a => a.ReturnDate ?? DateTime.MaxValue)
                .ToList();

            for (int i = 0; i < alerts.Count; i++)
                alerts[i].STT = i + 1;

            dgAlerts.ItemsSource = alerts;
        }

        private void Header_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
                DragMove();
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }

    public class LeaveReturnInfo
    {
        public int STT { get; set; }
        public string FullName { get; set; } = string.Empty;
        public string? Department { get; set; }
        public string LeaveType { get; set; } = string.Empty;
        public string AlertType { get; set; } = "upcoming";
        public DateTime? ReturnDate { get; set; }
        public int DaysLeft { get; set; }

        public string ReturnDateDisplay => AlertType == "open" ? "Chưa xác định" : ReturnDate?.ToString("dd/MM/yyyy") ?? "---";
        public string DepartmentDisplay => string.IsNullOrWhiteSpace(Department) ? "---" : Department;

        public string DaysLeftDisplay
        {
            get
            {
                if (AlertType == "open") return "Đang nghỉ";
                if (DaysLeft <= 0) return "Hôm nay";
                if (DaysLeft == 1) return "Ngày mai";
                return $"Còn {DaysLeft} ngày";
            }
        }

        public Brush DaysLeftBackground
        {
            get
            {
                if (AlertType == "open") return new SolidColorBrush(Color.FromRgb(241, 245, 249));
                if (DaysLeft <= 1) return new SolidColorBrush(Color.FromRgb(254, 226, 226));
                if (DaysLeft <= 3) return new SolidColorBrush(Color.FromRgb(254, 243, 199));
                return new SolidColorBrush(Color.FromRgb(224, 242, 254));
            }
        }

        public Brush DaysLeftForeground
        {
            get
            {
                if (AlertType == "open") return new SolidColorBrush(Color.FromRgb(100, 116, 139));
                if (DaysLeft <= 1) return new SolidColorBrush(Color.FromRgb(220, 38, 38));
                if (DaysLeft <= 3) return new SolidColorBrush(Color.FromRgb(217, 119, 6));
                return new SolidColorBrush(Color.FromRgb(2, 132, 199));
            }
        }

        public Brush DaysLeftBorder
        {
            get
            {
                if (AlertType == "open") return new SolidColorBrush(Color.FromRgb(203, 213, 225));
                if (DaysLeft <= 1) return new SolidColorBrush(Color.FromRgb(254, 202, 202));
                if (DaysLeft <= 3) return new SolidColorBrush(Color.FromRgb(253, 230, 138));
                return new SolidColorBrush(Color.FromRgb(186, 230, 253));
            }
        }

        public Brush LeaveTypeBackground
        {
            get
            {
                string lt = (LeaveType ?? "").ToLower();
                if (lt.Contains("thai sản")) return new SolidColorBrush(Color.FromRgb(253, 242, 248));
                if (lt.Contains("ốm")) return new SolidColorBrush(Color.FromRgb(255, 247, 237));
                if (lt.Contains("không lương") || lt.Contains("khong luong")) return new SolidColorBrush(Color.FromRgb(245, 243, 255));
                return new SolidColorBrush(Color.FromRgb(240, 253, 244));
            }
        }

        public Brush LeaveTypeForeground
        {
            get
            {
                string lt = (LeaveType ?? "").ToLower();
                if (lt.Contains("thai sản")) return new SolidColorBrush(Color.FromRgb(190, 24, 93));
                if (lt.Contains("ốm")) return new SolidColorBrush(Color.FromRgb(194, 65, 12));
                if (lt.Contains("không lương") || lt.Contains("khong luong")) return new SolidColorBrush(Color.FromRgb(109, 40, 217));
                return new SolidColorBrush(Color.FromRgb(21, 128, 61));
            }
        }
    }
}

