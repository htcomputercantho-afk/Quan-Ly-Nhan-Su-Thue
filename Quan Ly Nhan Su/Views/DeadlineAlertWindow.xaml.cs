using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;

namespace TaxPersonnelManagement.Views
{
    public partial class DeadlineAlertWindow : Window
    {
        public DeadlineAlertWindow(List<DeadlineAlertInfo> alerts)
        {
            InitializeComponent();

            int salary      = alerts.Count(a => a.AlertCategory == "salary");
            int retirement  = alerts.Count(a => a.AlertCategory == "retirement");
            int appointment = alerts.Count(a => a.AlertCategory == "appointment");
            int allowance   = alerts.Count(a => a.AlertCategory == "allowance");

            var parts = new List<string>();
            if (allowance   > 0) parts.Add($"{allowance} hết hạn bảo lưu phụ cấp");
            if (salary      > 0) parts.Add($"{salary} đến hạn nâng lương");
            if (retirement  > 0) parts.Add($"{retirement} sắp nghỉ hưu");
            if (appointment > 0) parts.Add($"{appointment} hết nhiệm kỳ bổ nhiệm");
            txtSubtitle.Text = string.Join("  •  ", parts);

            alerts = alerts
                .OrderBy(a => a.AlertCategory == "allowance" ? 0 : a.AlertCategory == "retirement" ? 1 : a.AlertCategory == "appointment" ? 2 : 3)
                .ThenBy(a => a.DaysLeft)
                .ToList();

            for (int i = 0; i < alerts.Count; i++)
                alerts[i].STT = i + 1;

            dgDeadlines.ItemsSource = alerts;
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

    public class DeadlineAlertInfo
    {
        public int STT { get; set; }
        public string FullName { get; set; } = string.Empty;
        public string? Department { get; set; }
        public string AlertCategory { get; set; } = "salary";
        public DateTime DeadlineDate { get; set; }
        public int DaysLeft { get; set; }

        public string DepartmentDisplay => string.IsNullOrWhiteSpace(Department) ? "---" : Department;
        public string DeadlineDateDisplay => DeadlineDate.ToString("dd/MM/yyyy");

        public string AlertCategoryDisplay => AlertCategory switch
        {
            "allowance"   => "Hết hạn bảo lưu PC",
            "salary"      => "Nâng lương",
            "retirement"  => "Nghỉ hưu",
            "appointment" => "Bổ nhiệm",
            _             => AlertCategory
        };

        public string DaysLeftDisplay
        {
            get
            {
                if (DaysLeft == 0) return "Hôm nay";
                if (DaysLeft == 1) return "Ngày mai";
                return $"Còn {DaysLeft} ngày";
            }
        }

        public Brush AlertCategoryBackground => AlertCategory switch
        {
            "allowance"   => new SolidColorBrush(Color.FromRgb(243, 232, 255)),
            "salary"      => new SolidColorBrush(Color.FromRgb(254, 252, 232)),
            "retirement"  => new SolidColorBrush(Color.FromRgb(254, 226, 226)),
            "appointment" => new SolidColorBrush(Color.FromRgb(219, 234, 254)),
            _             => new SolidColorBrush(Color.FromRgb(241, 245, 249))
        };

        public Brush AlertCategoryForeground => AlertCategory switch
        {
            "allowance"   => new SolidColorBrush(Color.FromRgb(126, 34, 206)),
            "salary"      => new SolidColorBrush(Color.FromRgb(133, 77, 14)),
            "retirement"  => new SolidColorBrush(Color.FromRgb(185, 28, 28)),
            "appointment" => new SolidColorBrush(Color.FromRgb(29, 78, 216)),
            _             => new SolidColorBrush(Color.FromRgb(100, 116, 139))
        };

        public Brush DaysLeftBackground
        {
            get
            {
                if (DaysLeft <= 7)  return new SolidColorBrush(Color.FromRgb(254, 226, 226));
                if (DaysLeft <= 30) return new SolidColorBrush(Color.FromRgb(254, 243, 199));
                return new SolidColorBrush(Color.FromRgb(209, 250, 229));
            }
        }

        public Brush DaysLeftForeground
        {
            get
            {
                if (DaysLeft <= 7)  return new SolidColorBrush(Color.FromRgb(220, 38, 38));
                if (DaysLeft <= 30) return new SolidColorBrush(Color.FromRgb(217, 119, 6));
                return new SolidColorBrush(Color.FromRgb(5, 150, 105));
            }
        }

        public Brush DaysLeftBorder
        {
            get
            {
                if (DaysLeft <= 7)  return new SolidColorBrush(Color.FromRgb(254, 202, 202));
                if (DaysLeft <= 30) return new SolidColorBrush(Color.FromRgb(253, 230, 138));
                return new SolidColorBrush(Color.FromRgb(167, 243, 208));
            }
        }
    }
}
