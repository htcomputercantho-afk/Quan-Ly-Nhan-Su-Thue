using System;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using MaterialDesignThemes.Wpf;

namespace TaxPersonnelManagement.Views
{
    public partial class WarningWindow : Window
    {
        public WarningWindow(string param1, string param2 = "Thông báo")
        {
            InitializeComponent();

            string actualTitle = "Thông Báo";
            string actualMessage = "";

            string p1 = param1?.Trim() ?? "";
            string p2 = param2?.Trim() ?? "";

            // Xử lý trường hợp chỉ truyền 1 tham số (param2 là giá trị mặc định "Thông báo" hoặc rỗng)
            if (string.IsNullOrWhiteSpace(p2) || p2.Equals("Thông báo", StringComparison.OrdinalIgnoreCase))
            {
                actualMessage = p1;
                actualTitle = string.IsNullOrWhiteSpace(p2) ? "Thông Báo" : p2;
            }
            else
            {
                // Cả 2 tham số đều được truyền giá trị cụ thể.
                // Tự động phân định tham số nào là Title (tiêu đề ngắn), tham số nào là Message (nội dung chi tiết):
                if (p1.Contains('\n') && !p2.Contains('\n'))
                {
                    // p1 có xuống dòng -> 100% p1 là Message, p2 là Title
                    actualMessage = p1;
                    actualTitle = p2;
                }
                else if (p2.Contains('\n') && !p1.Contains('\n'))
                {
                    // p2 có xuống dòng -> 100% p2 là Message, p1 là Title
                    actualTitle = p1;
                    actualMessage = p2;
                }
                else if (p1.Length > p2.Length + 6)
                {
                    // p1 dài hơn p2 rõ rệt -> p1 là Message, p2 là Title
                    actualMessage = p1;
                    actualTitle = p2;
                }
                else if (p2.Length > p1.Length + 6)
                {
                    // p2 dài hơn p1 rõ rệt -> p1 là Title, p2 là Message
                    actualTitle = p1;
                    actualMessage = p2;
                }
                else
                {
                    // Độ dài xấp xỉ nhau: kiểm tra từ khóa tiêu đề phổ biến
                    string[] titleKeywords = new[] { 
                        "thông báo", "cảnh báo", "lỗi", "thiếu thông tin", "từ chối", 
                        "thất bại", "không thể", "trùng tên", "xác nhận", "nhập liệu", "error", "warning" 
                    };
                    bool p1IsKeyword = titleKeywords.Any(k => p1.ToLower().Contains(k));
                    bool p2IsKeyword = titleKeywords.Any(k => p2.ToLower().Contains(k));

                    if (p1IsKeyword && !p2IsKeyword)
                    {
                        actualTitle = p1;
                        actualMessage = p2;
                    }
                    else if (p2IsKeyword && !p1IsKeyword)
                    {
                        actualTitle = p2;
                        actualMessage = p1;
                    }
                    else
                    {
                        // Mặc định theo signature: param1 là message, param2 là title
                        actualMessage = p1;
                        actualTitle = p2;
                    }
                }
            }

            if (string.IsNullOrWhiteSpace(actualTitle))
            {
                actualTitle = "Thông Báo";
            }

            txtTitle.Text = actualTitle;
            txtMessage.Text = actualMessage;

            ApplyTheme(actualTitle, actualMessage);

            // Set parent window for center owner positioning
            if (this.Owner == null)
            {
                if (Application.Current?.MainWindow != null && Application.Current.MainWindow.IsVisible)
                {
                    this.Owner = Application.Current.MainWindow;
                    this.WindowStartupLocation = WindowStartupLocation.CenterOwner;
                }
                else
                {
                    this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
                }
            }

            // Keyboard navigation
            this.KeyDown += (s, e) =>
            {
                if (e.Key == Key.Escape || e.Key == Key.Enter)
                {
                    this.Close();
                }
            };

            // Allow dragging the window anywhere on card
            this.MouseDown += (s, e) =>
            {
                if (e.ChangedButton == MouseButton.Left)
                {
                    try
                    {
                        this.DragMove();
                    }
                    catch { }
                }
            };
        }

        private void ApplyTheme(string title, string message)
        {
            string combined = ((title ?? "") + " " + (message ?? "")).ToLower();

            // 1. Error / Denied Theme (Đỏ sang trọng)
            if (combined.Contains("lỗi") || combined.Contains("error") || combined.Contains("thất bại") || combined.Contains("từ chối") || combined.Contains("không thể xóa"))
            {
                SetColors(
                    outerBg: "#FEF2F2",
                    innerBg: "#FEE2E2",
                    iconColor: "#DC2626",
                    iconKind: PackIconKind.AlertCircleOutline,
                    accentColor: "#DC2626"
                );
            }
            // 2. Info / Blue Theme (Xanh dương cho thông báo chung, hướng dẫn)
            else if ((title?.ToLower().Contains("thông báo") == true || title?.ToLower().Contains("hướng dẫn") == true)
                     && !combined.Contains("thiếu") && !combined.Contains("chưa") && !combined.Contains("vui lòng") && !combined.Contains("trùng") && !combined.Contains("không có"))
            {
                SetColors(
                    outerBg: "#EFF6FF",
                    innerBg: "#DBEAFE",
                    iconColor: "#2563EB",
                    iconKind: PackIconKind.InformationOutline,
                    accentColor: "#2563EB"
                );
            }
            // 3. Warning / Prompt Theme (Vàng hổ phách Amber ấm áp cho thiếu thông tin, cảnh báo, nhắc nhở)
            else
            {
                SetColors(
                    outerBg: "#FFFBEB",
                    innerBg: "#FEF3C7",
                    iconColor: "#D97706",
                    iconKind: PackIconKind.AlertCircleOutline,
                    accentColor: "#D97706"
                );
            }
        }

        private void SetColors(string outerBg, string innerBg, string iconColor, PackIconKind iconKind, string accentColor)
        {
            try
            {
                var brushConverter = new BrushConverter();
                var colorConverter = (Color)ColorConverter.ConvertFromString(accentColor);

                if (brushConverter.ConvertFromString(outerBg) is Brush outer) borderIconOuter.Background = outer;
                if (brushConverter.ConvertFromString(innerBg) is Brush inner) borderIconInner.Background = inner;
                if (brushConverter.ConvertFromString(iconColor) is Brush iconBr) iconType.Foreground = iconBr;
                iconType.Kind = iconKind;

                if (brushConverter.ConvertFromString(accentColor) is Brush accentBrush)
                {
                    btnConfirm.Background = accentBrush;
                    btnConfirm.BorderBrush = accentBrush;
                }
                btnShadow.Color = colorConverter;
            }
            catch { }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
