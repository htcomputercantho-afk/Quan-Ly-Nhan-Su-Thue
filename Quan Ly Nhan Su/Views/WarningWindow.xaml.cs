using System;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using MaterialDesignThemes.Wpf;

namespace TaxPersonnelManagement.Views
{
    public partial class WarningWindow : Window
    {
        public WarningWindow(string message, string title = "Thông báo")
        {
            InitializeComponent();

            // Handle legacy reversed arguments if title was passed first
            if ((message == "Lỗi" || message == "Cảnh báo" || message == "Cảnh Báo" || message == "Thông báo" || message == "Từ chối truy cập")
                && !string.IsNullOrEmpty(title) && title.Length > message.Length)
            {
                string tmp = message;
                message = title;
                title = tmp;
            }

            txtMessage.Text = message;
            txtTitle.Text = string.IsNullOrWhiteSpace(title) ? "Thông Báo" : title;

            ApplyTheme(title, message);

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
            string lowerTitle = title?.ToLower() ?? "";
            string lowerMsg = message?.ToLower() ?? "";

            // 1. Error theme (Red)
            if (lowerTitle.Contains("lỗi") || lowerTitle.Contains("error") || lowerTitle.Contains("thất bại"))
            {
                SetColors(
                    outerBg: "#FEF2F2",
                    innerBg: "#FEE2E2",
                    iconColor: "#DC2626",
                    iconKind: PackIconKind.AlertCircle,
                    accentColor: "#DC2626"
                );
            }
            // 2. Info / Blue theme (General notifications without warning intent)
            else if (lowerTitle.Contains("thông báo") && !lowerMsg.Contains("không có") && !lowerMsg.Contains("chưa") && !lowerMsg.Contains("vui lòng") && !lowerMsg.Contains("lỗi"))
            {
                SetColors(
                    outerBg: "#EFF6FF",
                    innerBg: "#DBEAFE",
                    iconColor: "#0284C7",
                    iconKind: PackIconKind.InformationOutline,
                    accentColor: "#0284C7"
                );
            }
            // 3. Warning theme (Warm Amber / Gold) - default for "Không có dữ liệu", warnings, prompts
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
