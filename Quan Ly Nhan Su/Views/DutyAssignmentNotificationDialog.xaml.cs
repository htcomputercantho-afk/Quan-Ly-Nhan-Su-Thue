using System;
using System.Windows;
using System.Windows.Media;
using MaterialDesignThemes.Wpf;

namespace TaxPersonnelManagement.Views
{
    public enum DutyNotificationType
    {
        Success,
        Warning,
        Error,
        Confirm
    }

    /// <summary>
    /// Hộp thoại thông báo, cảnh báo và xác nhận hiện đại dành cho chức năng Phân công nhiệm vụ.
    /// </summary>
    public partial class DutyAssignmentNotificationDialog : Window
    {
        public bool Confirmed { get; private set; }

        public DutyAssignmentNotificationDialog(
            DutyNotificationType type,
            string message,
            string title,
            bool isDestructive = false,
            string confirmButtonText = "ĐỒNG Ý")
        {
            InitializeComponent();

            txtTitle.Text = title.ToUpperInvariant();
            txtMessage.Text = message;

            switch (type)
            {
                case DutyNotificationType.Success:
                    ApplySuccessTheme();
                    break;
                case DutyNotificationType.Warning:
                    ApplyWarningTheme();
                    break;
                case DutyNotificationType.Error:
                    ApplyErrorTheme();
                    break;
                case DutyNotificationType.Confirm:
                    ApplyConfirmTheme(isDestructive, confirmButtonText);
                    break;
            }
        }

        private void ApplySuccessTheme()
        {
            brdTopIndicator.Background = new SolidColorBrush(Color.FromRgb(16, 185, 129)); // #10B981
            brdIconContainer.Background = new SolidColorBrush(Color.FromRgb(236, 253, 245)); // #ECFDF5
            iconMain.Kind = PackIconKind.CheckCircle;
            iconMain.Foreground = new SolidColorBrush(Color.FromRgb(5, 150, 105)); // #059669
            txtTitle.Foreground = new SolidColorBrush(Color.FromRgb(6, 95, 70)); // #065F46

            btnSingleOk.Visibility = Visibility.Visible;
            btnSingleOk.Background = new SolidColorBrush(Color.FromRgb(5, 150, 105));
            txtSingleOk.Text = "ĐÃ HIỂU";
            pnlConfirmButtons.Visibility = Visibility.Collapsed;
        }

        private void ApplyWarningTheme()
        {
            brdTopIndicator.Background = new SolidColorBrush(Color.FromRgb(245, 158, 11)); // #F59E0B
            brdIconContainer.Background = new SolidColorBrush(Color.FromRgb(255, 251, 235)); // #FFFBEB
            iconMain.Kind = PackIconKind.AlertCircleOutline;
            iconMain.Foreground = new SolidColorBrush(Color.FromRgb(217, 119, 6)); // #D97706
            txtTitle.Foreground = new SolidColorBrush(Color.FromRgb(146, 64, 14)); // #92400E

            btnSingleOk.Visibility = Visibility.Visible;
            btnSingleOk.Background = new SolidColorBrush(Color.FromRgb(217, 119, 6));
            txtSingleOk.Text = "ĐÃ HIỂU";
            pnlConfirmButtons.Visibility = Visibility.Collapsed;
        }

        private void ApplyErrorTheme()
        {
            brdTopIndicator.Background = new SolidColorBrush(Color.FromRgb(239, 68, 68)); // #EF4444
            brdIconContainer.Background = new SolidColorBrush(Color.FromRgb(254, 242, 242)); // #FEF2F2
            iconMain.Kind = PackIconKind.CloseCircleOutline;
            iconMain.Foreground = new SolidColorBrush(Color.FromRgb(220, 38, 38)); // #DC2626
            txtTitle.Foreground = new SolidColorBrush(Color.FromRgb(153, 27, 27)); // #991B1B

            btnSingleOk.Visibility = Visibility.Visible;
            btnSingleOk.Background = new SolidColorBrush(Color.FromRgb(220, 38, 38));
            txtSingleOk.Text = "ĐÓNG";
            pnlConfirmButtons.Visibility = Visibility.Collapsed;
        }

        private void ApplyConfirmTheme(bool isDestructive, string confirmText)
        {
            btnSingleOk.Visibility = Visibility.Collapsed;
            pnlConfirmButtons.Visibility = Visibility.Visible;
            txtConfirmBtn.Text = confirmText;

            if (isDestructive)
            {
                brdTopIndicator.Background = new SolidColorBrush(Color.FromRgb(239, 68, 68));
                brdIconContainer.Background = new SolidColorBrush(Color.FromRgb(254, 242, 242));
                iconMain.Kind = PackIconKind.AlertCircleOutline;
                iconMain.Foreground = new SolidColorBrush(Color.FromRgb(220, 38, 38));
                txtTitle.Foreground = new SolidColorBrush(Color.FromRgb(153, 27, 27));

                btnConfirm.Background = new SolidColorBrush(Color.FromRgb(220, 38, 38));
                iconConfirmBtn.Kind = PackIconKind.DeleteOutline;
            }
            else
            {
                brdTopIndicator.Background = new SolidColorBrush(Color.FromRgb(37, 99, 235));
                brdIconContainer.Background = new SolidColorBrush(Color.FromRgb(239, 246, 255));
                iconMain.Kind = PackIconKind.HelpCircleOutline;
                iconMain.Foreground = new SolidColorBrush(Color.FromRgb(29, 78, 216));
                txtTitle.Foreground = new SolidColorBrush(Color.FromRgb(30, 58, 138));

                btnConfirm.Background = new SolidColorBrush(Color.FromRgb(29, 78, 216));
                iconConfirmBtn.Kind = PackIconKind.Check;
            }
        }

        private void BtnSingleOk_Click(object sender, RoutedEventArgs e)
        {
            Confirmed = true;
            DialogResult = true;
            Close();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            Confirmed = false;
            DialogResult = false;
            Close();
        }

        private void BtnConfirm_Click(object sender, RoutedEventArgs e)
        {
            Confirmed = true;
            DialogResult = true;
            Close();
        }

        // ============================================================
        // Helper Methods
        // ============================================================
        public static void ShowSuccess(Window? owner, string message, string title = "Thành công")
        {
            var dlg = new DutyAssignmentNotificationDialog(DutyNotificationType.Success, message, title);
            if (owner != null && owner.IsLoaded) dlg.Owner = owner;
            dlg.ShowDialog();
        }

        public static void ShowWarning(Window? owner, string message, string title = "Lưu ý")
        {
            var dlg = new DutyAssignmentNotificationDialog(DutyNotificationType.Warning, message, title);
            if (owner != null && owner.IsLoaded) dlg.Owner = owner;
            dlg.ShowDialog();
        }

        public static void ShowError(Window? owner, string message, string title = "Thông báo lỗi")
        {
            var dlg = new DutyAssignmentNotificationDialog(DutyNotificationType.Error, message, title);
            if (owner != null && owner.IsLoaded) dlg.Owner = owner;
            dlg.ShowDialog();
        }

        public static bool ShowConfirm(Window? owner, string message, string title = "Xác nhận", bool isDestructive = true, string confirmButtonText = "ĐỒNG Ý XÓA")
        {
            var dlg = new DutyAssignmentNotificationDialog(DutyNotificationType.Confirm, message, title, isDestructive, confirmButtonText);
            if (owner != null && owner.IsLoaded) dlg.Owner = owner;
            dlg.ShowDialog();
            return dlg.Confirmed;
        }
    }
}
