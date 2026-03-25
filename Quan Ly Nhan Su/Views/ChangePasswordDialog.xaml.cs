using System;
using System.Linq;
using System.Windows;
using TaxPersonnelManagement.Data;

namespace TaxPersonnelManagement.Views
{
    public partial class ChangePasswordDialog : Window
    {
        public bool IsSuccess { get; private set; } = false;
        private int _userId;

        public ChangePasswordDialog(int userId, string username)
        {
            InitializeComponent();
            _userId = userId;
            txtUsernameDisplay.Text = username;
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(pwdNewPassword.Password))
            {
                var warning = new WarningWindow("Cảnh báo", "Vui lòng nhập mật khẩu mới.");
                warning.Owner = this;
                warning.ShowDialog();
                return;
            }

            if (pwdNewPassword.Password != pwdConfirmPassword.Password)
            {
                var warning = new WarningWindow("Cảnh báo", "Mật khẩu xác nhận không khớp.");
                warning.Owner = this;
                warning.ShowDialog();
                return;
            }

            try
            {
                using (var db = new AppDbContext())
                {
                    var user = db.Users.Find(_userId);
                    if (user != null)
                    {
                        user.PasswordHash = pwdNewPassword.Password;
                        db.SaveChanges();
                        IsSuccess = true;
                    }
                }
                this.Close();
            }
            catch (Exception ex)
            {
                var warning = new WarningWindow("Lỗi", "Đổi mật khẩu thất bại: " + ex.Message);
                warning.Owner = this;
                warning.ShowDialog();
            }
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
