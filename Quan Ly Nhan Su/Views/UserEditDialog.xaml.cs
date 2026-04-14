using System;
using System.Linq;
using System.Windows;
using TaxPersonnelManagement.Data;
using TaxPersonnelManagement.Models;

namespace TaxPersonnelManagement.Views
{
    public partial class UserEditDialog : Window
    {
        public bool IsSuccess { get; private set; } = false;
        private int? _editingUserId = null;

        public UserEditDialog(int? userId = null)
        {
            InitializeComponent();
            _editingUserId = userId;
            
            // Populate Roles
            var roles = Enum.GetValues(typeof(UserRole)).Cast<UserRole>()
                            .Where(r => r == UserRole.Admin || r == UserRole.Staff)
                            .Select(r => new KeyValuePair<UserRole, string>(r, GetRoleDisplayName(r))).ToList();
            cboRole.ItemsSource = roles;

            if (_editingUserId.HasValue)
            {
                txtTitle.Text = "CẬP NHẬT TÀI KHOẢN";
                txtUsername.IsReadOnly = true; 
                pwdPassword.Visibility = Visibility.Collapsed;
                txtPasswordHint.Visibility = Visibility.Collapsed; // We will use a separate dialog for password
                txtPasswordLabel.Visibility = Visibility.Collapsed;
                LoadData();
            }
            else
            {
                txtTitle.Text = "THÊM MỚI TÀI KHOẢN";
                txtUsername.IsReadOnly = false;
                pwdPassword.Visibility = Visibility.Visible;
                txtPasswordHint.Visibility = Visibility.Collapsed;
                txtPasswordLabel.Visibility = Visibility.Visible;
                cboRole.SelectedValue = UserRole.Staff; // Default to Staff
            }
        }

        private string GetRoleDisplayName(UserRole role)
        {
            return role switch
            {
                UserRole.Admin => "Quản trị viên (Admin)",
                UserRole.Manager => "Quản lý (Manager)",
                UserRole.Staff => "Nhân viên (Staff)",
                _ => role.ToString()
            };
        }

        private void LoadData()
        {
            try
            {
                using (var db = new AppDbContext())
                {
                    var user = db.Users.Find(_editingUserId);
                    if (user != null)
                    {
                        txtUsername.Text = user.Username;
                        txtFullName.Text = user.FullName;
                        cboRole.SelectedValue = user.Role;
                    }
                }
            }
            catch (Exception ex)
            {
                var warning = new WarningWindow("Lỗi", "Không thể tải dữ liệu: " + ex.Message);
                warning.Owner = this;
                warning.ShowDialog();
            }
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtUsername.Text) || 
                string.IsNullOrWhiteSpace(txtFullName.Text) || 
                cboRole.SelectedValue == null)
            {
                var warning = new WarningWindow("Cảnh báo", "Vui lòng nhập đầy đủ Tên đăng nhập, Họ tên và chọn Quyền hạn.");
                warning.Owner = this;
                warning.ShowDialog();
                return;
            }

            try
            {
                using (var db = new AppDbContext())
                {
                    if (_editingUserId.HasValue)
                    {
                        var user = db.Users.Find(_editingUserId);
                        if (user != null)
                        {
                            user.FullName = txtFullName.Text;
                            user.Role = (UserRole)cboRole.SelectedValue;
                            db.SaveChanges();
                            IsSuccess = true;
                        }
                    }
                    else
                    {
                        if (string.IsNullOrWhiteSpace(pwdPassword.Password))
                        {
                            var warning = new WarningWindow("Cảnh báo", "Vui lòng nhập Mật khẩu cho tài khoản mới.");
                            warning.Owner = this;
                            warning.ShowDialog();
                            return;
                        }

                        // Check duplicate username
                        if (db.Users.Any(u => u.Username.ToLower() == txtUsername.Text.ToLower()))
                        {
                            var warning = new WarningWindow("Cảnh báo", "Tên đăng nhập đã tồn tại.");
                            warning.Owner = this;
                            warning.ShowDialog();
                            return;
                        }

                        var newUser = new User
                        {
                            Username = txtUsername.Text,
                            FullName = txtFullName.Text,
                            Role = (UserRole)cboRole.SelectedValue,
                            PasswordHash = pwdPassword.Password // Simplified for demo as per current system
                        };
                        db.Users.Add(newUser);
                        db.SaveChanges();
                        IsSuccess = true;
                    }
                }
                this.Close();
            }
            catch (Exception ex)
            {
                var warning = new WarningWindow("Lỗi", "Quá trình lưu thất bại: " + ex.Message);
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
