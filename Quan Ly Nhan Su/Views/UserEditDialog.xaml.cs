using System;
using System.Linq;
using System.Windows;
using TaxPersonnelManagement.Data;
using TaxPersonnelManagement.Models;
using TaxPersonnelManagement.Services;

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

            LoadRoles();

            if (_editingUserId.HasValue)
            {
                txtTitle.Text = "CẬP NHẬT TÀI KHOẢN";
                txtUsername.IsReadOnly = true;
                pwdPassword.Visibility = Visibility.Collapsed;
                txtPasswordHint.Visibility = Visibility.Collapsed;
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
            }
        }

        private void LoadRoles()
        {
            try
            {
                using var db = new AppDbContext();
                var roles = db.Roles.OrderBy(r => r.Id).ToList();
                cboRole.ItemsSource = roles;
                if (!_editingUserId.HasValue && roles.Count > 0)
                {
                    var defaultRole = roles.FirstOrDefault(r => !r.IsSystem) ?? roles[0];
                    cboRole.SelectedValue = defaultRole.Id;
                }
            }
            catch { }
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
                        cboRole.SelectedValue = user.RoleId ?? (user.Role == UserRole.Admin ? 1 : null);
                    }
                }
            }
            catch (Exception ex)
            {
                var warning = new WarningWindow("Lỗi", "Không thể nạp dữ liệu: " + ex.Message);
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
                var warning = new WarningWindow("Cảnh báo", "Vui lòng nhập đầy đủ Tên đăng nhập, Họ tên và chọn Vai trò.");
                warning.Owner = this;
                warning.ShowDialog();
                return;
            }

            int selectedRoleId = (int)cboRole.SelectedValue;

            try
            {
                using (var db = new AppDbContext())
                {
                    var selectedRole = db.Roles.Find(selectedRoleId);
                    UserRole mappedEnumRole = UserRole.Staff;
                    if (selectedRole != null && (selectedRole.IsSystem || selectedRole.Name.Contains("Admin", StringComparison.OrdinalIgnoreCase) || (selectedRole.PermissionsJson != null && selectedRole.PermissionsJson.Contains("*"))))
                    {
                        mappedEnumRole = UserRole.Admin;
                    }

                    if (_editingUserId.HasValue)
                    {
                        var user = db.Users.Find(_editingUserId);
                        if (user != null)
                        {
                            user.FullName = txtFullName.Text.Trim();
                            user.RoleId = selectedRoleId;
                            user.Role = mappedEnumRole;
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

                        if (db.Users.Any(u => u.Username.ToLower() == txtUsername.Text.Trim().ToLower()))
                        {
                            var warning = new WarningWindow("Cảnh báo", "Tên đăng nhập đã tồn tại trong hệ thống.");
                            warning.Owner = this;
                            warning.ShowDialog();
                            return;
                        }

                        var newUser = new User
                        {
                            Username = txtUsername.Text.Trim(),
                            FullName = txtFullName.Text.Trim(),
                            RoleId = selectedRoleId,
                            Role = mappedEnumRole,
                            PasswordHash = pwdPassword.Password
                        };
                        db.Users.Add(newUser);
                        db.SaveChanges();
                        IsSuccess = true;
                    }
                }
                PermissionService.InvalidateCache();
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
