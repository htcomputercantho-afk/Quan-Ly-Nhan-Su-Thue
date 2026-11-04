using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using TaxPersonnelManagement.Data;
using TaxPersonnelManagement.Models;
using TaxPersonnelManagement.Services;

namespace TaxPersonnelManagement.Views
{
    public partial class RoleEditDialog : Window
    {
        public bool IsSuccess { get; private set; } = false;
        public string SavedRoleName { get; private set; } = string.Empty;
        private readonly int? _editingRoleId;
        private List<PermissionGroup> _groups = new();

        public RoleEditDialog(int? roleId = null)
        {
            InitializeComponent();
            _editingRoleId = roleId;

            // Nạp danh sách nhóm quyền động từ PermissionRegistry
            _groups = PermissionRegistry.GetFreshGroups();
            icPermissionGroups.ItemsSource = _groups;

            // Lắng nghe sự kiện thay đổi lựa chọn trên từng quyền để cập nhật số đếm trực quan
            foreach (var group in _groups)
            {
                foreach (var item in group.Permissions)
                {
                    item.SelectionChanged += (s, e) => UpdateSelectedCount();
                }
            }

            if (_editingRoleId.HasValue)
            {
                txtDialogTitle.Text = "CHỈNH SỬA VAI TRÒ & PHÂN QUYỀN HỆ THỐNG";
                LoadData();
            }
            else
            {
                txtDialogTitle.Text = "THÊM VAI TRÒ & PHÂN QUYỀN MỚI";
                UpdateSelectedCount();
            }
        }

        private void UpdateSelectedCount()
        {
            int total = _groups.Sum(g => g.Permissions.Count);
            int selected = _groups.Sum(g => g.Permissions.Count(p => p.IsSelected));
            if (txtSelectedCount != null)
            {
                txtSelectedCount.Text = $"{selected} / {total} quyền đã chọn";
            }
        }

        private void LoadData()
        {
            try
            {
                using var db = new AppDbContext();
                var role = db.Roles.Find(_editingRoleId);
                if (role == null) return;

                txtRoleName.Text = role.Name;
                txtDescription.Text = role.Description;

                if (role.IsSystem && role.Id == 1)
                {
                    txtRoleName.IsReadOnly = true;
                    txtRoleName.ToolTip = "Vai trò Quản trị viên hệ thống gốc không thể đổi tên";
                }

                var grantedCodes = PermissionService.DeserializePermissions(role.PermissionsJson);
                bool isSuperAdmin = grantedCodes.Contains("*");

                chkSuperAdmin.IsChecked = isSuperAdmin;

                foreach (var group in _groups)
                {
                    foreach (var item in group.Permissions)
                    {
                        // Khi là SuperAdmin thì danh sách bị khóa mờ, chỉ nạp quyền lẻ nếu không phải SuperAdmin
                        item.IsSelected = !isSuperAdmin && grantedCodes.Contains(item.Code, StringComparer.OrdinalIgnoreCase);
                    }
                }

                ApplySuperAdminState(isSuperAdmin);
            }
            catch (Exception ex)
            {
                var warning = new WarningWindow("Lỗi", "Không thể nạp thông tin vai trò: " + ex.Message);
                warning.Owner = this;
                warning.ShowDialog();
            }
        }

        private void ApplySuperAdminState(bool isSuperAdmin)
        {
            if (isSuperAdmin)
            {
                // 1. Ẩn 2 nút Chọn tất cả / Bỏ chọn tất cả
                if (pnlNormalSelectionButtons != null)
                {
                    pnlNormalSelectionButtons.Visibility = Visibility.Collapsed;
                }

                // 2. Khóa mờ danh sách 40 ô phân quyền chi tiết
                if (svPermissionsContainer != null)
                {
                    svPermissionsContainer.IsEnabled = false;
                    svPermissionsContainer.Opacity = 0.45;
                }

                // 3. Chuyển đổi nhãn trạng thái sang Toàn quyền tối cao
                if (borderSelectedCount != null)
                {
                    borderSelectedCount.Visibility = Visibility.Collapsed;
                }
                if (borderSuperAdminBadge != null)
                {
                    borderSuperAdminBadge.Visibility = Visibility.Visible;
                }
            }
            else
            {
                // 1. Hiện lại 2 nút Chọn tất cả / Bỏ chọn tất cả
                if (pnlNormalSelectionButtons != null)
                {
                    pnlNormalSelectionButtons.Visibility = Visibility.Visible;
                }

                // 2. Mở khóa và làm sáng danh sách phân quyền chi tiết
                if (svPermissionsContainer != null)
                {
                    svPermissionsContainer.IsEnabled = true;
                    svPermissionsContainer.Opacity = 1.0;
                }

                // 3. Chuyển lại badge đếm số quyền đã chọn
                if (borderSelectedCount != null)
                {
                    borderSelectedCount.Visibility = Visibility.Visible;
                }
                if (borderSuperAdminBadge != null)
                {
                    borderSuperAdminBadge.Visibility = Visibility.Collapsed;
                }
                UpdateSelectedCount();
            }
        }

        private void BtnSelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var group in _groups)
            {
                foreach (var item in group.Permissions)
                {
                    item.IsSelected = true;
                }
            }
            UpdateSelectedCount();
        }

        private void BtnUnselectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var group in _groups)
            {
                foreach (var item in group.Permissions)
                {
                    item.IsSelected = false;
                }
            }
            UpdateSelectedCount();
        }

        private void ChkSuperAdmin_Click(object sender, RoutedEventArgs e)
        {
            bool isChecked = chkSuperAdmin.IsChecked == true;
            ApplySuperAdminState(isChecked);
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            string roleName = txtRoleName.Text.Trim();
            if (string.IsNullOrWhiteSpace(roleName))
            {
                var warning = new WarningWindow("Vui lòng nhập Tên vai trò / Chức danh trước khi lưu cấu hình.", "Thiếu thông tin");
                warning.Owner = this;
                warning.ShowDialog();
                txtRoleName.Focus();
                return;
            }

            try
            {
                using var db = new AppDbContext();

                // Kiểm tra trùng tên vai trò
                bool isDuplicate = db.Roles.Any(r => r.Name.ToLower() == roleName.ToLower() && (!_editingRoleId.HasValue || r.Id != _editingRoleId.Value));
                if (isDuplicate)
                {
                    var warning = new WarningWindow($"Tên vai trò '{roleName}' đã tồn tại trong hệ thống.\n\nVui lòng đặt một tên khác để phân biệt.", "Trùng tên vai trò");
                    warning.Owner = this;
                    warning.ShowDialog();
                    txtRoleName.Focus();
                    return;
                }

                // Thu thập danh sách mã quyền được chọn
                List<string> selectedCodes = new();
                if (chkSuperAdmin.IsChecked == true)
                {
                    selectedCodes.Add("*");
                }
                else
                {
                    foreach (var group in _groups)
                    {
                        foreach (var item in group.Permissions)
                        {
                            if (item.IsSelected)
                            {
                                selectedCodes.Add(item.Code);
                            }
                        }
                    }
                }

                string permissionsJson = PermissionService.SerializePermissions(selectedCodes);

                if (_editingRoleId.HasValue)
                {
                    var existingRole = db.Roles.Find(_editingRoleId.Value);
                    if (existingRole != null)
                    {
                        if (!existingRole.IsSystem || existingRole.Id != 1)
                        {
                            existingRole.Name = roleName;
                        }
                        existingRole.Description = txtDescription.Text.Trim();
                        existingRole.PermissionsJson = permissionsJson;
                        db.SaveChanges();
                    }
                }
                else
                {
                    var newRole = new Role
                    {
                        Name = roleName,
                        Description = txtDescription.Text.Trim(),
                        PermissionsJson = permissionsJson,
                        IsSystem = false,
                        CreatedAt = DateTime.Now
                    };
                    db.Roles.Add(newRole);
                    db.SaveChanges();
                }

                // Làm mới cache phân quyền ngay lập tức
                PermissionService.InvalidateCache();

                SavedRoleName = roleName;
                IsSuccess = true;
                this.Close();
            }
            catch (Exception ex)
            {
                var warning = new WarningWindow("Quá trình lưu vai trò thất bại: " + ex.Message, "Lỗi lưu trữ");
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
