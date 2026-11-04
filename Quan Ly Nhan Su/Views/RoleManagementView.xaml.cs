using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Microsoft.EntityFrameworkCore;
using TaxPersonnelManagement.Data;
using TaxPersonnelManagement.Helpers;
using TaxPersonnelManagement.Models;
using TaxPersonnelManagement.Services;

namespace TaxPersonnelManagement.Views
{
    public partial class RoleManagementView : UserControl
    {
        private List<RoleItemViewModel> _allRoles = new();

        public RoleManagementView()
        {
            InitializeComponent();
            LoadData();
        }

        public void LoadData()
        {
            try
            {
                using var db = new AppDbContext();
                var roles = db.Roles.Include(r => r.Users).AsNoTracking().OrderBy(r => r.Id).ToList();

                int stt = 1;
                _allRoles = roles.Select(r =>
                {
                    var perms = PermissionService.DeserializePermissions(r.PermissionsJson);
                    bool isSuperAdmin = perms.Contains("*");

                    return new RoleItemViewModel
                    {
                        STT = stt++,
                        Id = r.Id,
                        Name = r.Name,
                        Description = string.IsNullOrWhiteSpace(r.Description) ? "—" : r.Description,
                        IsSystem = r.IsSystem,
                        UserCount = r.Users?.Count ?? 0,
                        PermissionsCount = perms.Count,
                        IsSuperAdmin = isSuperAdmin,
                        CreatedAt = r.CreatedAt
                    };
                }).ToList();

                // Cập nhật các thẻ KPI tổng quan
                txtKpiTotalRoles.Text = _allRoles.Count.ToString();
                txtKpiSystemRoles.Text = _allRoles.Count(r => r.IsSystem).ToString();
                txtKpiAssignedUsers.Text = _allRoles.Sum(r => r.UserCount).ToString();

                FilterData();
            }
            catch (Exception ex)
            {
                var warning = new WarningWindow("Lỗi nạp dữ liệu", "Không thể nạp danh sách vai trò từ CSDL:\n\n" + ex.Message);
                warning.Owner = Window.GetWindow(this);
                warning.ShowDialog();
            }
        }

        private void FilterData()
        {
            string keyword = txtSearch.Text.Trim();
            if (string.IsNullOrWhiteSpace(keyword))
            {
                dgRoles.ItemsSource = _allRoles;
            }
            else
            {
                dgRoles.ItemsSource = _allRoles.Where(r =>
                    SearchHelper.IsMatch(r.Name, keyword) ||
                    SearchHelper.IsMatch(r.Description, keyword)).ToList();
            }
        }

        private void TxtSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            FilterData();
        }

        private void BtnAddRole_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new RoleEditDialog();
            dialog.Owner = Window.GetWindow(this);
            dialog.ShowDialog();

            if (dialog.IsSuccess)
            {
                LoadData();
                var success = new SuccessWindow($"Đã tạo vai trò phân quyền '{dialog.SavedRoleName}' thành công!");
                success.Owner = Window.GetWindow(this);
                success.ShowDialog();
            }
        }

        private void BtnEditRole_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is int roleId)
            {
                var dialog = new RoleEditDialog(roleId);
                dialog.Owner = Window.GetWindow(this);
                dialog.ShowDialog();

                if (dialog.IsSuccess)
                {
                    LoadData();
                    var success = new SuccessWindow($"Đã cập nhật cấu hình vai trò '{dialog.SavedRoleName}' thành công!");
                    success.Owner = Window.GetWindow(this);
                    success.ShowDialog();
                }
            }
        }

        private void BtnDeleteRole_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is int roleId)
            {
                var role = _allRoles.FirstOrDefault(r => r.Id == roleId);
                if (role == null) return;

                if (role.IsSystem || role.Id == 1)
                {
                    var warning = new WarningWindow("Không thể xóa vai trò Quản trị viên hệ thống gốc.\n\nĐây là vai trò bảo đảm an toàn quyền truy cập cao nhất cho toàn bộ phần mềm.", "Từ chối thao tác");
                    warning.Owner = Window.GetWindow(this);
                    warning.ShowDialog();
                    return;
                }

                if (role.UserCount > 0)
                {
                    var warning = new WarningWindow($"Vai trò '{role.Name}' hiện đang được phân cho {role.UserCount} tài khoản người dùng.\n\nĐể đảm bảo an toàn dữ liệu, vui lòng chuyển các tài khoản này sang vai trò khác trong màn hình 'Tài khoản' trước khi xóa.", "Không thể xóa vai trò");
                    warning.Owner = Window.GetWindow(this);
                    warning.ShowDialog();
                    return;
                }

                var confirm = new ConfirmDialog(
                    $"Bạn có chắc chắn muốn xóa vai trò '{role.Name}' khỏi hệ thống?\n\nThao tác này sẽ xóa vĩnh viễn cấu hình quyền và không thể phục hồi.",
                    "Xác Nhận Xóa Vai Trò",
                    "XÓA VAI TRÒ"
                );
                confirm.Owner = Window.GetWindow(this);

                if (confirm.ShowDialog() == true)
                {
                    try
                    {
                        using var db = new AppDbContext();
                        var dbRole = db.Roles.Find(roleId);
                        if (dbRole != null)
                        {
                            db.Roles.Remove(dbRole);
                            db.SaveChanges();

                            PermissionService.InvalidateCache();
                            LoadData();

                            var success = new SuccessWindow($"Đã xóa vai trò '{role.Name}' thành công!");
                            success.Owner = Window.GetWindow(this);
                            success.ShowDialog();
                        }
                    }
                    catch (Exception ex)
                    {
                        var err = new WarningWindow("Quá trình xóa vai trò thất bại:\n\n" + ex.Message, "Lỗi khi xóa vai trò");
                        err.Owner = Window.GetWindow(this);
                        err.ShowDialog();
                    }
                }
            }
        }
    }

    public class RoleItemViewModel
    {
        public int STT { get; set; }
        public int Id { get; set; }
        public string Name { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public bool IsSystem { get; set; }
        public int UserCount { get; set; }
        public int PermissionsCount { get; set; }
        public bool IsSuperAdmin { get; set; }
        public DateTime CreatedAt { get; set; }

        public string UserCountText => $"{UserCount} tài khoản";
        public Visibility IsSystemBadgeVisibility => IsSystem ? Visibility.Visible : Visibility.Collapsed;
        public string CreatedAtFormatted => CreatedAt.ToString("dd/MM/yyyy");

        public string PermissionsSummary => IsSuperAdmin ? "Toàn quyền (SuperAdmin)" : $"{PermissionsCount} quyền hạn";
        public string PermissionBadgeBackground => IsSuperAdmin ? "#FEF3C7" : "#EFF6FF";
        public string PermissionBadgeForeground => IsSuperAdmin ? "#B45309" : "#1D4ED8";

        public bool CanDelete => !IsSystem && Id != 1;
        public string DeleteIconForeground => CanDelete ? "#EF4444" : "#94A3B8";
        public string DeleteButtonBackground => CanDelete ? "#FEF2F2" : "#F8FAFC";
    }
}
