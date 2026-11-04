using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Windows;
using Microsoft.EntityFrameworkCore;
using TaxPersonnelManagement.Data;
using TaxPersonnelManagement.Models;
using TaxPersonnelManagement.Views;

namespace TaxPersonnelManagement.Services
{
    /// <summary>
    /// Dịch vụ trung tâm xử lý và kiểm tra phân quyền (RBAC) cho toàn bộ ứng dụng.
    /// Hỗ trợ bộ nhớ đệm (Cache) tốc độ cao để đảm bảo hiệu năng chuyển menu 0ms.
    /// </summary>
    public static class PermissionService
    {
        private static HashSet<string>? _cachedCurrentUserPermissions = null;
        private static int? _cachedUserId = null;

        /// <summary>
        /// Xóa bộ nhớ cache quyền khi người dùng đăng nhập lại hoặc khi quyền của vai trò được cập nhật.
        /// </summary>
        public static void InvalidateCache()
        {
            _cachedCurrentUserPermissions = null;
            _cachedUserId = null;
        }

        /// <summary>
        /// Nạp lại danh sách quyền của người dùng hiện tại từ CSDL.
        /// </summary>
        public static void ReloadCurrentUserPermissions()
        {
            InvalidateCache();
            if (App.CurrentUser == null) return;

            try
            {
                using var db = new AppDbContext();
                // Nạp thông tin Role mới nhất của user
                var userWithRole = db.Users.Include(u => u.UserRoleEntity).FirstOrDefault(u => u.Id == App.CurrentUser.Id);
                if (userWithRole != null)
                {
                    App.CurrentUser.UserRoleEntity = userWithRole.UserRoleEntity;
                    App.CurrentUser.RoleId = userWithRole.RoleId;
                    App.CurrentUser.Role = userWithRole.Role;
                }
            }
            catch { }

            EnsureCacheLoaded();
        }

        private static void EnsureCacheLoaded()
        {
            if (App.CurrentUser == null)
            {
                _cachedCurrentUserPermissions = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                _cachedUserId = null;
                return;
            }

            if (_cachedCurrentUserPermissions != null && _cachedUserId == App.CurrentUser.Id)
            {
                return;
            }

            _cachedUserId = App.CurrentUser.Id;
            var perms = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // 1. Quản trị viên hệ thống (Admin) luôn có toàn quyền tuyệt đối
            if (App.CurrentUser.Role == UserRole.Admin)
            {
                perms.Add("*");
                _cachedCurrentUserPermissions = perms;
                return;
            }

            // 2. Nạp từ RoleEntity nếu có
            string? permJson = App.CurrentUser.UserRoleEntity?.PermissionsJson;

            // Nếu UserRoleEntity chưa nạp nhưng có RoleId, truy vấn CSDL
            if (string.IsNullOrEmpty(permJson) && App.CurrentUser.RoleId.HasValue)
            {
                try
                {
                    using var db = new AppDbContext();
                    var role = db.Roles.Find(App.CurrentUser.RoleId.Value);
                    if (role != null)
                    {
                        App.CurrentUser.UserRoleEntity = role;
                        permJson = role.PermissionsJson;
                    }
                }
                catch { }
            }

            if (!string.IsNullOrWhiteSpace(permJson))
            {
                try
                {
                    var list = JsonSerializer.Deserialize<List<string>>(permJson);
                    if (list != null)
                    {
                        foreach (var code in list)
                        {
                            perms.Add(code);
                        }
                    }
                }
                catch { }
            }

            // 3. Fallback cho tài khoản cũ chưa gán vai trò RBAC mới
            if (perms.Count == 0)
            {
                if (App.CurrentUser.Role == UserRole.Manager)
                {
                    // Lãnh đạo: Xem được tất cả các menu nghiệp vụ & xuất Excel
                    perms.Add(PermissionRegistry.Menu_Dashboard);
                    perms.Add(PermissionRegistry.Menu_Statistics);
                    perms.Add(PermissionRegistry.Statistics_Export);
                    perms.Add(PermissionRegistry.Menu_Personnel);
                    perms.Add(PermissionRegistry.Personnel_ImportExport);
                    perms.Add(PermissionRegistry.Menu_Salary);
                    perms.Add(PermissionRegistry.Salary_Export);
                    perms.Add(PermissionRegistry.Menu_AnnualIncome);
                    perms.Add(PermissionRegistry.Income_Export);
                    perms.Add(PermissionRegistry.Menu_LeaveDetail);
                    perms.Add(PermissionRegistry.Menu_PositionDuration);
                    perms.Add(PermissionRegistry.Menu_EmulationReward);
                    perms.Add(PermissionRegistry.Menu_Evaluation);
                    perms.Add(PermissionRegistry.Menu_Training);
                    perms.Add(PermissionRegistry.Menu_Planning);
                    perms.Add(PermissionRegistry.Menu_Rotation);
                    perms.Add(PermissionRegistry.Menu_DutyAssignment);
                }
                else
                {
                    // Nhân viên: Xem được các menu cơ bản
                    perms.Add(PermissionRegistry.Menu_Dashboard);
                    perms.Add(PermissionRegistry.Menu_Statistics);
                    perms.Add(PermissionRegistry.Menu_Personnel);
                    perms.Add(PermissionRegistry.Menu_Salary);
                    perms.Add(PermissionRegistry.Menu_AnnualIncome);
                    perms.Add(PermissionRegistry.Menu_LeaveDetail);
                    perms.Add(PermissionRegistry.Menu_PositionDuration);
                    perms.Add(PermissionRegistry.Menu_EmulationReward);
                    perms.Add(PermissionRegistry.Menu_Evaluation);
                    perms.Add(PermissionRegistry.Menu_Training);
                    perms.Add(PermissionRegistry.Menu_Planning);
                    perms.Add(PermissionRegistry.Menu_Rotation);
                    perms.Add(PermissionRegistry.Menu_DutyAssignment);
                }
            }

            _cachedCurrentUserPermissions = perms;
        }

        /// <summary>
        /// Kiểm tra xem người dùng hiện tại có quyền thực hiện thao tác hay không.
        /// Tốc độ thực thi cực nhanh (O(1)) nhờ bộ nhớ cache in-memory.
        /// </summary>
        public static bool HasPermission(string permissionCode)
        {
            if (App.CurrentUser == null) return false;

            EnsureCacheLoaded();

            if (_cachedCurrentUserPermissions == null) return false;

            // Quyền wildcard toàn quyền
            if (_cachedCurrentUserPermissions.Contains("*")) return true;

            return _cachedCurrentUserPermissions.Contains(permissionCode);
        }

        /// <summary>
        /// Kiểm tra quyền, nếu không có quyền thì hiển thị hộp thoại cảnh báo và trả về false.
        /// </summary>
        public static bool CheckPermissionOrAlert(Window? owner, string permissionCode, string featureName)
        {
            if (HasPermission(permissionCode)) return true;

            string message = $"Tài khoản của bạn không được phân quyền để thực hiện:\n\n• {featureName}\n\nVui lòng liên hệ Quản trị viên để được cấp quyền này.";
            var warning = new WarningWindow(message, "Từ chối truy cập");
            if (owner != null) warning.Owner = owner;
            warning.ShowDialog();

            return false;
        }

        /// <summary>
        /// Chuyển đổi danh sách mã quyền sang chuỗi JSON lưu vào CSDL.
        /// </summary>
        public static string SerializePermissions(IEnumerable<string> permissions)
        {
            return JsonSerializer.Serialize(permissions.Distinct().ToList());
        }

        /// <summary>
        /// Giải mã chuỗi JSON từ CSDL thành danh sách mã quyền.
        /// </summary>
        public static List<string> DeserializePermissions(string? json)
        {
            if (string.IsNullOrWhiteSpace(json)) return new List<string>();
            try
            {
                return JsonSerializer.Deserialize<List<string>>(json) ?? new List<string>();
            }
            catch
            {
                return new List<string>();
            }
        }
    }
}

