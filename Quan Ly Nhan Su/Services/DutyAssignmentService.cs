using System;
using System.Collections.Generic;
using System.Linq;
using TaxPersonnelManagement.Data;
using TaxPersonnelManagement.Models;

namespace TaxPersonnelManagement.Services
{
    /// <summary>
    /// Thông tin hiển thị của một phiên bản phân công nhiệm vụ (không chứa nội dung file PDF để tải danh sách nhanh).
    /// </summary>
    public class DutyVersionItem
    {
        public int Id { get; set; }
        public string Department { get; set; } = string.Empty;
        public int Year { get; set; }
        public int Month { get; set; }
        public DateTime EffectiveDate { get; set; }
        public string? Title { get; set; }
        public string? Note { get; set; }
        public string FileName { get; set; } = string.Empty;
        public long FileSize { get; set; }
        public DateTime UploadedAt { get; set; }
        public string? UploadedBy { get; set; }

        /// <summary>True nếu đây là phiên bản đang có hiệu lực tại thời điểm hiện tại</summary>
        public bool IsCurrent { get; set; }

        public string MonthText => $"Tháng {Month}/{Year}";
        public string EffectiveText => $"Hiệu lực từ {EffectiveDate:dd/MM/yyyy}";
        public string DisplayTitle => string.IsNullOrWhiteSpace(Title) ? "Phân công nhiệm vụ" : Title!;
        public string SizeText => FileSize >= 1024 * 1024
            ? $"{FileSize / 1024d / 1024d:0.0} MB"
            : $"{Math.Max(1, FileSize / 1024)} KB";
        public string UploadInfo => $"Tải lên {UploadedAt:dd/MM/yyyy HH:mm}" + (string.IsNullOrWhiteSpace(UploadedBy) ? "" : $" • {UploadedBy}");
    }

    /// <summary>
    /// Truy vấn dữ liệu cho chức năng Phân công nhiệm vụ.
    /// </summary>
    public static class DutyAssignmentService
    {
        /// <summary>Danh sách Tổ / bộ phận (danh mục + các tên đã từng có file), sắp xếp theo quy ước chung.</summary>
        public static List<string> GetDepartments()
        {
            using var db = new AppDbContext();
            var names = db.Departments.Select(d => d.Name).ToList();
            var fromRecords = db.DutyAssignmentRecords.Select(r => r.Department).Distinct().ToList();

            return names
                .Concat(fromRecords)
                .Where(n => !string.IsNullOrWhiteSpace(n))
                .Select(n => n.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(n => DepartmentSorter.GetSortKey(n).Order)
                .ThenBy(n => DepartmentSorter.GetSortKey(n).Number)
                .ThenBy(n => DepartmentSorter.GetSortKey(n).Name)
                .ToList();
        }

        /// <summary>
        /// Lấy toàn bộ phiên bản (không kèm nội dung file), sắp xếp mới nhất trước.
        /// Đánh dấu IsCurrent cho phiên bản đang có hiệu lực của mỗi Tổ.
        /// </summary>
        public static List<DutyVersionItem> GetVersions(string? department = null)
        {
            using var db = new AppDbContext();
            var query = db.DutyAssignmentRecords.AsQueryable();
            if (!string.IsNullOrWhiteSpace(department))
                query = query.Where(r => r.Department == department);

            var list = query
                .Select(r => new DutyVersionItem
                {
                    Id = r.Id,
                    Department = r.Department,
                    Year = r.Year,
                    Month = r.Month,
                    EffectiveDate = r.EffectiveDate,
                    Title = r.Title,
                    Note = r.Note,
                    FileName = r.FileName,
                    FileSize = r.FileSize,
                    UploadedAt = r.UploadedAt,
                    UploadedBy = r.UploadedBy
                })
                .AsEnumerable()
                .OrderByDescending(r => r.EffectiveDate)
                .ThenByDescending(r => r.Year * 100 + r.Month)
                .ThenByDescending(r => r.Id)
                .ToList();

            // Phiên bản hiện hành của từng Tổ = mới nhất có ngày hiệu lực <= hôm nay
            var today = DateTime.Today;
            foreach (var group in list.GroupBy(r => r.Department))
            {
                var current = group.FirstOrDefault(r => r.EffectiveDate.Date <= today);
                if (current != null) current.IsCurrent = true;
            }
            return list;
        }

        /// <summary>Đọc nội dung file PDF của một phiên bản.</summary>
        public static byte[]? LoadFileData(int id)
        {
            using var db = new AppDbContext();
            return db.DutyAssignmentRecords
                .Where(r => r.Id == id)
                .Select(r => r.FileData)
                .FirstOrDefault();
        }

        public static void Delete(int id)
        {
            using var db = new AppDbContext();
            var record = db.DutyAssignmentRecords.Find(id);
            if (record != null)
            {
                db.DutyAssignmentRecords.Remove(record);
                db.SaveChanges();
            }
        }
    }
}
