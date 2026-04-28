using Microsoft.EntityFrameworkCore;
using TaxPersonnelManagement.Models;
using System.IO;

namespace TaxPersonnelManagement.Data
{
    /// <summary>
    /// Lớp ngữ cảnh cơ sở dữ liệu chính của ứng dụng.
    /// Kế thừa từ DbContext của Entity Framework Core để quản lý kết nối và ánh xạ các model (lớp) với các bảng trong cơ sở dữ liệu SQLite.
    /// </summary>
    public class AppDbContext : DbContext
    {
        // Danh sách các bảng (DbSet) trong Cơ sở dữ liệu
        public DbSet<User> Users { get; set; }                           // Bảng người dùng hệ thống (tài khoản đăng nhập)
        public DbSet<Personnel> Personnel { get; set; }                  // Bảng hồ sơ công chức/viên chức
        public DbSet<Department> Departments { get; set; }               // Bảng danh mục phòng ban
        public DbSet<Position> Positions { get; set; }                   // Bảng danh mục chức vụ
        public DbSet<Rank> Ranks { get; set; }                           // Bảng danh mục ngạch (ví dụ: Chuyên viên chính)
        public DbSet<SalaryRecord> SalaryRecords { get; set; }           // Bảng lịch sử quá trình lương
        public DbSet<LeaveHistory> LeaveHistories { get; set; }          // Bảng lịch sử nghỉ phép/công tác
        public DbSet<RankSalarySpec> RankSalarySpecs { get; set; }       // Bảng cấu hình hệ số lương theo ngạch/bậc
        public DbSet<SalaryDelayReason> SalaryDelayReasons { get; set; } // Bảng danh mục lý do chậm nâng lương
        public DbSet<DisciplineType> DisciplineTypes { get; set; }       // Bảng danh mục các loại kỷ luật
        public DbSet<IncomeRecord> IncomeRecords { get; set; }           // Bảng ghi nhận tổng thu nhập hàng tháng/năm
        public DbSet<PublicHoliday> PublicHolidays { get; set; }         // Bảng danh mục ngày nghỉ lễ


        /// <summary>
        /// Phương thức cấu hình cơ sở dữ liệu khi ứng dụng khởi chạy.
        /// Sử dụng cơ sở dữ liệu SQLite cục bộ được lưu trong file 'tax_personnel.db' cùng thư mục chạy ứng dụng.
        /// </summary>
        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            // Thiết lập đường dẫn tới file SQLite cục bộ
            string dbPath = Path.Combine(System.AppContext.BaseDirectory, "tax_personnel.db");
            optionsBuilder.UseSqlite($"Data Source={dbPath}");
        }

        /// <summary>
        /// Phương thức cấu hình Model khi khởi tạo CSDL lần đầu, dùng để Seed (khởi tạo) dữ liệu mặc định.
        /// </summary>
        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

            // Khởi tạo tài khoản Quản trị viên (Admin) mặc định khi tạo CSDL mới
            modelBuilder.Entity<User>().HasData(new User
            {
                Id = 1,
                Username = "admin",
                PasswordHash = "admin", // Trong thực tế, mật khẩu cần được băm (hash) để bảo mật!
                FullName = "Administrator",
                Role = UserRole.Admin   // Quyền Quản trị viên có toàn quyền truy cập
            });
        }
    }
}
