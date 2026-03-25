using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace TaxPersonnelManagement.Models
{
    /// <summary>
    /// Các vai trò (quyền hạn) của người dùng trong hệ thống.
    /// Quyết định xem người dùng có thể xem/thêm/sửa/xóa dữ liệu nào.
    /// </summary>
    public enum UserRole
    {
        Admin,      // Quản trị viên (Toàn quyền)
        Manager,    // Lãnh đạo (Xem và phê duyệt báo cáo)
        Staff       // Nhân viên (Chỉ xem dữ liệu, không được chỉnh sửa)
    }

    /// <summary>
    /// Model lưu trữ thông tin tài khoản dùng để đăng nhập vào phần mềm.
    /// </summary>
    public class User
    {
        [Key]
        public int Id { get; set; }
        [Required]
        public string Username { get; set; } = string.Empty; // Tên đăng nhập
        [Required]
        public string PasswordHash { get; set; } = string.Empty; // Mật khẩu (bản rõ hoặc mã hóa)
        public string FullName { get; set; } = string.Empty; // Họ và tên người dùng
        public UserRole Role { get; set; } // Phân quyền
    }

    /// <summary>
    /// Model cốt lõi lưu trữ toàn bộ hồ sơ chi tiết của một Cán bộ / Công chức.
    /// Bao gồm các thông tin cá nhân, công tác, lương, đảng phí, v.v.
    /// </summary>
    public class Personnel
    {
        [Key]
        public int Id { get; set; }
        [System.ComponentModel.DataAnnotations.Schema.NotMapped]
        public int STT { get; set; }
        public string? StaffId { get; set; } // Số hiệu cán bộ
        [Required]
        public string? FullName { get; set; } = string.Empty;
        public DateTime? DateOfBirth { get; set; }
        public string? Gender { get; set; } // Nam/Nữ
        public string? PhoneNumber { get; set; } // Số điện thoại
        public string? IdentityCardNumber { get; set; } // CCCD/CMND
        public string? IdentityCardPlace { get; set; } // Nơi cấp
        public string? SocialSecurityNumber { get; set; } // Mã BHXH
        public string? Email { get; set; }
        public string? BirthPlace { get; set; } // Nơi sinh
        public string? Ethnicity { get; set; } // Dân tộc
        public string? Religion { get; set; } // Tôn giáo

        // Work Info
        public string? Department { get; set; } // Phòng ban/Bộ phận
        public string? Position { get; set; } // Chức vụ
        public string? RankCode { get; set; } // Mã ngạch
        public string? RankName { get; set; } // Tên ngạch
        public DateTime? TaxAuthorityStartDate { get; set; } // Thời gian công tác tại Cơ quan Thuế
        public DateTime? StartDate { get; set; } // Ngày bắt đầu (General)
        public string? Status { get; set; } // Trạng thái (Đang công tác, Nghỉ việc...)

        // Tab 2: Position History Info
        public DateTime? PositionDecisionDate { get; set; } // Thời gian công tác tính theo QĐ gần nhất
        public DateTime? PositionCalculationDate { get; set; } // Thời điểm tính thời gian công tác
        public string? PositionYear { get; set; } // Năm giữ vị trí công tác
        public string? DetailedWorkHistory { get; set; } // Quá trình công tác chi tiết

        // Tab 3: Retirement Info
        public DateTime? RetirementDate { get; set; } // Ngày về hưu

        // Tab 5: Leave Info
        public int TotalAnnualLeaveDays { get; set; } = 12; // Mặc định 12 ngày
        public List<LeaveHistory> LeaveHistories { get; set; } = new List<LeaveHistory>();

        // Tab 4: Party Info
        public DateTime? PartyEntryDate { get; set; } // Ngày vào Đảng
        public DateTime? PartyOfficialDate { get; set; } // Ngày chính thức

        // Education
        public string? EducationLevel { get; set; } // Trình độ văn hóa (12/12)
        public string? Major { get; set; } // Chuyên ngành
        public string? University { get; set; } // Trường tốt nghiệp
        public string? StateManagementLevel { get; set; } // QLNN
        public string? PoliticalTheoryLevel { get; set; } // LLCT
        public string? ITSkillLevel { get; set; } // Tin học
        public string? LanguageSkillLevel { get; set; } // Ngoại ngữ

        // Avatar
        public string? AvatarBase64 { get; set; } // Ảnh đại diện (Base64 string)

        // Tab 6: Salary Info
        public string? CurrentSalaryStep { get; set; } // Bậc lương
        public double CurrentSalaryCoefficient { get; set; } // Hệ số lương
        public double ExceedFramePercent { get; set; } // % Vượt khung
        public string? PositionAllowance { get; set; } // PC Chức vụ
        public DateTime? SalaryReservationDeadline { get; set; } // Thời hạn bảo lưu
        public DateTime? NextSalaryStepDate { get; set; } // Thời điểm tính bậc lương lần sau
        public string? SalaryIncreaseDelayType { get; set; } // Lùi thời gian nâng lương
        public DateTime? ExpectedSalaryIncreaseDate { get; set; } // Dự kiến lên lương
        public string? SalaryHistoryLog { get; set; } // Diễn biến quá trình lương

        // Tab 7: Reward Info
        public string? EmulationTitles { get; set; } // Danh hiệu thi đua
        public string? RewardForms { get; set; } // Hình thức khen thưởng

        // Tab 8: Discipline Info
        public string? DisciplineType { get; set; } // Hình thức kỷ luật (Configurable)
        public string? DisciplineDecisionNumber { get; set; } // Số quyết định
        public DateTime? DisciplineDecisionDate { get; set; } // Ngày ký QĐ
        public string? DisciplineReason { get; set; } // Nội dung / Lý do kỷ luật

        // Navigation
        public List<SalaryRecord> SalaryRecords { get; set; } = new List<SalaryRecord>();

        [System.ComponentModel.DataAnnotations.Schema.NotMapped]
        public string SalaryIncreaseStatus
        {
            get
            {
                if (!ExpectedSalaryIncreaseDate.HasValue)
                    return string.Empty;

                var days = (ExpectedSalaryIncreaseDate.Value.Date - DateTime.Now.Date).TotalDays;
                if (days > 0)
                    return $"Còn {days} ngày";
                else if (days == 0)
                    return "Đến hạn hôm nay";
                else
                    return $"Quá hạn {Math.Abs(days)} ngày";
            }
        }
    }

    /// <summary>
    /// Bảng danh mục Phòng ban / Đội quản lý.
    /// </summary>
    public class Department
    {
        [Key]
        public int Id { get; set; }
        [Required]
        public string Name { get; set; } = string.Empty; // Tên phòng ban
    }

    /// <summary>
    /// Bảng danh mục Chức vụ (Ví dụ: Chi cục trưởng, Phó phòng...).
    /// </summary>
    public class Position
    {
        [Key]
        public int Id { get; set; }
        [Required]
        public string Name { get; set; } = string.Empty; // Tên chức vụ
    }

    /// <summary>
    /// Bảng danh mục Ngạch lương (Ví dụ: Chuyên viên chính, Kiểm tra viên thuế...).
    /// </summary>
    public class Rank
    {
        [Key]
        public int Id { get; set; }
        [Required]
        public string Code { get; set; } = string.Empty; // Mã ngạch (VD: 01.003)
        [Required]
        public string Name { get; set; } = string.Empty; // Tên ngạch (VD: Chuyên viên chính)
    }

    /// <summary>
    /// Bảng thiết lập Hệ số lương chuẩn theo từng Ngạch và Bậc.
    /// Dùng để tra cứu hệ số tự động khi người dùng chọn Ngạch và Bậc.
    /// </summary>
    public class RankSalarySpec
    {
        [Key]
        public int Id { get; set; }
        [Required]
        public string RankCode { get; set; } = string.Empty; // Liên kết với Rank.Code
        [Required]
        public string SalaryStep { get; set; } = string.Empty; // Bậc lương (VD: "1/9")
        public double Coefficient { get; set; } // Hệ số lương tương ứng (VD: 2.34)
    }

    /// <summary>
    /// Lịch sử quá trình lương của Cán bộ.
    /// Lưu trữ các quyết định nâng lương, thâm niên.
    /// </summary>
    public class SalaryRecord
    {
        [Key]
        public int Id { get; set; }
        public int PersonnelId { get; set; } // Khóa ngoại liên kết tới Cán bộ
        public Personnel? Personnel { get; set; }
        
        public string SalaryLevel { get; set; } = string.Empty; // Ngạch/Bậc lương
        public decimal Coefficient { get; set; } // Hệ số lương áp dụng
        public DateTime DecisionDate { get; set; } // Ngày ký quyết định
        public DateTime EffectiveDate { get; set; } // Ngày bắt đầu hưởng lương mới
        public DateTime NextIncreaseDate { get; set; } // Ngày nâng lương dự kiến tiếp theo
        public string? Note { get; set; } // Ghi chú thêm
    }

    public class LeaveHistory
    {
        [Key]
        public int Id { get; set; }
        public int PersonnelId { get; set; }
        public Personnel? Personnel { get; set; }

        [Required]
        public string LeaveType { get; set; } = string.Empty; // Nghỉ phép, Ốm, Thai sản...
        public int? LeaveYear { get; set; } // Năm nghỉ phép (cho Phép năm)
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public double DurationDays { get; set; }
        public string? Reason { get; set; }
        
        [System.ComponentModel.DataAnnotations.Schema.NotMapped]
        public int STT { get; set; }

        [System.ComponentModel.DataAnnotations.Schema.NotMapped]
        public string UserReasonDisplay 
        { 
            get 
            {
                if (string.IsNullOrEmpty(Reason)) return "";
                // Split by |SYS: first
                var parts = Reason.Split(new[] { "|SYS:" }, StringSplitOptions.None);
                return parts[0].Trim();
            }
        }

        [System.ComponentModel.DataAnnotations.Schema.NotMapped]
        public string SystemMessageDisplay 
        { 
            get 
            {
                // First, try to extract specific SYS message
                string sysMsg = "";
                if (!string.IsNullOrEmpty(Reason))
                {
                    var parts = Reason.Split(new[] { "|SYS:" }, StringSplitOptions.None);
                    if (parts.Length > 1)
                    {
                        var sysPart = parts[1];
                        var linkParts = sysPart.Split(new[] { "|LINK:" }, StringSplitOptions.None);
                        sysMsg = linkParts[0].Trim();
                    }
                }

                // If a sys message exists, prioritize it.
                if (!string.IsNullOrEmpty(sysMsg))
                {
                    return sysMsg;
                }

                // If NO sys message exists AND it's NOT Annual Leave ("Phép năm"), 
                // calculate and display "X tháng Y ngày"
                if (LeaveType != "Phép năm")
                {
                    DateTime start = StartDate.Date;
                    DateTime end = EndDate.Date;
                    if (start <= end)
                    {
                        int months = 0;
                        DateTime tempDate = start;
                        while (tempDate.AddMonths(1) <= end)
                        {
                            months++;
                            tempDate = tempDate.AddMonths(1);
                        }
                        int days = (end - tempDate).Days + 1; // +1 to make it inclusive

                        if (months > 0 && days > 0)
                            return $"{months} tháng {days} ngày";
                        else if (months > 0)
                            return $"{months} tháng";
                        else if (days > 0)
                            return $"{days} ngày";
                    }
                }

                return ""; // Default empty
            }
        }

        [System.ComponentModel.DataAnnotations.Schema.NotMapped]
        public string? LinkId
        {
            get
            {
                if (string.IsNullOrEmpty(Reason)) return null;
                var parts = Reason.Split(new[] { "|LINK:" }, StringSplitOptions.None);
                if (parts.Length > 1)
                {
                    return parts[1].Trim();
                }
                return null;
            }
        }

        [System.ComponentModel.DataAnnotations.Schema.NotMapped]
        public bool ShowDeleteButton { get; set; } = true; // Default true so existing records show link

        [System.ComponentModel.DataAnnotations.Schema.NotMapped]
        public bool IsLinkedGroupHead { get; set; } = false; // True only if it's the first of a MULTI-record group
    }
    public class SalaryDelayReason
    {
        [Key]
        public int Id { get; set; }
        [Required]
        public string Name { get; set; } = string.Empty; // e.g. "Lùi 3 tháng (Khiển trách)"
    }

    public class DisciplineType
    {
        [Key]
        public int Id { get; set; }
        [Required]
        public string Name { get; set; } = string.Empty; // e.g. "Khiển trách", "Cảnh cáo"
    }

    public class IncomeRecord
    {
        [Key]
        public int Id { get; set; }
        public int PersonnelId { get; set; } // Foreign key to Personnel table
        public int Year { get; set; }
        public int Month { get; set; }
        public string IncomeType { get; set; } = string.Empty; // "Lương", "Làm thêm giờ", "Thu nhập khác"
        public decimal Amount { get; set; }
        public string? Note { get; set; }
    }
}
