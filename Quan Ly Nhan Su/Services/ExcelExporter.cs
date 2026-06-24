using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using TaxPersonnelManagement.Models;
using TaxPersonnelManagement.Helpers;

namespace TaxPersonnelManagement.Services
{
    /// <summary>
    /// Lớp xuất danh sách nhân sự ra file Excel (.xlsx).
    /// Bao gồm toàn bộ thông tin cá nhân, lương, khen thưởng, kỷ luật và cột Ghi chú.
    /// </summary>
    public static class ExcelExporter
    {
        /// <summary>
        /// Xuất danh sách nhân sự ra file Excel tại đường dẫn chỉ định.
        /// Cột Ghi chú tự động tính toán trạng thái nghỉ thai sản (chưa đủ 36 tháng) và nghỉ ốm.
        /// </summary>
        /// <param name="personnelList">Danh sách nhân sự cần xuất.</param>
        /// <param name="filePath">Đường dẫn lưu file Excel.</param>
        public static void Export(IEnumerable<Personnel> personnelList, string filePath)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("DanhSachNhanSu");

                // Tiêu đề các cột (40 cột)
                string[] headers =
                {
                    "STT", "Số hiệu CB", "Họ và Tên", // A, B, C
                    "Giới tính", "Ngày sinh", "Dân tộc", "Tôn giáo", "Nơi sinh", "Quê quán", "Nơi ở hiện nay", // D, E, F, G, H, I, J
                    "SĐT", "Email", // K, L
                    "CCCD", "Nơi cấp CCCD", "Số BHXH", // M, N, O
                    "Bộ phận", "Chức vụ", "Thời gian công tác tại cơ quan thuế", "Số năm công tác", // P, Q, R, S
                    "Trình độ CM", "Chuyên ngành", "Trường đào tạo", // T, U, V
                    "Lý luận CT", "QL Nhà nước", "Ngoại ngữ", "Tin học", // W, X, Y, Z
                    "Đảng viên", "Ngày vào Đảng", "Ngày chính thức", "Số năm tuổi Đảng", // AA, AB, AC, AD
                    "Mã ngạch", "Tên ngạch", "Bậc lương", "Hệ số", // AE, AF, AG, AH
                    "Phụ cấp CV", "Vượt khung %", // AI, AJ
                    "Danh hiệu thi đua", "Khen thưởng", "Kỷ luật", // AK, AL, AM
                    "Ghi chú", // AN
                    "Ngày về hưu (Dự kiến)" // AO
                };

                for (int i = 0; i < headers.Length; i++)
                {
                    worksheet.Cell(1, i + 1).Value = headers[i];
                    worksheet.Cell(1, i + 1).Style.Font.Bold = true;
                    worksheet.Cell(1, i + 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#1976D2"); // Blue header
                    worksheet.Cell(1, i + 1).Style.Font.FontColor = XLColor.White;
                    worksheet.Cell(1, i + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Cell(1, i + 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                }

                // Ghi dữ liệu từng dòng
                int row = 2;
                int stt = 1;
                var now = DateTime.Now.Date;

                foreach (var p in personnelList)
                {
                    worksheet.Cell(row, 1).Value = stt++;
                    worksheet.Cell(row, 2).Value = p.StaffId;
                    worksheet.Cell(row, 3).Value = p.FullName;

                    worksheet.Cell(row, 4).Value = p.Gender;
                    worksheet.Cell(row, 5).Value = p.DateOfBirth.HasValue ? DatePickerHelper.FormatDateForDisplay(p.DateOfBirth.Value) : "";

                    worksheet.Cell(row, 6).Value = p.Ethnicity;
                    worksheet.Cell(row, 7).Value = p.Religion;
                    worksheet.Cell(row, 8).Value = p.BirthPlace;
                    worksheet.Cell(row, 9).Value = p.Hometown;
                    worksheet.Cell(row, 10).Value = p.CurrentResidence;

                    worksheet.Cell(row, 11).Value = p.PhoneNumber;
                    worksheet.Cell(row, 12).Value = p.Email;

                    worksheet.Cell(row, 13).Value = "'" + p.IdentityCardNumber; // Force text to avoid scientific notation
                    worksheet.Cell(row, 14).Value = p.IdentityCardPlace;
                    worksheet.Cell(row, 15).Value = p.SocialSecurityNumber;

                    worksheet.Cell(row, 16).Value = p.Department;
                    worksheet.Cell(row, 17).Value = p.Position;
                    worksheet.Cell(row, 18).Value = p.TaxAuthorityStartDate.HasValue ? DatePickerHelper.FormatDateForDisplay(p.TaxAuthorityStartDate.Value) : "";
 
                    worksheet.Cell(row, 19).Value = CalcWorkingYears(p, now);

                    worksheet.Cell(row, 20).Value = p.EducationLevel;
                    worksheet.Cell(row, 21).Value = p.Major;
                    worksheet.Cell(row, 22).Value = p.University;

                    worksheet.Cell(row, 23).Value = p.PoliticalTheoryLevel;
                    worksheet.Cell(row, 24).Value = p.StateManagementLevel;
                    worksheet.Cell(row, 25).Value = p.LanguageSkillLevel;
                    worksheet.Cell(row, 26).Value = p.ITSkillLevel;

                    worksheet.Cell(row, 27).Value = p.PartyEntryDate.HasValue ? "Đảng viên" : "";
                    worksheet.Cell(row, 28).Value = p.PartyEntryDate.HasValue ? DatePickerHelper.FormatDateForDisplay(p.PartyEntryDate.Value) : "";
                    worksheet.Cell(row, 29).Value = p.PartyOfficialDate.HasValue ? DatePickerHelper.FormatDateForDisplay(p.PartyOfficialDate.Value) : "";
                    worksheet.Cell(row, 30).Value = p.CalculatedPartyAge;

                    worksheet.Cell(row, 31).Value = p.RankCode;
                    worksheet.Cell(row, 32).Value = p.RankName;
                    worksheet.Cell(row, 33).Value = "'" + p.CurrentSalaryStep; // e.g. "1/9" can be interpreted as date
                    worksheet.Cell(row, 34).Value = p.CurrentSalaryCoefficient;

                    worksheet.Cell(row, 35).Value = p.PositionAllowance;
                    worksheet.Cell(row, 36).Value = p.ExceedFramePercent > 0 ? $"{p.ExceedFramePercent}%" : "";

                    worksheet.Cell(row, 37).Value = p.EmulationTitles;
                    worksheet.Cell(row, 38).Value = p.RewardForms;
                    worksheet.Cell(row, 39).Value = p.DisciplineType == "---" ? "" : p.DisciplineType;

                    // Cột 40: Ghi chú - Theo dõi nghỉ thai sản (chưa đủ 36 tháng) và nghỉ ốm
                    string ghiChu = "";
                    if (p.LeaveHistories != null)
                    {
                        // Kiểm tra nghỉ thai sản: nếu chưa đủ 36 tháng từ ngày bắt đầu → ghi chú
                        var maternityLeave = p.LeaveHistories
                            .Where(l => (l.LeaveType == "Thai sản" || l.LeaveType == "Nghỉ thai sản"))
                            .OrderByDescending(l => l.StartDate)
                            .FirstOrDefault();

                        if (maternityLeave != null)
                        {
                            var endOf36Months = maternityLeave.StartDate.AddMonths(36);
                            if (now < endOf36Months)
                            {
                                ghiChu = $"Chưa đủ 36 tháng ({DatePickerHelper.FormatDateForDisplay(maternityLeave.StartDate)}-{DatePickerHelper.FormatDateForDisplay(endOf36Months)})";
                            }
                        }

                        // Kiểm tra nghỉ ốm: nếu công chức đang trong thời gian nghỉ ốm → ghi chú
                        if (string.IsNullOrEmpty(ghiChu))
                        {
                            var sickLeave = p.LeaveHistories
                                .Where(l => (l.LeaveType == "Nghỉ ốm" || l.LeaveType.Contains("ốm")) &&
                                            l.StartDate <= now && l.EndDate >= now)
                                .OrderByDescending(l => l.StartDate)
                                .FirstOrDefault();

                            if (sickLeave != null)
                            {
                                ghiChu = $"Đang nghỉ ốm ({DatePickerHelper.FormatDateForDisplay(sickLeave.StartDate)}-{(sickLeave.EndDate.HasValue ? DatePickerHelper.FormatDateForDisplay(sickLeave.EndDate.Value) : "")})";
                            }
                        }
                    }
                    worksheet.Cell(row, 40).Value = ghiChu;
                    worksheet.Cell(row, 41).Value = p.RetirementDate.HasValue ? DatePickerHelper.FormatDateForDisplay(p.RetirementDate.Value) : "";

                    // Áp dụng viền và căn chỉnh cho toàn bộ ô trong dòng
                    for (int c = 1; c <= headers.Length; c++)
                    {
                        worksheet.Cell(row, c).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        worksheet.Cell(row, c).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    }

                    row++;
                }

                // Tự động điều chỉnh độ rộng cột theo nội dung
                worksheet.Columns().AdjustToContents();

                // Lưu file Excel
                workbook.SaveAs(filePath);
            }
        }

        private static string CalcWorkingYears(Personnel p, DateTime now)
        {
            if (!p.TaxAuthorityStartDate.HasValue) return "";
            DateTime start = p.TaxAuthorityStartDate.Value;
            if (start > now) return "0 năm";

            DateTime temp = start;
            int y = 0;
            while (temp.AddYears(1) <= now) { y++; temp = temp.AddYears(1); }
            int m = 0;
            while (temp.AddMonths(1) <= now) { m++; temp = temp.AddMonths(1); }
            int d = (now - temp).Days;

            if (m == 0 && d == 0) return $"{y} năm";
            if (d == 0) return $"{y} năm {m} tháng";
            return $"{y} năm {m} tháng {d} ngày";
        }
    }
}
