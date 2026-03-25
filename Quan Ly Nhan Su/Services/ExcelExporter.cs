using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using TaxPersonnelManagement.Models;

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

                // Tiêu đề các cột (35 cột, từ A đến AI)
                string[] headers = 
                {
                    "STT", "Số hiệu CB", "Họ và Tên", // A, B, C
                    "Giới tính", "Ngày sinh", "Dân tộc", "Tôn giáo", "Nơi sinh", // D, E, F, G, H
                    "SĐT", "Email", // I, J
                    "CCCD", "Nơi cấp CCCD", "Số BHXH", // K, L, M
                    "Phòng ban", "Chức vụ", "Ngày về hưu (Dự kiến)", // N, O, P
                    "Trình độ CM", "Chuyên ngành", "Trường đào tạo", // Q, R, S
                    "Lý luận CT", "QL Nhà nước", "Ngoại ngữ", "Tin học", // T, U, V, W
                    "Ngày vào Đảng", "Ngày chính thức", // X, Y
                    "Mã ngạch", "Tên ngạch", "Bậc lương", "Hệ số", // Z, AA, AB, AC
                    "Phụ cấp CV", "Vượt khung %", // AD, AE
                    "Danh hiệu thi đua", "Khen thưởng", "Kỷ luật", // AF, AG, AH
                    "Ghi chú" // AI
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
                    worksheet.Cell(row, 5).Value = p.DateOfBirth; 
                    worksheet.Cell(row, 5).Style.DateFormat.Format = "dd/MM/yyyy";

                    worksheet.Cell(row, 6).Value = p.Ethnicity;
                    worksheet.Cell(row, 7).Value = p.Religion;
                    worksheet.Cell(row, 8).Value = p.BirthPlace;
                    
                    worksheet.Cell(row, 9).Value = p.PhoneNumber;
                    worksheet.Cell(row, 10).Value = p.Email;
                    
                    worksheet.Cell(row, 11).Value = "'" + p.IdentityCardNumber; // Force text to avoid scientific notation
                    worksheet.Cell(row, 12).Value = p.IdentityCardPlace;
                    worksheet.Cell(row, 13).Value = "'" + p.SocialSecurityNumber;

                    worksheet.Cell(row, 14).Value = p.Department;
                    worksheet.Cell(row, 15).Value = p.Position;
                    worksheet.Cell(row, 16).Value = p.RetirementDate;
                     worksheet.Cell(row, 16).Style.DateFormat.Format = "dd/MM/yyyy";

                    worksheet.Cell(row, 17).Value = p.EducationLevel;
                    worksheet.Cell(row, 18).Value = p.Major;
                    worksheet.Cell(row, 19).Value = p.University;

                    worksheet.Cell(row, 20).Value = p.PoliticalTheoryLevel;
                    worksheet.Cell(row, 21).Value = p.StateManagementLevel;
                    worksheet.Cell(row, 22).Value = p.LanguageSkillLevel;
                    worksheet.Cell(row, 23).Value = p.ITSkillLevel;

                    worksheet.Cell(row, 24).Value = p.PartyEntryDate;
                    worksheet.Cell(row, 24).Style.DateFormat.Format = "dd/MM/yyyy";
                    worksheet.Cell(row, 25).Value = p.PartyOfficialDate;
                    worksheet.Cell(row, 25).Style.DateFormat.Format = "dd/MM/yyyy";

                    worksheet.Cell(row, 26).Value = p.RankCode;
                    worksheet.Cell(row, 27).Value = p.RankName;
                    worksheet.Cell(row, 28).Value = "'" + p.CurrentSalaryStep; // e.g. "1/9" can be interpreted as date
                    worksheet.Cell(row, 29).Value = p.CurrentSalaryCoefficient;

                    worksheet.Cell(row, 30).Value = p.PositionAllowance;
                    worksheet.Cell(row, 31).Value = p.ExceedFramePercent;

                    worksheet.Cell(row, 32).Value = p.EmulationTitles;
                    worksheet.Cell(row, 33).Value = p.RewardForms;
                    worksheet.Cell(row, 34).Value = p.DisciplineType == "---" ? "" : p.DisciplineType;

                    // Cột 35: Ghi chú - Theo dõi nghỉ thai sản (chưa đủ 36 tháng) và nghỉ ốm
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
                                ghiChu = $"Chưa đủ 36 tháng ({maternityLeave.StartDate:dd/MM/yyyy}-{endOf36Months:dd/MM/yyyy})";
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
                                ghiChu = $"Đang nghỉ ốm ({sickLeave.StartDate:dd/MM/yyyy}-{sickLeave.EndDate:dd/MM/yyyy})";
                            }
                        }
                    }
                    worksheet.Cell(row, 35).Value = ghiChu;

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
    }
}
