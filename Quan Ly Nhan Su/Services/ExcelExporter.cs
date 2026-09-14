using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using Microsoft.EntityFrameworkCore;
using TaxPersonnelManagement.Data;
using TaxPersonnelManagement.Models;
using TaxPersonnelManagement.Helpers;

namespace TaxPersonnelManagement.Services
{
    /// <summary>
    /// Lớp xuất danh sách nhân sự ra file Excel (.xlsx).
    /// Bao gồm toàn bộ thông tin cá nhân, lương, khen thưởng, kỷ luật, toàn bộ văn bằng chứng chỉ và cột Ghi chú.
    /// </summary>
    public static class ExcelExporter
    {
        /// <summary>
        /// Xuất danh sách nhân sự ra file Excel tại đường dẫn chỉ định.
        /// Các bằng cấp, chứng chỉ nếu có nhiều dòng sẽ được hiển thị nhiều dòng (Alt+Enter) trong cùng 1 ô.
        /// </summary>
        /// <param name="personnelList">Danh sách nhân sự cần xuất.</param>
        /// <param name="filePath">Đường dẫn lưu file Excel.</param>
        public static void Export(IEnumerable<Personnel> personnelList, string filePath)
        {
            var personnelListArray = personnelList as Personnel[] ?? personnelList.ToArray();
            var ids = personnelListArray.Select(p => p.Id).Where(id => id > 0).ToList();

            // Tải toàn bộ văn bằng từ CSDL để đảm bảo đầy đủ ngay cả khi danh sách truyền vào chưa được Include
            Dictionary<int, List<PersonnelDegree>> degreesByPersonnel = new();
            if (ids.Any())
            {
                try
                {
                    using var db = new AppDbContext();
                    degreesByPersonnel = db.PersonnelDegrees
                        .Where(d => ids.Contains(d.PersonnelId))
                        .AsNoTracking()
                        .ToList()
                        .GroupBy(d => d.PersonnelId)
                        .ToDictionary(g => g.Key, g => g.ToList());
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("Error loading degrees for export: " + ex.Message);
                }
            }

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("DanhSachNhanSu");

                // Tiêu đề các cột (42 cột chuẩn)
                string[] headers =
                {
                    "STT", "Số hiệu CB", "Họ và Tên", // A, B, C (1, 2, 3)
                    "Giới tính", "Ngày sinh", "Dân tộc", "Tôn giáo", "Nơi sinh", "Quê quán", "Nơi ở hiện nay", // D-J (4-10)
                    "SĐT", "Email", // K, L (11, 12)
                    "CCCD", "Nơi cấp CCCD", "Số BHXH", // M, N, O (13, 14, 15)
                    "Bộ phận", "Chức vụ", "Thời gian công tác tại cơ quan thuế", "Số năm công tác", "Thời gian công tác theo QĐ gần nhất", // P, Q, R, S, T (16, 17, 18, 19, 20)
                    "Trình độ CM", "Chuyên ngành", "Trường đào tạo", // U, V, W (21, 22, 23)
                    "Lý luận CT", "QL Nhà nước", "Ngoại ngữ", "Tin học", // X, Y, Z, AA (24, 25, 26, 27)
                    "Đảng viên", "Ngày vào Đảng", "Ngày chính thức", "Số năm tuổi Đảng", // AB, AC, AD, AE (28, 29, 30, 31)
                    "Mã ngạch", "Tên ngạch", "Bậc lương", "Hệ số", // AF, AG, AH, AI (32, 33, 34, 35)
                    "Phụ cấp CV", "Vượt khung %", // AJ, AK (36, 37)
                    "Danh hiệu thi đua", "Khen thưởng", "Kỷ luật", // AL, AM, AN (38, 39, 40)
                    "Ghi chú", // AO (41)
                    "Ngày về hưu (Dự kiến)" // AP (42)
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

                foreach (var p in personnelListArray)
                {
                    var allDegrees = GetDegreesForPersonnel(p, degreesByPersonnel);

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
                    worksheet.Cell(row, 20).Value = p.PositionDecisionDate.HasValue ? DatePickerHelper.FormatDateForDisplay(p.PositionDecisionDate.Value) : "";

                    // 1. Chuyên môn (Hiển thị nhiều dòng trong cùng 1 ô nếu có nhiều bằng)
                    var cmList = allDegrees.Where(d => d.DegreeType == "Chuyên môn")
                                           .OrderByDescending(d => d.IsPrimary)
                                           .ThenBy(d => d.Id)
                                           .ToList();
                    string trinhDoCM = cmList.Any()
                        ? string.Join("\n", cmList.Select(d => d.DegreeName))
                        : (p.EducationLevel ?? "");

                    string chuyenNganh = cmList.Any()
                        ? string.Join("\n", cmList.Select(d => d.Major ?? ""))
                        : (p.Major ?? "");

                    string truongDaoTao = cmList.Any()
                        ? string.Join("\n", cmList.Select(d => d.Institution ?? ""))
                        : (p.University ?? "");

                    worksheet.Cell(row, 21).Value = trinhDoCM;
                    worksheet.Cell(row, 22).Value = chuyenNganh;
                    worksheet.Cell(row, 23).Value = truongDaoTao;

                    // 2. Lý luận chính trị (Nhiều dòng nếu có nhiều bằng)
                    var polList = allDegrees.Where(d => d.DegreeType == "Lý luận chính trị")
                                            .OrderByDescending(d => d.IsPrimary)
                                            .ThenBy(d => d.Id)
                                            .ToList();
                    worksheet.Cell(row, 24).Value = polList.Any()
                        ? string.Join("\n", polList.Select(d => d.DegreeName))
                        : (p.PoliticalTheoryLevel ?? "");

                    // 3. Quản lý Nhà nước (Nhiều dòng nếu có nhiều bằng)
                    var stateList = allDegrees.Where(d => d.DegreeType == "Quản lý Nhà nước")
                                              .OrderByDescending(d => d.IsPrimary)
                                              .ThenBy(d => d.Id)
                                              .ToList();
                    worksheet.Cell(row, 25).Value = stateList.Any()
                        ? string.Join("\n", stateList.Select(d => d.DegreeName))
                        : (p.StateManagementLevel ?? "");

                    // 4. Ngoại ngữ (Nhiều dòng nếu có nhiều bằng, ví dụ: A, B, C...)
                    var langList = allDegrees.Where(d => d.DegreeType == "Ngoại ngữ")
                                             .OrderByDescending(d => d.IsPrimary)
                                             .ThenBy(d => d.Id)
                                             .ToList();
                    worksheet.Cell(row, 26).Value = langList.Any()
                        ? string.Join("\n", langList.Select(d => d.DegreeName))
                        : (p.LanguageSkillLevel ?? "");

                    // 5. Tin học (Nhiều dòng nếu có nhiều bằng, ví dụ: A, B, C...)
                    var itList = allDegrees.Where(d => d.DegreeType == "Tin học")
                                           .OrderByDescending(d => d.IsPrimary)
                                           .ThenBy(d => d.Id)
                                           .ToList();
                    worksheet.Cell(row, 27).Value = itList.Any()
                        ? string.Join("\n", itList.Select(d => d.DegreeName))
                        : (p.ITSkillLevel ?? "");

                    // Thông tin Đảng viên
                    worksheet.Cell(row, 28).Value = p.PartyEntryDate.HasValue ? 1 : "";
                    worksheet.Cell(row, 28).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Cell(row, 29).Value = p.PartyEntryDate.HasValue ? DatePickerHelper.FormatDateForDisplay(p.PartyEntryDate.Value) : "";
                    worksheet.Cell(row, 30).Value = p.PartyOfficialDate.HasValue ? DatePickerHelper.FormatDateForDisplay(p.PartyOfficialDate.Value) : "";
                    worksheet.Cell(row, 31).Value = p.CalculatedPartyAge;

                    worksheet.Cell(row, 32).Value = p.RankCode;
                    worksheet.Cell(row, 33).Value = p.RankName;
                    worksheet.Cell(row, 34).Value = "'" + p.CurrentSalaryStep; // e.g. "1/9" can be interpreted as date
                    worksheet.Cell(row, 35).Value = p.CurrentSalaryCoefficient;

                    worksheet.Cell(row, 36).Value = p.PositionAllowance;
                    worksheet.Cell(row, 37).Value = p.ExceedFramePercent > 0 ? $"{p.ExceedFramePercent}%" : "";

                    worksheet.Cell(row, 38).Value = p.EmulationTitles;
                    worksheet.Cell(row, 39).Value = p.RewardForms;
                    worksheet.Cell(row, 40).Value = p.DisciplineType == "---" ? "" : p.DisciplineType;

                    // Cột 41: Ghi chú - Theo dõi nghỉ thai sản, nghỉ ốm và chứng chỉ khác (nếu có)
                    string ghiChu = "";
                    if (p.LeaveHistories != null)
                    {
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

                    // Thêm chứng chỉ khác vào Ghi chú nếu có
                    var otherList = allDegrees.Where(d => d.DegreeType == "Chứng chỉ khác" ||
                                                          (!new[] { "Tin học", "Ngoại ngữ", "Chuyên môn", "Quản lý Nhà nước", "Lý luận chính trị" }.Contains(d.DegreeType)))
                                               .OrderByDescending(d => d.IsPrimary)
                                               .ThenBy(d => d.Id)
                                               .ToList();
                    if (otherList.Any())
                    {
                        string otherStr = string.Join(", ", otherList.Select(d => d.DegreeName));
                        if (!string.IsNullOrWhiteSpace(otherStr))
                        {
                            ghiChu = string.IsNullOrWhiteSpace(ghiChu) ? $"CC khác: {otherStr}" : $"{ghiChu}\nCC khác: {otherStr}";
                        }
                    }

                    worksheet.Cell(row, 41).Value = ghiChu;
                    worksheet.Cell(row, 42).Value = p.RetirementDate.HasValue ? DatePickerHelper.FormatDateForDisplay(p.RetirementDate.Value) : "";

                    // Áp dụng viền, căn chỉnh giữa và cho phép hiển thị nhiều dòng trong ô (WrapText)
                    for (int c = 1; c <= headers.Length; c++)
                    {
                        worksheet.Cell(row, c).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        worksheet.Cell(row, c).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(row, c).Style.Alignment.WrapText = true;
                    }

                    row++;
                }

                // Tự động điều chỉnh độ rộng cột và chiều cao dòng theo nội dung
                worksheet.Columns().AdjustToContents();
                worksheet.Rows().AdjustToContents();

                // Lưu file Excel
                workbook.SaveAs(filePath);
            }
        }

        private static List<PersonnelDegree> GetDegreesForPersonnel(Personnel p, Dictionary<int, List<PersonnelDegree>> dbDegreesMap)
        {
            if (p.Id > 0 && dbDegreesMap.TryGetValue(p.Id, out var dbList) && dbList.Any())
            {
                return dbList;
            }
            if (p.PersonnelDegrees != null && p.PersonnelDegrees.Any())
            {
                return p.PersonnelDegrees.ToList();
            }

            // Tự động kế thừa từ các trường truyền thống của cán bộ nếu chưa có bản ghi trong bảng PersonnelDegrees
            var fallback = new List<PersonnelDegree>();
            if (!string.IsNullOrWhiteSpace(p.EducationLevel))
            {
                fallback.Add(new PersonnelDegree
                {
                    PersonnelId = p.Id,
                    DegreeType = "Chuyên môn",
                    DegreeName = p.EducationLevel.Trim(),
                    Major = p.Major,
                    Institution = p.University,
                    IsPrimary = true
                });
            }
            if (!string.IsNullOrWhiteSpace(p.PoliticalTheoryLevel))
            {
                fallback.Add(new PersonnelDegree
                {
                    PersonnelId = p.Id,
                    DegreeType = "Lý luận chính trị",
                    DegreeName = p.PoliticalTheoryLevel.Trim(),
                    IsPrimary = true
                });
            }
            if (!string.IsNullOrWhiteSpace(p.StateManagementLevel))
            {
                fallback.Add(new PersonnelDegree
                {
                    PersonnelId = p.Id,
                    DegreeType = "Quản lý Nhà nước",
                    DegreeName = p.StateManagementLevel.Trim(),
                    IsPrimary = true
                });
            }
            if (!string.IsNullOrWhiteSpace(p.LanguageSkillLevel))
            {
                fallback.Add(new PersonnelDegree
                {
                    PersonnelId = p.Id,
                    DegreeType = "Ngoại ngữ",
                    DegreeName = p.LanguageSkillLevel.Trim(),
                    IsPrimary = true
                });
            }
            if (!string.IsNullOrWhiteSpace(p.ITSkillLevel))
            {
                fallback.Add(new PersonnelDegree
                {
                    PersonnelId = p.Id,
                    DegreeType = "Tin học",
                    DegreeName = p.ITSkillLevel.Trim(),
                    IsPrimary = true
                });
            }
            return fallback;
        }

        private static int GetDegreeTypeSortOrder(string? degreeType)
        {
            return (degreeType ?? "").Trim() switch
            {
                "Chuyên môn" => 1,
                "Tin học" => 2,
                "Ngoại ngữ" => 3,
                "Quản lý Nhà nước" => 4,
                "Lý luận chính trị" => 5,
                "Chứng chỉ khác" => 6,
                _ => 7
            };
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
