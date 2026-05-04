using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using TaxPersonnelManagement.Models;

namespace TaxPersonnelManagement.Services
{
    public static class ExcelImporter
    {
        public static List<Personnel> Import(string filePath)
        {
            var list = new List<Personnel>();

            // Mở file bằng FileStream với quyền Share ReadWrite để tránh lỗi khi file đang mở trong Excel
            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var workbook = new XLWorkbook(stream))
            {
                var worksheet = workbook.Worksheet(1);
                var range = worksheet.RangeUsed();
                if (range == null) return list;
                var rows = range.RowsUsed().Skip(1); // Bỏ qua dòng tiêu đề

                foreach (var row in rows)
                {
                    try
                    {
                        var p = new Personnel();
                        
                        // Ánh xạ theo cấu trúc file DS.xlsx (Bắt đầu từ cột B là STT)
                        p.StaffId = GetFormattedString(row.Cell(3));         // Cột C: Mã NV
                        p.FullName = GetValue(row.Cell(4)) ?? "";  // Cột D: Tên NV
                        p.Department = GetValue(row.Cell(5));      // Cột E: Bộ phận
                        
                        // Cột F (6): Thời gian công tác tại cơ quan Thuế
                        p.TaxAuthorityStartDate = GetDate(row.Cell(6));
                        
                        p.RankCode = GetValue(row.Cell(7));        // Cột G: Mã ngạch
                        p.RankName = GetValue(row.Cell(8));        // Cột H: Ngạch công chức
                        p.Position = GetValue(row.Cell(9));        // Cột I: Chức vụ
                        p.Gender = GetValue(row.Cell(10));         // Cột J: Giới tính
                        p.DateOfBirth = GetDate(row.Cell(11));     // Cột K: Ngày sinh
                        
                        // Cột L (12): Hôn nhân (Bỏ qua)
                        
                        p.Ethnicity = GetValue(row.Cell(13));      // Cột M: Dân tộc
                        p.Religion = GetValue(row.Cell(14));       // Cột N: Tôn giáo
                        p.IdentityCardNumber = GetFormattedString(row.Cell(15)); // Cột O: CCCD
                        
                        // Nơi cấp CCCD: Luôn để mặc định theo yêu cầu
                        p.IdentityCardPlace = "Cục Cảnh sát quản lý hành chính về trật tự xã hội";
                        
                        p.SocialSecurityNumber = GetFormattedString(row.Cell(16)); // Cột P: BHXH
                        
                        p.EducationLevel = GetValue(row.Cell(17)); // Cột Q: Trình độ
                        p.Major = GetValue(row.Cell(18));          // Cột R: Chuyên ngành
                        p.University = GetValue(row.Cell(19));     // Cột S: Trường đào tạo
                        
                        p.Email = GetValue(row.Cell(20));          // Cột T: Email
                        p.PhoneNumber = GetFormattedString(row.Cell(21));    // Cột U: SĐT
                        p.BirthPlace = GetValue(row.Cell(22));     // Cột V: Nơi đăng ký khai sinh
                        
                        // Cột W (23): Đảng viên
                        
                        p.PartyEntryDate = GetDate(row.Cell(24));  // Cột X: Ngày vào Đảng
                        p.PartyOfficialDate = GetDate(row.Cell(25)); // Cột Y: Ngày chính thức
                        
                        // Cột Z (26): Số thẻ đảng (Bỏ qua)
                        
                        p.StateManagementLevel = GetValue(row.Cell(27)); // Cột AA: QL Nhà nước
                        p.PoliticalTheoryLevel = GetValue(row.Cell(28)); // Cột AB: Lý luận CT
                        p.ITSkillLevel = GetValue(row.Cell(29));         // Cột AC: Tin học
                        p.LanguageSkillLevel = GetValue(row.Cell(30));    // Cột AD: Ngoại ngữ
                        
                        // Cột AE (31): Trạng thái TV (Bỏ qua)
                        // Cột AF (32): Tỉnh/TP (Bỏ qua)
                        
                        p.RetirementDate = GetDate(row.Cell(33));  // Cột AG: Ngày nghỉ hưu
                        p.PositionAllowance = GetValue(row.Cell(34)); // Cột AH: PC Chức vụ

                        // Cột AI (35): Lý do nghỉ việc (Bỏ qua)
                        
                        // Cột AJ (36): Thời gian công tác theo QĐ gần nhất
                        p.PositionDecisionDate = GetDate(row.Cell(36));

                        if (!string.IsNullOrWhiteSpace(p.FullName))
                        {
                            list.Add(p);
                        }
                    }
                    catch (Exception)
                    {
                        // Bỏ qua dòng lỗi
                    }
                }
            }

            return list;
        }

        private static string? GetValue(IXLCell cell)
        {
            var val = cell.Value.ToString();
            if (string.IsNullOrWhiteSpace(val) || val == "0") return null;
            return val.Trim();
        }

        private static string? GetFormattedString(IXLCell cell)
        {
            var val = GetValue(cell);
            if (string.IsNullOrEmpty(val)) return null;

            // Nếu là SĐT và không bắt đầu bằng số 0, nhưng có 9 chữ số (thiếu số 0 đầu)
            if (val.Length == 9 && char.IsDigit(val[0]) && !val.StartsWith("0"))
            {
                return "0" + val;
            }
            
            return val;
        }

        private static DateTime? GetDate(IXLCell cell)
        {
            if (cell.DataType == XLDataType.DateTime) return cell.GetDateTime();
            
            string val = cell.Value.ToString().Trim();
            if (string.IsNullOrEmpty(val)) return null;

            if (DateTime.TryParse(val, out DateTime dt)) return dt;
            
            return null;
        }
    }
}
