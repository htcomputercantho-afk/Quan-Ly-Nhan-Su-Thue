using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using TaxPersonnelManagement.Models;
using TaxPersonnelManagement.Helpers;

namespace TaxPersonnelManagement.Services
{
    public static class PlanningExcelExporter
    {
        public static void Export(IEnumerable<PlanningRecord> records, string filePath)
        {
            string planningType = records.FirstOrDefault()?.PlanningType ?? "Chuyên môn";

            using (var workbook = new XLWorkbook())
            {
                // Group 1: Chuyển tiếp (Status != "Bổ sung quy hoạch")
                var transferRecords = records.Where(r => r.Status != "Bổ sung quy hoạch").ToList();
                var sheetTransfer = workbook.Worksheets.Add("RS Chuyển tiếp");
                WriteTransferSheet(sheetTransfer, transferRecords, planningType);

                // Group 2: Bổ sung (Status == "Bổ sung quy hoạch")
                var addRecords = records.Where(r => r.Status == "Bổ sung quy hoạch").ToList();
                var sheetAdd = workbook.Worksheets.Add("RS Bổ sung");
                WriteAddSheet(sheetAdd, addRecords, planningType);

                workbook.SaveAs(filePath);
            }
        }

        private static void WriteTransferSheet(IXLWorksheet worksheet, List<PlanningRecord> records, string planningType)
        {
            worksheet.Style.Font.FontName = "Times New Roman";
            worksheet.PageSetup.ShowGridlines = true;

            // Row 1: TÊN ĐƠN VỊ ...
            worksheet.Cell(1, 1).Value = "TÊN ĐƠN VỊ ...";
            worksheet.Cell(1, 1).Style.Font.Bold = true;
            worksheet.Cell(1, 1).Style.Font.FontSize = 11;
            worksheet.Row(1).Height = 20;

            // Row 2: Title
            worksheet.Range("A2:L2").Merge().Value = "Phụ lục 1: BẢNG RÀ SOÁT, CHUYỂN TIẾP QUY HOẠCH CÁC CHỨC DANH LÃNH ĐẠO (Đợt năm 2026)";
            var titleStyle = worksheet.Range("A2:L2").Style;
            titleStyle.Font.Bold = true;
            titleStyle.Font.FontSize = 14;
            titleStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            titleStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            worksheet.Row(2).Height = 25;

            // Row 3: Subtitle 1
            worksheet.Range("A3:L3").Merge().Value = "của đơn vị ....";
            var subtitleStyle1 = worksheet.Range("A3:L3").Style;
            subtitleStyle1.Font.Bold = true;
            subtitleStyle1.Font.FontSize = 12;
            subtitleStyle1.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            subtitleStyle1.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            worksheet.Row(3).Height = 20;

            // Row 4: Subtitle 2
            worksheet.Range("A4:L4").Merge().Value = "(kèm theo Tờ trình số ..../TTr-... ngày .../.../2026 của ....)";
            var subtitleStyle2 = worksheet.Range("A4:L4").Style;
            subtitleStyle2.Font.Italic = true;
            subtitleStyle2.Font.FontSize = 11;
            subtitleStyle2.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            subtitleStyle2.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            worksheet.Row(4).Height = 20;

            // Row 5 & 6: Headers
            worksheet.Range("A5:A6").Merge().Value = "Số\nTT";
            worksheet.Range("B5:B6").Merge().Value = "Họ và tên";
            worksheet.Range("C5:D5").Merge().Value = "Ngày sinh";
            worksheet.Cell(6, 3).Value = "Nam";
            worksheet.Cell(6, 4).Value = "Nữ";
            worksheet.Range("E5:E6").Merge().Value = "Chức vụ hiện tại";
            worksheet.Range("F5:F6").Merge().Value = "Chức danh đã được quy hoạch";
            worksheet.Range("G5:G6").Merge().Value = "Chức danh chuyển tiếp quy hoạch";
            worksheet.Range("H5:I5").Merge().Value = "Trình độ đào tạo,\nbồi dưỡng";
            worksheet.Cell(6, 8).Value = "Trình độ chuyên môn";
            worksheet.Cell(6, 9).Value = "Lý luận chính trị";
            worksheet.Range("J5:J6").Merge().Value = "QĐ quy hoạch";
            worksheet.Range("K5:K6").Merge().Value = "Kết quả phân loại công chức\n03 năm gần nhất";
            worksheet.Range("L5:L6").Merge().Value = "Đề xuất của đơn vị";

            // Headers Styling
            var headerRange = worksheet.Range("A5:L6");
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Font.FontSize = 10;
            headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            headerRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            headerRange.Style.Alignment.WrapText = true;
            headerRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            headerRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            worksheet.Row(5).Height = 25;
            worksheet.Row(6).Height = 25;

            // Row 7: Index Row
            for (int col = 1; col <= 12; col++)
            {
                worksheet.Cell(7, col).Value = col;
            }
            var indexRange = worksheet.Range("A7:L7");
            indexRange.Style.Font.Italic = true;
            indexRange.Style.Font.FontSize = 9;
            indexRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            indexRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            indexRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            indexRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            worksheet.Row(7).Height = 18;

            int currentRow = 8;

            // Group by PlanningTerm
            var termGroups = records.GroupBy(r => r.PlanningTerm ?? "Chưa rõ").OrderBy(g => g.Key).ToList();
            char termLetter = 'A';

            foreach (var termGroup in termGroups)
            {
                var termHeaderRange = worksheet.Range($"A{currentRow}:L{currentRow}");
                termHeaderRange.Merge().Value = $"{termLetter}. NHIỆM KỲ {termGroup.Key}";
                termHeaderRange.Style.Font.Bold = true;
                termHeaderRange.Style.Font.FontSize = 11;
                termHeaderRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                termHeaderRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                termHeaderRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                termHeaderRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                worksheet.Row(currentRow).Height = 24;
                currentRow++;
                termLetter++;

                var transferRecords = termGroup.Where(r => r.Status == "Tiếp tục quy hoạch").ToList();
                var removeRecords = termGroup.Where(r => r.Status == "Đưa ra khỏi quy hoạch").ToList();

                List<PlanningRecord> group1List;
                List<PlanningRecord> group2List;
                string group1Title;
                string group2Title;
                string group3Title;

                if (planningType == "Đảng")
                {
                    group1List = transferRecords.Where(r => IsPartyBithuOrEquivalent(r.PlannedPosition ?? "")).ToList();
                    group2List = transferRecords.Where(r => !IsPartyBithuOrEquivalent(r.PlannedPosition ?? "")).ToList();
                    group1Title = "Danh sách chuyển tiếp quy hoạch sang chức danh Bí thư và tương đương";
                    group2Title = "Danh sách chuyển tiếp quy hoạch sang chức danh Phó Bí thư và tương đương";
                    group3Title = "Danh sách đưa ra khỏi quy hoạch chức danh Bí thư, Phó Bí thư và tương đương";
                }
                else
                {
                    group1List = transferRecords.Where(r => IsTruongPhongOrEquivalent(r.PlannedPosition ?? "")).ToList();
                    group2List = transferRecords.Where(r => !IsTruongPhongOrEquivalent(r.PlannedPosition ?? "")).ToList();
                    group1Title = "Danh sách chuyển tiếp quy hoạch sang chức danh Trưởng phòng và tương đương";
                    group2Title = "Danh sách chuyển tiếp quy hoạch sang chức danh Phó Trưởng phòng và tương đương";
                    group3Title = "Danh sách đưa ra khỏi quy hoạch chức danh Trưởng phòng, Phó Trưởng phòng và tương đương";
                }

                if (group1List.Any())
                {
                    currentRow = WriteTransferGroupHeader(worksheet, "I", group1Title, currentRow);
                    currentRow = WriteTransferDataRows(worksheet, group1List, currentRow);
                }

                if (group2List.Any())
                {
                    currentRow = WriteTransferGroupHeader(worksheet, "II", group2Title, currentRow);
                    currentRow = WriteTransferDataRows(worksheet, group2List, currentRow);
                }

                if (removeRecords.Any())
                {
                    currentRow = WriteTransferGroupHeader(worksheet, "III", group3Title, currentRow);
                    currentRow = WriteTransferDataRows(worksheet, removeRecords, currentRow);
                }
            }

            // Auto fit column widths
            worksheet.Columns().AdjustToContents(5, currentRow - 1);
            worksheet.Column(1).Width = 6;  // Số TT
            worksheet.Column(2).Width = 25; // Họ và tên
            worksheet.Column(3).Width = 12; // Nam
            worksheet.Column(4).Width = 12; // Nữ
            worksheet.Column(5).Width = 20; // Chức vụ hiện tại
            worksheet.Column(6).Width = 22; // Chức danh đã được quy hoạch
            worksheet.Column(7).Width = 22; // Chức danh chuyển tiếp
            worksheet.Column(8).Width = 18; // Trình độ chuyên môn
            worksheet.Column(9).Width = 15; // Lý luận chính trị
            worksheet.Column(10).Width = 22; // QĐ quy hoạch
            worksheet.Column(11).Width = 26; // Kết quả phân loại
            worksheet.Column(12).Width = 24; // Đề xuất
        }

        private static int WriteTransferGroupHeader(IXLWorksheet worksheet, string romanNumeral, string title, int currentRow)
        {
            var range = worksheet.Range($"A{currentRow}:L{currentRow}");
            range.Merge().Value = $"{romanNumeral} | {title}";
            range.Style.Font.Bold = true;
            range.Style.Font.FontSize = 10;
            range.Style.Font.FontColor = XLColor.FromHtml("#6B21A8"); // Purple
            range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            worksheet.Row(currentRow).Height = 24;
            return currentRow + 1;
        }

        private static int WriteTransferDataRows(IXLWorksheet worksheet, List<PlanningRecord> records, int currentRow)
        {
            int stt = 1;
            foreach (var record in records)
            {
                var p = record.Personnel;

                worksheet.Cell(currentRow, 1).Value = stt++; // Số TT
                worksheet.Cell(currentRow, 2).Value = p?.FullName ?? ""; // Họ và tên
                
                if (p != null)
                {
                    string dobStr = p.DateOfBirth.HasValue ? p.DateOfBirth.Value.ToString("dd/M/yyyy") : "";
                    if (p.Gender == "Nam")
                    {
                        worksheet.Cell(currentRow, 3).Value = dobStr;
                        worksheet.Cell(currentRow, 4).Value = "";
                    }
                    else
                    {
                        worksheet.Cell(currentRow, 3).Value = "";
                        worksheet.Cell(currentRow, 4).Value = dobStr;
                    }
                }
                else
                {
                    worksheet.Cell(currentRow, 3).Value = "";
                    worksheet.Cell(currentRow, 4).Value = "";
                }

                worksheet.Cell(currentRow, 5).Value = !string.IsNullOrEmpty(record.CurrentPosition) ? record.CurrentPosition : (p?.Position ?? "");
                worksheet.Cell(currentRow, 6).Value = record.PlannedPosition ?? "";
                worksheet.Cell(currentRow, 7).Value = record.PlannedTransitionPosition ?? "";
                worksheet.Cell(currentRow, 8).Value = !string.IsNullOrEmpty(record.TrainingLevel) ? record.TrainingLevel : (p?.EducationLevel ?? "");
                worksheet.Cell(currentRow, 9).Value = !string.IsNullOrEmpty(record.PoliticalTheoryLevel) ? record.PoliticalTheoryLevel : (p?.PoliticalTheoryLevel ?? "");
                
                if (!string.IsNullOrEmpty(record.DecisionNumber) || record.DecisionDate.HasValue)
                {
                    string decNum = !string.IsNullOrEmpty(record.DecisionNumber) ? record.DecisionNumber : "...";
                    string decDate = record.DecisionDate.HasValue ? record.DecisionDate.Value.ToString("dd/M/yyyy") : ".../../....";
                    worksheet.Cell(currentRow, 10).Value = $"Quyết định số {decNum} ngày {decDate}";
                }
                else
                {
                    worksheet.Cell(currentRow, 10).Value = "---";
                }

                worksheet.Cell(currentRow, 11).Value = FormatEvaluationForExcel(record.Evaluation3Years);
                worksheet.Cell(currentRow, 12).Value = record.Note ?? "";

                var rowRange = worksheet.Range($"A{currentRow}:L{currentRow}");
                rowRange.Style.Font.FontSize = 10;
                rowRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                rowRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                rowRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                rowRange.Style.Alignment.WrapText = true;

                worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell(currentRow, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell(currentRow, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell(currentRow, 10).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                int lineCount = 1;
                if (!string.IsNullOrEmpty(record.Evaluation3Years))
                {
                    lineCount = Math.Max(lineCount, record.Evaluation3Years.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None).Length);
                }
                if (!string.IsNullOrEmpty(record.Note))
                {
                    lineCount = Math.Max(lineCount, record.Note.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None).Length);
                }
                worksheet.Row(currentRow).Height = Math.Max(28, 15 * lineCount + 10);

                currentRow++;
            }
            return currentRow;
        }

        private static void WriteAddSheet(IXLWorksheet worksheet, List<PlanningRecord> records, string planningType)
        {
            worksheet.Style.Font.FontName = "Times New Roman";
            worksheet.PageSetup.ShowGridlines = true;

            // Row 1: TÊN ĐƠN VỊ ...
            worksheet.Cell(1, 1).Value = "TÊN ĐƠN VỊ ...";
            worksheet.Cell(1, 1).Style.Font.Bold = true;
            worksheet.Cell(1, 1).Style.Font.FontSize = 11;
            worksheet.Row(1).Height = 20;

            // Row 2: Title
            worksheet.Range("A2:J2").Merge().Value = "Phụ lục 2: BẢNG RÀ SOÁT, BỔ SUNG QUY HOẠCH CÁC CHỨC DANH LÃHO ĐẠO (Đợt năm 2026)";
            var titleStyle = worksheet.Range("A2:J2").Style;
            titleStyle.Font.Bold = true;
            titleStyle.Font.FontSize = 14;
            titleStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            titleStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            worksheet.Row(2).Height = 25;

            // Row 3: Subtitle 1
            worksheet.Range("A3:J3").Merge().Value = "của đơn vị ....";
            var subtitleStyle1 = worksheet.Range("A3:J3").Style;
            subtitleStyle1.Font.Bold = true;
            subtitleStyle1.Font.FontSize = 12;
            subtitleStyle1.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            subtitleStyle1.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            worksheet.Row(3).Height = 20;

            // Row 4: Subtitle 2
            worksheet.Range("A4:J4").Merge().Value = "(kèm theo Tờ trình số ..../TTr-... ngày .../.../2026 của ....)";
            var subtitleStyle2 = worksheet.Range("A4:J4").Style;
            subtitleStyle2.Font.Italic = true;
            subtitleStyle2.Font.FontSize = 11;
            subtitleStyle2.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            subtitleStyle2.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            worksheet.Row(4).Height = 20;

            // Row 5 & 6: Headers
            worksheet.Range("A5:A6").Merge().Value = "Số\nTT";
            worksheet.Range("B5:B6").Merge().Value = "Họ và tên";
            worksheet.Range("C5:D5").Merge().Value = "Ngày sinh";
            worksheet.Cell(6, 3).Value = "Nam";
            worksheet.Cell(6, 4).Value = "Nữ";
            worksheet.Range("E5:E6").Merge().Value = "Chức vụ hiện tại";
            worksheet.Range("F5:F6").Merge().Value = "Chức danh đề xuất quy hoạch";
            worksheet.Range("G5:H5").Merge().Value = "Trình độ đào tạo,\nbồi dưỡng";
            worksheet.Cell(6, 7).Value = "Trình độ chuyên môn";
            worksheet.Cell(6, 8).Value = "Lý luận chính trị";
            worksheet.Range("I5:I6").Merge().Value = "Kết quả phân loại công chức\n03 năm gần nhất";
            worksheet.Range("J5:J6").Merge().Value = "Đề xuất của đơn vị";

            var headerRange = worksheet.Range("A5:J6");
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Font.FontSize = 10;
            headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            headerRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            headerRange.Style.Alignment.WrapText = true;
            headerRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            headerRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            worksheet.Row(5).Height = 25;
            worksheet.Row(6).Height = 25;

            // Row 7: Index Row
            for (int col = 1; col <= 10; col++)
            {
                worksheet.Cell(7, col).Value = col;
            }
            var indexRange = worksheet.Range("A7:J7");
            indexRange.Style.Font.Italic = true;
            indexRange.Style.Font.FontSize = 9;
            indexRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            indexRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            indexRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            indexRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            worksheet.Row(7).Height = 18;

            int currentRow = 8;

            // Group by PlanningTerm
            var termGroups = records.GroupBy(r => r.PlanningTerm ?? "Chưa rõ").OrderBy(g => g.Key).ToList();
            char termLetter = 'A';

            foreach (var termGroup in termGroups)
            {
                var termHeaderRange = worksheet.Range($"A{currentRow}:J{currentRow}");
                termHeaderRange.Merge().Value = $"{termLetter}. NHIỆM KỲ {termGroup.Key}";
                termHeaderRange.Style.Font.Bold = true;
                termHeaderRange.Style.Font.FontSize = 11;
                termHeaderRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                termHeaderRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                termHeaderRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                termHeaderRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                worksheet.Row(currentRow).Height = 24;
                currentRow++;
                termLetter++;

                var addRecords = termGroup.ToList();
                List<PlanningRecord> group1List;
                List<PlanningRecord> group2List;
                string group1Title;
                string group2Title;

                if (planningType == "Đảng")
                {
                    group1List = addRecords.Where(r => IsPartyBithuOrEquivalent(r.PlannedPosition ?? "")).ToList();
                    group2List = addRecords.Where(r => !IsPartyBithuOrEquivalent(r.PlannedPosition ?? "")).ToList();
                    group1Title = "Bổ sung quy hoạch Bí thư (hoặc tương đương):";
                    group2Title = "Bổ sung quy hoạch Phó Bí thư (hoặc tương đương):";
                }
                else
                {
                    group1List = addRecords.Where(r => IsTruongPhongOrEquivalent(r.PlannedPosition ?? "")).ToList();
                    group2List = addRecords.Where(r => !IsTruongPhongOrEquivalent(r.PlannedPosition ?? "")).ToList();
                    group1Title = "Bổ sung quy hoạch Trưởng phòng (hoặc tương đương):";
                    group2Title = "Bổ sung quy hoạch Phó Trưởng phòng (hoặc tương đương):";
                }

                if (group1List.Any())
                {
                    currentRow = WriteAddGroupHeader(worksheet, "I", group1Title, currentRow);
                    currentRow = WriteAddSpecialSummaryRows(worksheet, group1List.Count, currentRow);
                    currentRow = WriteAddDataRows(worksheet, group1List, currentRow);
                }

                if (group2List.Any())
                {
                    currentRow = WriteAddGroupHeader(worksheet, "II", group2Title, currentRow);
                    currentRow = WriteAddSpecialSummaryRows(worksheet, group2List.Count, currentRow);
                    currentRow = WriteAddDataRows(worksheet, group2List, currentRow);
                }
            }

            worksheet.Columns().AdjustToContents(5, currentRow - 1);
            worksheet.Column(1).Width = 6;  // Số TT
            worksheet.Column(2).Width = 25; // Họ và tên
            worksheet.Column(3).Width = 12; // Nam
            worksheet.Column(4).Width = 12; // Nữ
            worksheet.Column(5).Width = 20; // Chức vụ hiện tại
            worksheet.Column(6).Width = 22; // Chức danh đề xuất
            worksheet.Column(7).Width = 18; // Trình độ chuyên môn
            worksheet.Column(8).Width = 15; // Lý luận chính trị
            worksheet.Column(9).Width = 26; // Kết quả phân loại
            worksheet.Column(10).Width = 24; // Đề xuất
        }

        private static int WriteAddGroupHeader(IXLWorksheet worksheet, string romanNumeral, string title, int currentRow)
        {
            var range = worksheet.Range($"A{currentRow}:J{currentRow}");
            range.Merge().Value = $"{romanNumeral} | {title}";
            range.Style.Font.Bold = true;
            range.Style.Font.FontSize = 10;
            range.Style.Font.FontColor = XLColor.FromHtml("#6B21A8"); // Purple
            range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            worksheet.Row(currentRow).Height = 24;
            return currentRow + 1;
        }

        private static int WriteAddSpecialSummaryRows(IXLWorksheet worksheet, int count, int currentRow)
        {
            worksheet.Cell(currentRow, 1).Value = "-";
            worksheet.Cell(currentRow, 2).Value = "Số lượng đề nghị";
            worksheet.Cell(currentRow, 10).Value = count;
            
            worksheet.Cell(currentRow, 10).Style.Font.Underline = XLFontUnderlineValues.Single;
            worksheet.Cell(currentRow, 10).Style.Font.Bold = true;
            
            var rowRange1 = worksheet.Range($"A{currentRow}:J{currentRow}");
            rowRange1.Style.Font.FontSize = 10;
            rowRange1.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            rowRange1.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            rowRange1.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            worksheet.Cell(currentRow, 10).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            worksheet.Row(currentRow).Height = 22;
            currentRow++;

            worksheet.Cell(currentRow, 1).Value = "-";
            worksheet.Cell(currentRow, 2).Value = "Danh sách nhân sự giới thiệu:";
            worksheet.Cell(currentRow, 10).Value = count;

            worksheet.Cell(currentRow, 10).Style.Font.Underline = XLFontUnderlineValues.Single;
            worksheet.Cell(currentRow, 10).Style.Font.Bold = true;

            var rowRange2 = worksheet.Range($"A{currentRow}:J{currentRow}");
            rowRange2.Style.Font.FontSize = 10;
            rowRange2.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            rowRange2.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            rowRange2.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            worksheet.Cell(currentRow, 10).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            worksheet.Row(currentRow).Height = 22;
            
            return currentRow + 1;
        }

        private static int WriteAddDataRows(IXLWorksheet worksheet, List<PlanningRecord> records, int currentRow)
        {
            int stt = 1;
            foreach (var record in records)
            {
                var p = record.Personnel;

                worksheet.Cell(currentRow, 1).Value = stt++; // Số TT
                worksheet.Cell(currentRow, 2).Value = p?.FullName ?? ""; // Họ và tên

                if (p != null)
                {
                    string dobStr = p.DateOfBirth.HasValue ? p.DateOfBirth.Value.ToString("dd/M/yyyy") : "";
                    if (p.Gender == "Nam")
                    {
                        worksheet.Cell(currentRow, 3).Value = dobStr;
                        worksheet.Cell(currentRow, 4).Value = "";
                    }
                    else
                    {
                        worksheet.Cell(currentRow, 3).Value = "";
                        worksheet.Cell(currentRow, 4).Value = dobStr;
                    }
                }
                else
                {
                    worksheet.Cell(currentRow, 3).Value = "";
                    worksheet.Cell(currentRow, 4).Value = "";
                }

                worksheet.Cell(currentRow, 5).Value = !string.IsNullOrEmpty(record.CurrentPosition) ? record.CurrentPosition : (p?.Position ?? "");
                worksheet.Cell(currentRow, 6).Value = record.PlannedPosition ?? "";
                worksheet.Cell(currentRow, 7).Value = !string.IsNullOrEmpty(record.TrainingLevel) ? record.TrainingLevel : (p?.EducationLevel ?? "");
                worksheet.Cell(currentRow, 8).Value = !string.IsNullOrEmpty(record.PoliticalTheoryLevel) ? record.PoliticalTheoryLevel : (p?.PoliticalTheoryLevel ?? "");
                worksheet.Cell(currentRow, 9).Value = FormatEvaluationForExcel(record.Evaluation3Years);
                worksheet.Cell(currentRow, 10).Value = record.Note ?? "";

                var rowRange = worksheet.Range($"A{currentRow}:J{currentRow}");
                rowRange.Style.Font.FontSize = 10;
                rowRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                rowRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                rowRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                rowRange.Style.Alignment.WrapText = true;

                worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell(currentRow, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell(currentRow, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                int lineCount = 1;
                if (!string.IsNullOrEmpty(record.Evaluation3Years))
                {
                    lineCount = Math.Max(lineCount, record.Evaluation3Years.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None).Length);
                }
                if (!string.IsNullOrEmpty(record.Note))
                {
                    lineCount = Math.Max(lineCount, record.Note.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None).Length);
                }
                worksheet.Row(currentRow).Height = Math.Max(28, 15 * lineCount + 10);

                currentRow++;
            }
            return currentRow;
        }

        private static bool IsTruongPhongOrEquivalent(string position)
        {
            if (string.IsNullOrEmpty(position)) return true;
            string p = position.ToLower();
            if (p.Contains("phó")) return false;
            if (p.Contains("trưởng") || p.Contains("chánh") || p.Contains("đội trưởng") || p.Contains("chi cục trưởng") || p.Contains("tổ trưởng")) return true;
            return true;
        }

        private static bool IsPartyBithuOrEquivalent(string position)
        {
            if (string.IsNullOrEmpty(position)) return true;
            string p = position.ToLower();
            if (p.Contains("phó")) return false;
            if (p.Contains("bí thư")) return true;
            if (p.Contains("chi ủy viên") || p.Contains("ủy viên")) return false;
            return true;
        }

        private static string FormatEvaluationForExcel(string? evalStr)
        {
            if (string.IsNullOrEmpty(evalStr)) return "";
            var lines = evalStr.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            var formattedLines = lines.Select(line =>
            {
                var parts = line.Split(':');
                if (parts.Length == 2)
                {
                    string year = parts[0].Trim();
                    string rating = parts[1].Trim();
                    return $"{year}: {GetRatingShortCode(rating)}";
                }
                return line;
            });
            return string.Join(Environment.NewLine, formattedLines);
        }

        private static string GetRatingShortCode(string rating)
        {
            if (string.IsNullOrEmpty(rating)) return "Chưa đánh giá";
            string r = rating.ToLower();
            if (r.Contains("xuất sắc")) return "HTXSNV";
            if (r.Contains("hoàn thành tốt") || r.Contains("tốt")) return "HTTNV";
            if (r.Contains("không hoàn thành")) return "KHTNV";
            if (r.Contains("hoàn thành nhiệm vụ") || r.Contains("hoàn thành")) return "HTNV";
            return rating;
        }
    }
}
