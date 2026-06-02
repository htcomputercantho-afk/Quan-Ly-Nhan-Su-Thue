using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using System;
using System.IO;
using TaxPersonnelManagement.Models;
using TaxPersonnelManagement.Helpers;
using Colors = QuestPDF.Helpers.Colors;

namespace TaxPersonnelManagement.Services
{
    public static class PdfExporter
    {
        private static string FormatDate(DateTime? dt, string fallback = "")
        {
            return dt.HasValue ? DatePickerHelper.FormatDateForDisplay(dt.Value) : fallback;
        }
        public static void Export(Personnel p, string filePath)
        {
            QuestPDF.Settings.License = LicenseType.Community;

            Document.Create(container =>
            {
                container.Page(page =>
                {
                    page.Size(PageSizes.A4);
                    page.Margin(0); // Full Bleed
                    page.PageColor(Colors.White);
                    page.DefaultTextStyle(x => x.FontSize(11).FontFamily(Fonts.Arial));

                    // Page Content with Red Border and Distinct Backgrounds
                    page.Content()
                        .PaddingVertical(0)
                        .Border(5).BorderColor("#B71C1C") // Full Page Red Frame
                        .Padding(20) // More formatting padding since we have no margin
                        .Column(column =>
                        {
                            column.Spacing(15);

                            // Header on White Background
                            column.Item().Element(header => ComposeHeader(header, p));

                            // Body sections on Light Yellow Background
                            column.Item().Background("#FFF8E1").Border(1).BorderColor("#FFE0B2").Padding(10).Column(bodyCol =>
                            {
                                bodyCol.Spacing(15);
                                bodyCol.Item().Element(e => ComposeSection(e, "1. THÔNG TIN CHUNG", c => ComposeGeneralInfo(c, p)));
                                bodyCol.Item().Element(e => ComposeSection(e, "2. ĐÀO TẠO", c => ComposeEducation(c, p)));
                                bodyCol.Item().Element(e => ComposeSection(e, "3. THÔNG TIN LƯƠNG", c => ComposeSalary(c, p)));
                                bodyCol.Item().Element(e => ComposeSection(e, "4. THÔNG TIN NGHỈ PHÉP", c => ComposeLeave(c, p)));
                                bodyCol.Item().Element(e => ComposeSection(e, "5. CÔNG TÁC", c => ComposeWorkHistory(c, p)));
                                bodyCol.Item().Element(e => ComposeSection(e, "6. THÔNG TIN ĐẢNG VIÊN", c => ComposeParty(c, p)));
                                bodyCol.Item().Element(e => ComposeSection(e, "7. KHEN THƯỞNG & KỶ LUẬT", c => ComposeRewardDiscipline(c, p)));
                                bodyCol.Item().Element(e => ComposeSection(e, "8. LỊCH SỬ XẾP LOẠI", c => ComposeEvaluationHistory(c, p), "#E65100", "#FFE0B2"));
                            });

                            // Footer inside the border
                            column.Item().PaddingTop(10).AlignCenter().Text(x =>
                            {
                                x.Span("Hệ thống Quản lý Nhân sự - Xuất ngày " + DatePickerHelper.FormatDateForDisplay(DateTime.Now))
                                 .FontColor(Colors.Grey.Medium).FontSize(10).Italic();
                            });
                        });
                });
            })
            .GeneratePdf(filePath);
        }

        static void ComposeHeader(IContainer container, Personnel p)
        {
            container.Border(1).BorderColor(Colors.Grey.Lighten2).Padding(15).Row(row =>
            {
                // Avatar
                var avatarItem = row.ConstantItem(100).Height(120).Border(1).BorderColor(Colors.Grey.Lighten2).Background(Colors.White);

                bool hasImage = false;
                if (!string.IsNullOrEmpty(p.AvatarBase64))
                {
                    try
                    {
                        var base64 = p.AvatarBase64;
                        if (base64.Contains(",")) base64 = base64.Split(',')[1];
                        byte[] imageBytes = Convert.FromBase64String(base64);
                        avatarItem.Image(imageBytes).FitUnproportionally();
                        hasImage = true;
                    }
                    catch { /* Ignore error, show placeholder */ }
                }

                if (!hasImage)
                {
                    avatarItem.AlignCenter().AlignMiddle().Text("Ảnh").FontColor(Colors.Grey.Medium);
                }

                row.RelativeItem().PaddingLeft(20).Column(col =>
                {
                    col.Item().Text(p.FullName?.ToUpper() ?? "").FontSize(20).Bold().FontColor("#D32F2F"); // Red
                    col.Item().Text(p.Department?.ToUpper() ?? "").FontSize(14).Bold().FontColor("#B71C1C");

                    col.Item().PaddingTop(5).Text(p.Position ?? "").FontSize(14).Bold();

                    col.Item().PaddingTop(10).Row(r =>
                    {
                        r.AutoItem()
                         .Background("#9C2727") // Deep Red
                         .CornerRadius(15)      // Pill Shape
                         .PaddingHorizontal(15)
                         .PaddingVertical(5)
                         .Text($"Số hiệu công chức: {p.StaffId}")
                         .FontColor(Colors.White)
                         .Bold();
                    });
                });
            });
        }

        static void ComposeSection(IContainer container, string title, Action<ColumnDescriptor> content, string sectionColor = "#B71C1C", string borderColor = "#FBE9E7")
        {
            // Ensure section start has enough space, or push start to next page
            container.EnsureSpace(100)
                     .Border(1).BorderColor(borderColor).Background("#FFFAFA").Padding(15).Column(column =>
            {
                column.Item().Text(title).FontSize(14).Bold().FontColor(sectionColor);
                column.Item().PaddingTop(5).LineHorizontal(1).LineColor(Colors.Grey.Lighten2);
                column.Item().PaddingTop(10).Element(e => e.Column(content));
            });
        }

        static void ComposeGeneralInfo(ColumnDescriptor column, Personnel p)
        {
            column.Item().Row(row =>
            {
                row.RelativeItem().Column(c =>
                {
                    LabelValue(c, "Ngày sinh:", FormatDate(p.DateOfBirth));
                    LabelValue(c, "SĐT:", p.PhoneNumber);
                    LabelValue(c, "CCCD:", p.IdentityCardNumber);
                    LabelValue(c, "BHXH:", p.SocialSecurityNumber);
                });
                row.RelativeItem().Column(c =>
                {
                    LabelValue(c, "Giới tính:", p.Gender);
                    LabelValue(c, "Email:", p.Email);
                    LabelValue(c, "Nơi cấp:", p.IdentityCardPlace);
                    LabelValue(c, "Nơi sinh:", p.BirthPlace);
                });
            });
        }

        static void ComposeEducation(ColumnDescriptor column, Personnel p)
        {
            column.Item().Row(row =>
            {
                row.RelativeItem().Column(c =>
                {
                    LabelValue(c, "Trình độ:", p.EducationLevel);
                    LabelValue(c, "Tin học:", p.ITSkillLevel);
                    LabelValue(c, "Quản lý Nhà nước:", p.StateManagementLevel);
                });
                row.RelativeItem().Column(c =>
                {
                    LabelValue(c, "Chuyên ngành:", $"{p.Major} - {p.University}");
                    LabelValue(c, "Ngoại ngữ:", p.LanguageSkillLevel);
                    LabelValue(c, "Lý luận Chính trị:", p.PoliticalTheoryLevel);
                });
            });
        }

        static void ComposeSalary(ColumnDescriptor column, Personnel p)
        {
            // 1. Calculate Salary Dates matching UI logic
            DateTime? nextCalcDate = p.NextSalaryStepDate;

            // If p.NextSalaryStepDate is null, try to get it from the latest SalaryRecord
            if (!nextCalcDate.HasValue && p.SalaryRecords != null && p.SalaryRecords.Count > 0)
            {
                nextCalcDate = p.SalaryRecords.OrderByDescending(s => s.StartDate).FirstOrDefault()?.SalaryCalculationDate;
            }

            DateTime? expectedDate = null;
            if (nextCalcDate.HasValue)
            {
                DateTime baseDate = nextCalcDate.Value;
                int periodYears = 3;

                // Check Exceed Frame
                if (p.ExceedFramePercent > 0)
                {
                    periodYears = 1;
                }
                else
                {
                    // Check Rank Code
                    string rc = p.RankCode?.Trim() ?? "";
                    if (rc == "06.039-1" || rc == "01.011" || rc == "01.009")
                    {
                        periodYears = 2;
                    }
                }

                DateTime calcDate = baseDate.AddYears(periodYears);

                // Delay from Disciplinary Action
                string delayType = p.SalaryIncreaseDelayType ?? "";
                if (delayType.Contains("3 tháng")) calcDate = calcDate.AddMonths(3);
                else if (delayType.Contains("6 tháng")) calcDate = calcDate.AddMonths(6);
                else if (delayType.Contains("12 tháng")) calcDate = calcDate.AddMonths(12);

                // Delay from Unpaid Leave
                if (delayType != "-- Không lùi --" && p.LeaveHistories != null)
                {
                    double unpaidDays = p.LeaveHistories
                        .Where(h => h.LeaveType == "Không lương")
                        .Sum(h => h.DurationDays);

                    if (unpaidDays > 0)
                    {
                        calcDate = calcDate.AddDays(unpaidDays);
                    }
                }

                expectedDate = calcDate;
            }

            column.Item().PaddingTop(5).Background("#FAFAFA").CornerRadius(5).Padding(15).Row(row =>
            {
                void CenteredItem(RowDescriptor r, string label, string? val, bool highlight = false, bool last = false, int fontSize = 16)
                {
                    r.RelativeItem().BorderRight(last ? 0 : 2).BorderColor(Colors.Grey.Lighten2).PaddingHorizontal(10).Column(c =>
                    {
                        c.Item().AlignCenter().Text(label).FontSize(10).FontColor(Colors.Grey.Darken1);
                        c.Item().PaddingTop(5).AlignCenter().Text(val ?? "---").Bold().FontSize(fontSize).FontColor(highlight ? "#B71C1C" : Colors.Black);
                    });
                }

                int codeFontSize = (!string.IsNullOrEmpty(p.RankName)) ? 11 : 16;
                CenteredItem(row, "Mã ngạch", p.RankDisplayName, last: false, fontSize: codeFontSize);
                CenteredItem(row, "Bậc lương", p.CurrentSalaryStep);
                CenteredItem(row, "Hệ số", p.CurrentSalaryCoefficient.ToString("F2"), true, true);
            });

            column.Item().PaddingTop(10).Row(row =>
            {
                row.RelativeItem().Column(c => LabelValue(c, "% Vượt khung:", p.ExceedFramePercent + "%"));
                row.RelativeItem().Column(c => LabelValue(c, "PC Chức vụ:", p.PositionAllowance));
                row.RelativeItem().Column(c => LabelValue(c, "Thời hạn bảo lưu:", FormatDate(p.SalaryReservationDeadline)));
            });

            column.Item().PaddingTop(5).Row(row =>
            {
                row.RelativeItem().Column(c => LabelValue(c, "Mốc lương:", FormatDate(nextCalcDate)));
                row.RelativeItem().Column(c => LabelValueRed(c, "Lùi thời gian nâng lương:", p.SalaryIncreaseDelayType));
                row.RelativeItem().Column(c => LabelValue(c, "Dự kiến lên lương:", FormatDate(expectedDate)));
            });

            // Salary History Table - Keep title + table together to prevent page break between them
            column.Item().PaddingTop(15).ShowEntire().Column(salarySection =>
            {
                salarySection.Item().PaddingBottom(8).Text("DIỄN BIẼN QUÁ TRÌNH LƯƠNG").Bold().FontColor("#B71C1C").FontSize(12);
                salarySection.Item().Table(table =>
                {
                    table.ColumnsDefinition(columns =>
                    {
                        columns.RelativeColumn(1.0f); // RankCode
                        columns.RelativeColumn(1.0f); // SalaryStep
                        columns.ConstantColumn(35);   // Coeff
                        columns.ConstantColumn(30);   // %
                        columns.RelativeColumn(1.2f); // Start
                        columns.RelativeColumn(1.2f); // End
                        columns.RelativeColumn(1.2f); // Calc
                        columns.RelativeColumn(1.5f); // Decision No
                        columns.RelativeColumn(1.2f); // Decision Date
                    });

                    table.Header(header =>
                    {
                        header.Cell().Background("#B71C1C").Padding(5).AlignCenter().Text("Mã ngạch").Bold().FontSize(8).FontColor(Colors.White);
                        header.Cell().Background("#B71C1C").Padding(5).AlignCenter().Text("Bậc lương").Bold().FontSize(8).FontColor(Colors.White);
                        header.Cell().Background("#B71C1C").Padding(5).AlignCenter().Text("Hệ số").Bold().FontSize(8).FontColor(Colors.White);
                        header.Cell().Background("#B71C1C").Padding(5).AlignCenter().Text("%").Bold().FontSize(8).FontColor(Colors.White);
                        header.Cell().Background("#B71C1C").Padding(5).AlignCenter().Text("Bắt đầu").Bold().FontSize(8).FontColor(Colors.White);
                        header.Cell().Background("#B71C1C").Padding(5).AlignCenter().Text("Kết thúc").Bold().FontSize(8).FontColor(Colors.White);
                        header.Cell().Background("#B71C1C").Padding(5).AlignCenter().Text("Mốc tính").Bold().FontSize(8).FontColor(Colors.White);
                        header.Cell().Background("#B71C1C").Padding(5).AlignCenter().Text("Số VB/QĐ").Bold().FontSize(8).FontColor(Colors.White);
                        header.Cell().Background("#B71C1C").Padding(5).AlignCenter().Text("Ngày ký QĐ").Bold().FontSize(8).FontColor(Colors.White);
                    });

                    if (p.SalaryRecords != null && p.SalaryRecords.Count > 0)
                    {
                        foreach (var h in p.SalaryRecords.OrderByDescending(s => s.StartDate))
                        {
                            table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten3).Padding(5).AlignCenter().Text(h.RankCode ?? "").FontSize(8);
                            table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten3).Padding(5).AlignCenter().Text(h.SalaryStep ?? "").FontSize(8);
                            table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten3).Padding(5).AlignCenter().Text(h.Coefficient ?? "").FontSize(8).Bold();
                            table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten3).Padding(5).AlignCenter().Text($"{h.Percentage}%").FontSize(8);
                            table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten3).Padding(5).AlignCenter().Text(FormatDate(h.StartDate)).FontSize(8);
                            table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten3).Padding(5).AlignCenter().Text(FormatDate(h.EndDate)).FontSize(8);
                            table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten3).Padding(5).AlignCenter().Text(FormatDate(h.SalaryCalculationDate)).FontSize(8);
                            table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten3).Padding(5).AlignCenter().Text(h.DecisionNumber ?? "").FontSize(8);
                            table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten3).Padding(5).AlignCenter().Text(FormatDate(h.DecisionDate)).FontSize(8);
                        }
                    }
                    else
                    {
                        table.Cell().ColumnSpan(9).Padding(10).AlignCenter().Text("Chưa có dữ liệu lịch sử lương").Italic().FontColor(Colors.Grey.Medium);
                    }
                });
            });
        }

        static void ComposeLeave(ColumnDescriptor column, Personnel p)
        {
            // 1. Calculate stats matching UI logic
            double annualTakenCurrentYear = 0;
            double annualTakenOldYear = 0;
            int currentYear = DateTime.Now.Year;

            if (p.LeaveHistories != null)
            {
                foreach (var item in p.LeaveHistories)
                {
                    if (item.StartDate.Year == currentYear && item.LeaveType == "Phép năm")
                    {
                        if (item.LeaveYear.HasValue && item.LeaveYear.Value < currentYear)
                            annualTakenOldYear += item.DurationDays;
                        else
                            annualTakenCurrentYear += item.DurationDays;
                    }
                }
            }

            int used = (int)(annualTakenCurrentYear + annualTakenOldYear);
            int remaining = p.TotalAnnualLeaveDays - (int)annualTakenCurrentYear;
            if (remaining < 0) remaining = 0;

            // Stats
            column.Item().PaddingBottom(10).Row(row =>
            {
                row.RelativeItem().Background("#E3F2FD").Padding(10).Column(c =>
                {
                    c.Item().AlignCenter().Text("Tổng phép").FontSize(9).FontColor(Colors.Grey.Darken1);
                    c.Item().AlignCenter().Text(p.TotalAnnualLeaveDays.ToString()).FontSize(14).Bold();
                });

                row.Spacing(10);

                row.RelativeItem().Background("#FFEBEE").Padding(10).Column(c =>
                {
                    c.Item().AlignCenter().Text("Đã nghỉ").FontSize(9).FontColor(Colors.Red.Darken1);
                    c.Item().AlignCenter().Text(used.ToString()).FontSize(14).Bold().FontColor(Colors.Red.Darken1);
                });

                row.Spacing(10);

                row.RelativeItem().Background("#E8F5E9").Padding(10).Column(c =>
                {
                    c.Item().AlignCenter().Text("Còn lại").FontSize(9).FontColor(Colors.Green.Darken1);
                    c.Item().AlignCenter().Text(remaining.ToString()).FontSize(14).Bold().FontColor(Colors.Green.Darken1);
                });
            });

            // Table
            column.Item().Table(table =>
            {
                table.ColumnsDefinition(columns =>
                {
                    columns.ConstantColumn(80);   // Loại
                    columns.ConstantColumn(120);  // Thời gian
                    columns.ConstantColumn(60);   // Số ngày
                    columns.RelativeColumn(1.5f); // Lí do
                    columns.RelativeColumn(2.5f); // Ghi chú hệ thống
                });

                table.Header(header =>
                {
                    header.Cell().Background("#EEEEEE").Padding(5).AlignCenter().Text("Loại").Bold().FontSize(9);
                    header.Cell().Background("#EEEEEE").Padding(5).AlignCenter().Text("Thời gian").Bold().FontSize(9);
                    header.Cell().Background("#EEEEEE").Padding(5).AlignCenter().Text("Số ngày").Bold().FontSize(9);
                    header.Cell().Background("#EEEEEE").Padding(5).AlignCenter().Text("Lí do").Bold().FontSize(9);
                    header.Cell().Background("#EEEEEE").Padding(5).AlignCenter().Text("Ghi chú hệ thống").Bold().FontSize(9);
                });

                if (p.LeaveHistories != null)
                {
                    foreach (var h in p.LeaveHistories.OrderByDescending(x => x.StartDate).ThenByDescending(x => x.Id))
                    {
                        table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten3).Padding(5).AlignCenter().Text(h.LeaveType).FontSize(8).Bold();
                        table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten3).Padding(5).AlignCenter().Text($"{DatePickerHelper.FormatDateForDisplay(h.StartDate)} - {FormatDate(h.EndDate, "---")}").FontSize(8);
                        table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten3).Padding(5).AlignCenter().Text(h.DurationDays.ToString() + " ngày").FontSize(8);
                        table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten3).Padding(5).AlignCenter().Text(h.UserReasonDisplay).FontSize(8).Italic().FontColor("#455A64");
                        table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten3).Padding(5).AlignCenter().Text(h.SystemMessageDisplay).FontSize(8).FontColor("#D32F2F");
                    }
                }

                if (p.LeaveHistories == null || p.LeaveHistories.Count == 0)
                {
                    table.Cell().ColumnSpan(5).Padding(10).AlignCenter().Text("Chưa có dữ liệu lịch sử nghỉ").Italic().FontColor(Colors.Grey.Medium);
                }
            });
        }

        static void ComposeWorkHistory(ColumnDescriptor column, Personnel p)
        {
            column.Item().Text("A. THÔNG TIN CÔNG TÁC").Bold().FontColor("#B71C1C").FontSize(12);
            column.Item().PaddingTop(5).Row(row =>
            {
                row.RelativeItem().Column(c => LabelValue(c, "Thời gian công tác tại cơ quan thuế:", FormatDate(p.TaxAuthorityStartDate), true));
            });
            column.Item().PaddingTop(5).Row(row =>
            {
                row.RelativeItem().Column(c => LabelValue(c, "Thời gian công tác tính theo QĐ gần nhất:", FormatDate(p.PositionDecisionDate)));
                row.RelativeItem().Column(c => LabelValue(c, "Thời điểm tính thời gian công tác:", DatePickerHelper.FormatDateForDisplay(p.DisplayPositionCalculationDate)));
            });

            // Calculated Stats
            int wYears = 0;
            int wMonths = 0;
            if (p.PositionDecisionDate.HasValue)
            {
                DateTime startDate = p.PositionDecisionDate.Value;
                DateTime endDate = p.PositionCalculationDate ?? DateTime.Now;

                if (endDate >= startDate)
                {
                    // Clean calculation matching View logic
                    wYears = endDate.Year - startDate.Year;
                    if (startDate.Date > endDate.AddYears(-wYears)) wYears--;

                    DateTime tmpDate = startDate.AddYears(wYears);
                    while (tmpDate.AddMonths(1) <= endDate)
                    {
                        wMonths++;
                        tmpDate = tmpDate.AddMonths(1);
                    }
                }
            }

            // Box
            column.Item().PaddingTop(10).Background("#FAFAFA").CornerRadius(5).Padding(15).Row(row =>
            {
                void CenteredItem(RowDescriptor r, string label, string? val, bool last = false)
                {
                    r.RelativeItem().BorderRight(last ? 0 : 2).BorderColor(Colors.Grey.Lighten2).PaddingHorizontal(10).Column(c =>
                    {
                        c.Item().AlignCenter().Text(label).FontSize(10).FontColor(Colors.Grey.Darken1).AlignCenter();
                        c.Item().PaddingTop(5).AlignCenter().Text(val).Bold().FontSize(14);
                    });
                }

                CenteredItem(row, "Số năm công tác tính đến thời điểm hiện tại", wYears.ToString());
                CenteredItem(row, "Số tháng công tác tính đến thời điểm hiện tại", wMonths.ToString());
                CenteredItem(row, "Năm giữ vị trí công tác", p.CalculatedPositionYear, true);
            });

            // Retirement Stats
            string retYearsWorked = "---";
            string retRemaining = "---";
            DateTime now = DateTime.Now.Date;

            string CalcDuration(DateTime start, DateTime end)
            {
                if (start > end) return "0 năm 0 tháng 0 ngày";
                DateTime temp = start;
                int y = 0;
                while (temp.AddYears(1) <= end) { y++; temp = temp.AddYears(1); }
                int m = 0;
                while (temp.AddMonths(1) <= end) { m++; temp = temp.AddMonths(1); }
                int d = (end - temp).Days;
                return $"{y} năm {m} tháng {d} ngày";
            }

            if (p.TaxAuthorityStartDate.HasValue)
                retYearsWorked = CalcDuration(p.TaxAuthorityStartDate.Value, now);

            if (p.RetirementDate.HasValue)
                retRemaining = CalcDuration(now, p.RetirementDate.Value);

            column.Item().PaddingTop(15).Text("B. THÔNG TIN NGHỈ HƯU").Bold().FontColor("#B71C1C").FontSize(12);
            column.Item().PaddingTop(5).Row(row =>
            {
                row.RelativeItem().Column(c => LabelValue(c, "Ngày về hưu", FormatDate(p.RetirementDate)));
                row.RelativeItem().Column(c => LabelValueBlue(c, "Số năm công tác", retYearsWorked));
                row.RelativeItem().Column(c => LabelValueRed(c, "Số năm còn lại", retRemaining));
            });

            column.Item().PaddingTop(15).Text("C. QUÁ TRÌNH CÔNG TÁC CHI TIẾT").Bold().FontColor("#B71C1C").FontSize(12);
            column.Item().PaddingTop(5).Border(1).BorderColor(Colors.Grey.Lighten2).Padding(10).Text(p.DetailedWorkHistory ?? "").FontSize(10);
        }

        static void ComposeParty(ColumnDescriptor column, Personnel p)
        {
            column.Item().Row(row =>
            {
                row.RelativeItem().Column(c => LabelValue(c, "Ngày vào Đảng:", FormatDate(p.PartyEntryDate)));
                row.RelativeItem().Column(c => LabelValue(c, "Ngày chính thức:", FormatDate(p.PartyOfficialDate)));
            });
        }

        static void ComposeRewardDiscipline(ColumnDescriptor column, Personnel p)
        {
            column.Item().Text("A. THÔNG TIN KHEN THƯỞNG").Bold().FontColor("#1976D2");
            column.Item().PaddingTop(5).Column(c => LabelValueBox(c, "Danh hiệu thi đua:", p.EmulationTitles));
            column.Item().PaddingTop(5).Column(c => LabelValueBox(c, "Hình thức khen thưởng:", p.RewardForms));

            column.Item().PaddingTop(15).Text("B. THÔNG TIN KỶ LUẬT").Bold().FontColor("#D32F2F");

            if (!string.IsNullOrEmpty(p.DisciplineType) && p.DisciplineType != "-- Không có --" && p.DisciplineType != "---")
            {
                column.Item().PaddingTop(5).Background("#FFEBEE").Padding(15).Column(c =>
                {
                    c.Item().Text("Hình thức kỷ luật:").FontSize(10).Bold();
                    c.Item().Text(p.DisciplineType).FontSize(14).Bold().FontColor("#D32F2F");

                    c.Item().PaddingTop(10).Row(row =>
                    {
                        row.RelativeItem().Text(t => { t.Span("Số QĐ: ").Bold(); t.Span(p.DisciplineDecisionNumber ?? ""); });
                        row.RelativeItem().Text(t => { t.Span("Ngày ký: ").Bold(); t.Span(FormatDate(p.DisciplineDecisionDate)); });
                    });

                    c.Item().PaddingTop(10).LineHorizontal(1).LineColor("#EF9A9A");

                    c.Item().PaddingTop(5).Text("Nội dung / Lý do:").Bold();
                    c.Item().Text(p.DisciplineReason ?? "").Italic();
                });
            }
            else
            {
                column.Item().PaddingTop(5).Background("#E8F5E9").Padding(10).Text("Không có ghi nhận vi phạm kỷ luật").FontColor(Colors.Green.Darken2);
            }
        }

        // Helpers
        static void LabelValue(ColumnDescriptor c, string label, string? value, bool highlight = false)
        {
            c.Item().PaddingBottom(5).Column(col =>
            {
                col.Item().Text(label).FontSize(10).FontColor(Colors.Grey.Darken1);
                col.Item().Text(value ?? "---").Bold().FontSize(12).FontColor(highlight ? "#B71C1C" : Colors.Black);
            });
        }

        static void LabelValueBig(ColumnDescriptor c, string label, string? value, bool highlight = false)
        {
            c.Item().PaddingBottom(5).Column(col =>
            {
                col.Item().Text(label).FontSize(10).FontColor(Colors.Grey.Darken1);
                col.Item().Text(value ?? "---").Bold().FontSize(14).FontColor(highlight ? "#B71C1C" : Colors.Black);
            });
        }

        static void LabelValueBlue(ColumnDescriptor c, string label, string? value)
        {
            c.Item().PaddingBottom(5).Column(col =>
            {
                col.Item().Text(label).FontSize(10).FontColor(Colors.Grey.Darken1);
                col.Item().Text(value ?? "---").Bold().FontSize(12).FontColor("#1976D2");
            });
        }
        static void LabelValueRed(ColumnDescriptor c, string label, string? value)
        {
            c.Item().PaddingBottom(5).Column(col =>
            {
                col.Item().Text(label).FontSize(10).FontColor(Colors.Grey.Darken1);
                col.Item().Text(value ?? "---").Bold().FontSize(12).FontColor("#D32F2F");
            });
        }

        static void LabelValueBox(ColumnDescriptor c, string label, string? value)
        {
            c.Item().Text(label).FontSize(10).Bold();
            c.Item().Background("#F5F5F5").Padding(8).Text(value ?? "---");
        }

        static void ComposeEvaluationHistory(ColumnDescriptor column, Personnel p)
        {
            string headerBgColor = "#E65100";
            column.Item().Table(table =>
            {
                table.ColumnsDefinition(columns =>
                {
                    columns.RelativeColumn(1.0f); // Năm
                    columns.RelativeColumn(3.0f); // Xếp loại
                    columns.RelativeColumn(1.5f); // Số QĐ
                    columns.RelativeColumn(1.5f); // Ngày ký QĐ
                    columns.RelativeColumn(2.0f); // Đơn vị ra QĐ
                });

                table.Header(header =>
                {
                    header.Cell().Background(headerBgColor).Padding(5).AlignCenter().Text("Năm").Bold().FontSize(8).FontColor(Colors.White);
                    header.Cell().Background(headerBgColor).Padding(5).AlignCenter().Text("Xếp loại").Bold().FontSize(8).FontColor(Colors.White);
                    header.Cell().Background(headerBgColor).Padding(5).AlignCenter().Text("Số QĐ").Bold().FontSize(8).FontColor(Colors.White);
                    header.Cell().Background(headerBgColor).Padding(5).AlignCenter().Text("Ngày ký QĐ").Bold().FontSize(8).FontColor(Colors.White);
                    header.Cell().Background(headerBgColor).Padding(5).AlignCenter().Text("Đơn vị ra QĐ").Bold().FontSize(8).FontColor(Colors.White);
                });

                if (p.EvaluationRecords != null && p.EvaluationRecords.Count > 0)
                {
                    foreach (var h in p.EvaluationRecords.OrderByDescending(e => e.Year))
                    {
                        table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten3).Padding(5).AlignCenter().Text(h.Year.ToString()).FontSize(8);
                        table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten3).Padding(5).AlignCenter().Text(h.Rating ?? "").FontSize(8);
                        table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten3).Padding(5).AlignCenter().Text(h.DecisionNumber ?? "---").FontSize(8);
                        table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten3).Padding(5).AlignCenter().Text(FormatDate(h.DecisionDate, "---")).FontSize(8);
                        table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten3).Padding(5).AlignCenter().Text(h.DecisionAgency ?? "---").FontSize(8);
                    }
                }
                else
                {
                    table.Cell().ColumnSpan(5).Padding(10).AlignCenter().Text("Chưa có dữ liệu xếp loại").Italic().FontColor(Colors.Grey.Medium);
                }
            });
        }
    }
}
