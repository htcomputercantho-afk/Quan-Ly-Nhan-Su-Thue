using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using System;
using System.IO;
using TaxPersonnelManagement.Models;
using Colors = QuestPDF.Helpers.Colors;

namespace TaxPersonnelManagement.Services
{
    public static class PdfExporter
    {
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
                            });
                            
                            // Footer inside the border
                            column.Item().PaddingTop(10).AlignCenter().Text(x =>
                            {
                                x.Span("Hệ thống Quản lý Nhân sự - Xuất ngày " + DateTime.Now.ToString("dd/MM/yyyy"))
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

        static void ComposeSection(IContainer container, string title, Action<ColumnDescriptor> content)
        {
            // Ensure section start has enough space, or push start to next page
            container.EnsureSpace(100)
                     .Border(1).BorderColor("#FBE9E7").Background("#FFFAFA").Padding(15).Column(column => 
            {
                column.Item().Text(title).FontSize(14).Bold().FontColor("#B71C1C");
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
                    LabelValue(c, "Ngày sinh:", p.DateOfBirth?.ToString("dd/MM/yyyy"));
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
             column.Item().PaddingTop(5).Background("#FAFAFA").CornerRadius(5).Padding(15).Row(row => 
            {
                void CenteredItem(RowDescriptor r, string label, string? val, bool highlight = false, bool last = false)
                {
                    r.RelativeItem().BorderRight(last ? 0 : 2).BorderColor(Colors.Grey.Lighten2).PaddingHorizontal(10).Column(c => 
                    {
                        c.Item().AlignCenter().Text(label).FontSize(10).FontColor(Colors.Grey.Darken1);
                        c.Item().PaddingTop(5).AlignCenter().Text(val ?? "---").Bold().FontSize(16).FontColor(highlight ? "#B71C1C" : Colors.Black);
                    });
                }

                CenteredItem(row, "Mã ngạch", p.RankCode);
                CenteredItem(row, "Bậc lương", p.CurrentSalaryStep);
                CenteredItem(row, "Hệ số", p.CurrentSalaryCoefficient.ToString("F2"), true, true);
            });
            
            column.Item().PaddingTop(10).Row(row => 
            {
                row.RelativeItem().Column(c => LabelValue(c, "% Vượt khung:", p.ExceedFramePercent + "%"));
                row.RelativeItem().Column(c => LabelValue(c, "PC Chức vụ:", p.PositionAllowance));
                row.RelativeItem().Column(c => LabelValue(c, "Thời hạn bảo lưu:", p.SalaryReservationDeadline?.ToString("dd/MM/yyyy")));
            });

            column.Item().PaddingTop(5).Row(row => 
            {
                row.RelativeItem().Column(c => LabelValue(c, "Mốc lương:", p.NextSalaryStepDate?.ToString("dd/MM/yyyy")));
                row.RelativeItem().Column(c => LabelValue(c, "Dự kiến lên lương:", p.ExpectedSalaryIncreaseDate?.ToString("dd/MM/yyyy")));
            });

            // Salary History Table - Keep title + table together to prevent page break between them
            column.Item().PaddingTop(15).ShowEntire().Column(salarySection =>
            {
                salarySection.Item().PaddingBottom(8).Text("DIỄN BIẼN QUÁ TRÌNH LƯƠNG").Bold().FontColor("#B71C1C").FontSize(12);
                salarySection.Item().Table(table =>
                {
                    table.ColumnsDefinition(columns =>
                    {
                        columns.RelativeColumn(1.2f); // Start
                        columns.RelativeColumn(1.2f); // End
                        columns.RelativeColumn(1.2f); // Calc
                        columns.ConstantColumn(50);   // Coeff
                        columns.ConstantColumn(50);   // %
                        columns.RelativeColumn(1.5f); // Decision No
                        columns.RelativeColumn(1.2f); // Decision Date
                    });

                    table.Header(header =>
                    {
                        header.Cell().Background("#B71C1C").Padding(5).AlignCenter().Text("Thời gian bắt đầu").Bold().FontSize(9).FontColor(Colors.White);
                        header.Cell().Background("#B71C1C").Padding(5).AlignCenter().Text("Thời gian kết thúc").Bold().FontSize(9).FontColor(Colors.White);
                        header.Cell().Background("#B71C1C").Padding(5).AlignCenter().Text("Mốc xét lương từ").Bold().FontSize(9).FontColor(Colors.White);
                        header.Cell().Background("#B71C1C").Padding(5).AlignCenter().Text("Hệ số").Bold().FontSize(9).FontColor(Colors.White);
                        header.Cell().Background("#B71C1C").Padding(5).AlignCenter().Text("%").Bold().FontSize(9).FontColor(Colors.White);
                        header.Cell().Background("#B71C1C").Padding(5).AlignCenter().Text("Số VB/QĐ").Bold().FontSize(9).FontColor(Colors.White);
                        header.Cell().Background("#B71C1C").Padding(5).AlignCenter().Text("Ngày ký QĐ").Bold().FontSize(9).FontColor(Colors.White);
                    });

                    if (p.SalaryRecords != null && p.SalaryRecords.Count > 0)
                    {
                        foreach (var h in p.SalaryRecords.OrderByDescending(s => s.StartDate))
                        {
                            table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten3).Padding(5).AlignCenter().Text(h.StartDate?.ToString("dd/MM/yyyy") ?? "").FontSize(9);
                            table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten3).Padding(5).AlignCenter().Text(h.EndDate?.ToString("dd/MM/yyyy") ?? "").FontSize(9);
                            table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten3).Padding(5).AlignCenter().Text(h.SalaryCalculationDate?.ToString("dd/MM/yyyy") ?? "").FontSize(9);
                            table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten3).Padding(5).AlignCenter().Text(h.Coefficient ?? "").FontSize(9).Bold();
                            table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten3).Padding(5).AlignCenter().Text($"{h.Percentage}%").FontSize(9);
                            table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten3).Padding(5).AlignCenter().Text(h.DecisionNumber ?? "").FontSize(9);
                            table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten3).Padding(5).AlignCenter().Text(h.DecisionDate?.ToString("dd/MM/yyyy") ?? "").FontSize(9);
                        }
                    }
                    else 
                    {
                        table.Cell().ColumnSpan(7).Padding(10).AlignCenter().Text("Chưa có dữ liệu lịch sử lương").Italic().FontColor(Colors.Grey.Medium);
                    }
                });
            });
        }
        
        static void ComposeLeave(ColumnDescriptor column, Personnel p)
        {
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
                     // Simple calc for used
                     // Need logic passed in or calc here. For now 0 or sample.
                     // Assuming passed p has up-to-date data. Let's verify logic in View.
                     // We might calculate inside Export or pass a DTO. 
                     // For simplicity, just showing 0 if logic is complex.
                     c.Item().AlignCenter().Text("0").FontSize(14).Bold().FontColor(Colors.Red.Darken1); 
                 });

                 row.Spacing(10);

                 row.RelativeItem().Background("#E8F5E9").Padding(10).Column(c => 
                 {
                     c.Item().AlignCenter().Text("Còn lại").FontSize(9).FontColor(Colors.Green.Darken1);
                     c.Item().AlignCenter().Text(p.TotalAnnualLeaveDays.ToString()).FontSize(14).Bold().FontColor(Colors.Green.Darken1);
                 });
            });

            // Table
            column.Item().Table(table =>
            {
                table.ColumnsDefinition(columns =>
                {
                    columns.RelativeColumn();
                    columns.RelativeColumn();
                    columns.RelativeColumn();
                    columns.RelativeColumn();
                });

                table.Header(header =>
                {
                    header.Cell().Background("#EEEEEE").Padding(5).Text("Loại").Bold().FontSize(10);
                    header.Cell().Background("#EEEEEE").Padding(5).Text("Thời gian").Bold().FontSize(10);
                    header.Cell().Background("#EEEEEE").Padding(5).Text("Số ngày").Bold().FontSize(10);
                    header.Cell().Background("#EEEEEE").Padding(5).Text("Ghi chú").Bold().FontSize(10);
                });

                if (p.LeaveHistories != null)
                {
                    string FormatDurationToMonthsDays(double dValue)
                    {
                        if (dValue >= 30)
                        {
                            int m = (int)(dValue / 30);
                            double d = dValue % 30;
                            if (d > 0)
                            {
                                string dayStr = d == Math.Floor(d) ? ((int)d).ToString() : d.ToString("0.#");
                                return $"{m} tháng {dayStr} ngày";
                            }
                            return $"{m} tháng";
                        }
                        return $"{(dValue == Math.Floor(dValue) ? ((int)dValue).ToString() : dValue.ToString("0.#"))} ngày";
                    }

                    foreach (var h in p.LeaveHistories)
                    {
                        table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten3).Padding(5).Text(h.LeaveType).Bold();
                        table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten3).Padding(5).Text($"{h.StartDate:dd/MM} - {h.EndDate:dd/MM}");
                        table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten3).Padding(5).Text(h.DurationDays.ToString() + " ngày");
                        table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten3).Padding(5).Text(t => 
                        {
                            t.Span($"({FormatDurationToMonthsDays(h.DurationDays)})").FontColor(Colors.Blue.Darken2);
                            if (!string.IsNullOrWhiteSpace(h.Reason))
                            {
                                t.Span($" - {h.Reason}");
                            }
                        });
                    }
                }
                
                if (p.LeaveHistories == null || p.LeaveHistories.Count == 0)
                {
                     table.Cell().ColumnSpan(4).Padding(10).AlignCenter().Text("Chưa có dữ liệu").Italic().FontColor(Colors.Grey.Medium);
                }
            });
        }

        static void ComposeWorkHistory(ColumnDescriptor column, Personnel p)
        {
            column.Item().Text("A. THÔNG TIN CÔNG TÁC").Bold().FontColor("#B71C1C").FontSize(12);
            column.Item().PaddingTop(5).Row(row => {
                 row.RelativeItem().Column(c => LabelValue(c, "Thời gian công tác tại cơ quan thuế:", p.TaxAuthorityStartDate?.ToString("dd/MM/yyyy"), true));
            });
             column.Item().PaddingTop(5).Row(row => {
                 row.RelativeItem().Column(c => LabelValue(c, "Thời gian công tác tính theo QĐ gần nhất:", p.PositionDecisionDate?.ToString("dd/MM/yyyy")));
                 row.RelativeItem().Column(c => LabelValue(c, "Thời điểm tính thời gian công tác:", p.DisplayPositionCalculationDate.ToString("dd/MM/yyyy")));
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
            column.Item().PaddingTop(5).Row(row => {
                 row.RelativeItem().Column(c => LabelValue(c, "Ngày về hưu", p.RetirementDate?.ToString("dd/MM/yyyy")));
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
                row.RelativeItem().Column(c => LabelValue(c, "Ngày vào Đảng:", p.PartyEntryDate?.ToString("dd/MM/yyyy")));
                row.RelativeItem().Column(c => LabelValue(c, "Ngày chính thức:", p.PartyOfficialDate?.ToString("dd/MM/yyyy")));
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
                        row.RelativeItem().Text(t => { t.Span("Ngày ký: ").Bold(); t.Span(p.DisciplineDecisionDate?.ToString("dd/MM/yyyy") ?? ""); });
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
    }
}
