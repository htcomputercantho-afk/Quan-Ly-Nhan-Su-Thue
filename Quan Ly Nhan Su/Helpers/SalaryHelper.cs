using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.EntityFrameworkCore;
using TaxPersonnelManagement.Data;
using TaxPersonnelManagement.Models;

namespace TaxPersonnelManagement.Helpers
{
    /// <summary>
    /// Kết quả tính toán chu kỳ nâng bậc lương, thời gian lùi và ngày dự kiến lên lương.
    /// </summary>
    public class SalaryCalculationResult
    {
        public DateTime? BaseDate { get; set; }
        public int PeriodYears { get; set; } = 3;
        public DateTime? BaseExpectedDate { get; set; }
        public DateTime? ExpectedDate { get; set; }
        public string DelayType { get; set; } = "-- Không lùi --";
        public bool HasOngoingSickLeave { get; set; }
        public double TotalSickDaysInPeriod { get; set; }
        public double SickExcessDays { get; set; }
        public int SickDelayMonths { get; set; }
        public double UnpaidDays { get; set; }
        public int UnpaidDelayMonths { get; set; }
        public int DisciplinaryDelayMonths { get; set; }
        public string DisciplinaryReason { get; set; } = "";
    }

    /// <summary>
    /// Lớp hỗ trợ tập trung tính toán nâng bậc lương và quy đổi thời gian lùi nâng lương
    /// theo Điều 2 Thông tư 03/2021/TT-BNV của Bộ Nội vụ.
    /// </summary>
    public static class SalaryHelper
    {
        /// <summary>
        /// Quy đổi số ngày nghỉ lẻ thành số tháng lùi thời hạn nâng lương theo Điều 2 Thông tư 03/2021/TT-BNV:
        /// - Mỗi 30 ngày = 1 tháng.
        /// - Số ngày lẻ còn lại: nếu từ đủ 11 ngày trở lên (>= 11 ngày) thì làm tròn thành 1 tháng; nếu dưới 11 ngày thì không tính.
        /// </summary>
        public static int ConvertDaysToDelayMonths(double days)
        {
            if (days <= 0) return 0;
            int totalDays = (int)Math.Round(days);
            int months = totalDays / 30;
            int remainder = totalDays % 30;
            if (remainder >= 11)
            {
                months += 1;
            }
            return months;
        }

        /// <summary>
        /// Tính toán chi tiết mốc thời gian nâng bậc lương tiếp theo, chu kỳ nâng ngạch, các yếu tố lùi thời hạn
        /// (kỷ luật, nghỉ ốm vượt 180 ngày theo TT 03/2021, nghỉ không lương theo TT 03/2021).
        /// </summary>
        public static SalaryCalculationResult CalculateSalaryIncrease(
            Personnel? p,
            DateTime? baseDateOverride = null,
            string? rankCodeOverride = null,
            double? exceedFrameOverride = null,
            string? delayReasonOverride = null,
            DateTime? evaluationDate = null)
        {
            var res = new SalaryCalculationResult();
            DateTime evalDate = (evaluationDate ?? DateTime.Today).Date;

            if (p == null && !baseDateOverride.HasValue)
            {
                return res;
            }

            // 1. Mốc thời gian gốc: Thời điểm tính bậc lương lần sau
            DateTime? baseDate = baseDateOverride ?? p?.NextSalaryStepDate;
            if (!baseDate.HasValue && p?.SalaryRecords != null && p.SalaryRecords.Count > 0)
            {
                baseDate = p.SalaryRecords.OrderByDescending(s => s.StartDate).FirstOrDefault()?.SalaryCalculationDate;
            }

            if (!baseDate.HasValue)
            {
                res.DelayType = "-- Không lùi --";
                return res;
            }

            res.BaseDate = baseDate.Value.Date;

            // 2. Tính chu kỳ nâng bậc lương (1 năm, 2 năm hoặc 3 năm)
            int periodYears = 3;
            double exceedFrame = exceedFrameOverride ?? p?.ExceedFramePercent ?? 0;

            if (exceedFrame > 0)
            {
                // Đã lên % vượt khung thì 1 năm nâng 1 lần
                periodYears = 1;
            }
            else
            {
                // Ngạch 06.039-1, 01.011, 01.009 thì 2 năm nâng 1 lần
                string rc = (rankCodeOverride ?? p?.RankCode ?? "").Trim();
                if (string.Equals(rc, "06.039-1", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(rc, "01.011", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(rc, "01.009", StringComparison.OrdinalIgnoreCase))
                {
                    periodYears = 2;
                }
            }

            res.PeriodYears = periodYears;
            res.BaseExpectedDate = res.BaseDate.Value.AddYears(periodYears);
            DateTime expectedDate = res.BaseExpectedDate.Value;

            // 3. Thời gian lùi do kỷ luật
            string delayReason = delayReasonOverride ?? p?.SalaryIncreaseDelayType ?? "";
            int discMonths = 0;
            string discLabel = "";

            if (!string.IsNullOrEmpty(delayReason))
            {
                if (delayReason.Contains("Khiển trách") ||
                    (delayReason.Contains("Lùi 3 tháng") && !delayReason.Contains("ốm") && !delayReason.Contains("không lương")))
                {
                    discMonths = 3;
                    discLabel = "Khiển trách (Lùi 3 tháng)";
                }
                else if (delayReason.Contains("Cảnh cáo") ||
                         (delayReason.Contains("Lùi 6 tháng") && !delayReason.Contains("ốm") && !delayReason.Contains("không lương")))
                {
                    discMonths = 6;
                    discLabel = "Cảnh cáo (Lùi 6 tháng)";
                }
                else if (delayReason.Contains("Giáng chức") || delayReason.Contains("Cách chức") ||
                         (delayReason.Contains("Lùi 12 tháng") && !delayReason.Contains("ốm") && !delayReason.Contains("không lương")))
                {
                    discMonths = 12;
                    discLabel = "Giáng chức/Cách chức (Lùi 12 tháng)";
                }
            }

            res.DisciplinaryDelayMonths = discMonths;
            res.DisciplinaryReason = discLabel;
            if (discMonths > 0)
            {
                expectedDate = expectedDate.AddMonths(discMonths);
            }

            // 4. Lùi do nghỉ ốm vượt quá 6 tháng (180 ngày) theo Điều 2 TT 03/2021/TT-BNV
            const double SickLeaveThreshold = 180.0;
            double totalSickDaysInPeriod = 0;
            bool hasOngoingSickLeave = false;

            if (p?.LeaveHistories != null)
            {
                DateTime periodStart = res.BaseDate.Value;
                DateTime periodEnd = res.BaseExpectedDate.Value;

                foreach (var leave in p.LeaveHistories)
                {
                    if (leave.LeaveType != "Nghỉ ốm" &&
                        !leave.LeaveType.Contains("ốm", StringComparison.OrdinalIgnoreCase))
                        continue;

                    DateTime leaveStart = leave.StartDate.Date;

                    if (!leave.EndDate.HasValue)
                    {
                        // Đợt nghỉ ốm chưa có ngày kết thúc: tính tạm đến hôm nay
                        if (leaveStart <= evalDate)
                        {
                            hasOngoingSickLeave = true;
                            DateTime effectiveStart = leaveStart > periodStart ? leaveStart : periodStart;
                            DateTime effectiveEnd = evalDate < periodEnd ? evalDate : periodEnd;
                            if (effectiveEnd >= effectiveStart)
                            {
                                totalSickDaysInPeriod += (effectiveEnd - effectiveStart).TotalDays + 1;
                            }
                        }
                    }
                    else
                    {
                        // Đợt nghỉ đã kết thúc: lấy phần giao với kỳ nâng lương
                        DateTime leaveEnd = leave.EndDate.Value.Date;
                        DateTime effectiveStart = leaveStart > periodStart ? leaveStart : periodStart;
                        DateTime effectiveEnd = leaveEnd < periodEnd ? leaveEnd : periodEnd;
                        if (effectiveEnd >= effectiveStart)
                        {
                            totalSickDaysInPeriod += (effectiveEnd - effectiveStart).TotalDays + 1;
                        }
                    }
                }
            }

            double sickExcessDays = Math.Max(0, totalSickDaysInPeriod - SickLeaveThreshold);
            int sickDelayMonths = ConvertDaysToDelayMonths(sickExcessDays);

            res.TotalSickDaysInPeriod = totalSickDaysInPeriod;
            res.SickExcessDays = sickExcessDays;
            res.SickDelayMonths = sickDelayMonths;
            res.HasOngoingSickLeave = hasOngoingSickLeave;

            if (sickDelayMonths > 0)
            {
                expectedDate = expectedDate.AddMonths(sickDelayMonths);
            }

            // 5. Lùi do nghỉ không lương theo Điều 2 TT 03/2021/TT-BNV
            double unpaidDays = 0;
            if (p?.LeaveHistories != null)
            {
                foreach (var leave in p.LeaveHistories)
                {
                    if (leave.LeaveType != "Không lương" &&
                        !leave.LeaveType.Contains("không lương", StringComparison.OrdinalIgnoreCase))
                        continue;

                    if (leave.DurationDays > 0)
                    {
                        unpaidDays += leave.DurationDays;
                    }
                    else if (!leave.EndDate.HasValue)
                    {
                        DateTime leaveStart = leave.StartDate.Date;
                        if (leaveStart <= evalDate)
                        {
                            unpaidDays += (evalDate - leaveStart).TotalDays + 1;
                        }
                    }
                    else
                    {
                        unpaidDays += (leave.EndDate.Value.Date - leave.StartDate.Date).TotalDays + 1;
                    }
                }
            }

            int unpaidDelayMonths = ConvertDaysToDelayMonths(unpaidDays);
            res.UnpaidDays = unpaidDays;
            res.UnpaidDelayMonths = unpaidDelayMonths;

            if (unpaidDelayMonths > 0)
            {
                expectedDate = expectedDate.AddMonths(unpaidDelayMonths);
            }

            res.ExpectedDate = expectedDate;

            // 6. Xây dựng chuỗi hiển thị lý do lùi thời gian nâng lương chuẩn hóa
            var reasons = new List<string>();
            if (discMonths > 0)
            {
                reasons.Add(discLabel);
            }
            if (sickDelayMonths > 0)
            {
                reasons.Add($"Nghỉ ốm quá hạn {(int)Math.Round(sickExcessDays)} ngày (Lùi {sickDelayMonths} tháng)");
            }
            if (unpaidDelayMonths > 0)
            {
                reasons.Add($"Nghỉ không lương ({(int)Math.Round(unpaidDays)} ngày - Lùi {unpaidDelayMonths} tháng)");
            }

            if (reasons.Count > 0)
            {
                res.DelayType = string.Join("; ", reasons);
            }
            else
            {
                res.DelayType = "-- Không lùi --";
            }

            return res;
        }

        /// <summary>
        /// Đồng bộ kết quả tính toán nâng lương vào đối tượng Personnel và lưu vào CSDL nếu có thay đổi.
        /// </summary>
        public static bool SyncSalaryIncreaseForPersonnel(Personnel p, AppDbContext? db = null, DateTime? evaluationDate = null)
        {
            if (p == null) return false;

            var res = CalculateSalaryIncrease(p, evaluationDate: evaluationDate);
            bool isModified = false;

            if (p.SalaryIncreaseDelayType != res.DelayType)
            {
                p.SalaryIncreaseDelayType = res.DelayType;
                isModified = true;
            }

            if (p.ExpectedSalaryIncreaseDate != res.ExpectedDate)
            {
                p.ExpectedSalaryIncreaseDate = res.ExpectedDate;
                isModified = true;
            }

            if (isModified && p.Id > 0)
            {
                try
                {
                    if (db != null)
                    {
                        var entity = db.Personnel.Find(p.Id);
                        if (entity != null)
                        {
                            entity.SalaryIncreaseDelayType = p.SalaryIncreaseDelayType;
                            entity.ExpectedSalaryIncreaseDate = p.ExpectedSalaryIncreaseDate;
                            db.SaveChanges();
                        }
                    }
                    else
                    {
                        using (var newDb = new AppDbContext())
                        {
                            var entity = newDb.Personnel.Find(p.Id);
                            if (entity != null)
                            {
                                entity.SalaryIncreaseDelayType = p.SalaryIncreaseDelayType;
                                entity.ExpectedSalaryIncreaseDate = p.ExpectedSalaryIncreaseDate;
                                newDb.SaveChanges();
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    App.DebugLog($"SyncSalaryIncreaseForPersonnel error for Id {p.Id}: {ex.Message}");
                }
            }

            return isModified;
        }

        /// <summary>
        /// Quét toàn bộ nhân sự trong hệ thống có đợt nghỉ hoặc dữ liệu cũ cần chuẩn hóa để đồng bộ lại ngày dự kiến và chuỗi lùi lương.
        /// </summary>
        public static int SyncAllPersonnelSalaryIncreases(AppDbContext db, DateTime? evaluationDate = null)
        {
            int updatedCount = 0;
            try
            {
                var candidates = db.Personnel
                    .Include(p => p.LeaveHistories)
                    .Include(p => p.SalaryRecords)
                    .Where(p => string.IsNullOrEmpty(p.Status) || p.Status == "Đang công tác")
                    .ToList();

                foreach (var p in candidates)
                {
                    var res = CalculateSalaryIncrease(p, evaluationDate: evaluationDate);
                    bool changed = false;

                    if (p.SalaryIncreaseDelayType != res.DelayType)
                    {
                        p.SalaryIncreaseDelayType = res.DelayType;
                        changed = true;
                    }

                    if (p.ExpectedSalaryIncreaseDate != res.ExpectedDate)
                    {
                        p.ExpectedSalaryIncreaseDate = res.ExpectedDate;
                        changed = true;
                    }

                    if (changed)
                    {
                        updatedCount++;
                    }
                }

                if (updatedCount > 0)
                {
                    db.SaveChanges();
                    App.DebugLog($"SyncAllPersonnelSalaryIncreases: Updated {updatedCount} personnel records.");
                }
            }
            catch (Exception ex)
            {
                App.DebugLog($"SyncAllPersonnelSalaryIncreases error: {ex.Message}");
            }

            return updatedCount;
        }
    }
}
