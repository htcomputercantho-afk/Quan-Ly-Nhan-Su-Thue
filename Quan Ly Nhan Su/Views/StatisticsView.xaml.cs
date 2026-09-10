using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;
using TaxPersonnelManagement.Models;
using TaxPersonnelManagement.Data;
using Microsoft.EntityFrameworkCore;
using ClosedXML.Excel;
using Microsoft.Win32;

namespace TaxPersonnelManagement.Views
{
    public partial class StatisticsView : UserControl
    {
        // Cờ kiểm soát khởi tạo, tránh trigger sự kiện lọc trước khi UI sẵn sàng
        private bool _isInitialized = false;

        public StatisticsView()
        {
            InitializeComponent();
            InitializeFilters();
            _isInitialized = true;
            LoadStatistics();
        }

        // Khởi tạo dữ liệu cho các bộ lọc (Bộ phận, Năm)
        private void InitializeFilters()
        {
            try
            {
                // Bộ lọc Bộ phận
                cbDepartmentFilter.Items.Clear();
                cbDepartmentFilter.Items.Add("Tất cả bộ phận");

                using (var db = new AppDbContext())
                {
                    var departments = db.Departments.Select(d => d.Name).OrderBy(n => n).ToList();
                    foreach (var dept in departments)
                        cbDepartmentFilter.Items.Add(dept);
                }
                cbDepartmentFilter.SelectedIndex = 0;

                // Bộ lọc Năm (từ năm nhỏ nhất đến năm hiện tại, sắp xếp mới nhất lên đầu)
                cbYearFilter.Items.Clear();

                int currentYear = DateTime.Now.Year;
                int minYear = 1990;

                using (var db = new AppDbContext())
                {
                    if (db.Personnel.Any())
                    {
                        var yearsList = db.Personnel
                            .Where(p => string.IsNullOrEmpty(p.Status) || p.Status == "Đang công tác")
                            .Select(p => p.TaxAuthorityStartDate.HasValue
                                ? p.TaxAuthorityStartDate.Value.Year
                                : (p.StartDate.HasValue ? p.StartDate.Value.Year : 2020))
                            .Where(y => y > 1900)
                            .ToList();

                        if (yearsList.Any())
                        {
                            int empMin = yearsList.Min();
                            if (empMin > 1900 && empMin < currentYear)
                                minYear = empMin;
                        }
                    }
                }

                if (minYear < 1990) minYear = 1990;

                // Thêm năm mới nhất lên đầu danh sách
                for (int y = currentYear; y >= minYear; y--)
                    cbYearFilter.Items.Add(y);

                cbYearFilter.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                App.DebugLog("Lỗi khởi tạo bộ lọc thống kê: " + ex.Message);
            }
        }

        // Sự kiện khi thay đổi ComboBox bộ lọc
        private void Filter_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (!_isInitialized) return;
            LoadStatistics();
        }



        // Lấy danh sách nhân sự theo các bộ lọc hiện tại
        private List<Personnel> GetFilteredPersonnel(AppDbContext db, out int filterYear, out string departmentStr, out string yearStr)
        {
            var query = db.Personnel.AsQueryable();

            // Chỉ thống kê công chức đang làm việc / đang công tác (loại trừ luân chuyển điều động, nghỉ hưu, nghỉ việc)
            query = query.Where(p => string.IsNullOrEmpty(p.Status) || p.Status == "Đang công tác");

            // Lọc theo Bộ phận
            departmentStr = "Tất cả bộ phận";
            if (cbDepartmentFilter.SelectedIndex > 0 && cbDepartmentFilter.SelectedItem is string dept)
            {
                departmentStr = dept;
                query = query.Where(p => p.Department == dept);
            }

            // Lọc theo Năm
            filterYear = DateTime.Now.Year;
            if (cbYearFilter.SelectedItem is int yr)
            {
                filterYear = yr;
            }
            yearStr = filterYear.ToString();

            // Dùng biến local vì C# không cho phép capture tham số 'out' trong lambda
            int localYear = filterYear;
            var endOfYear = new DateTime(localYear, 12, 31);
            // Chỉ lấy những nhân sự đã bắt đầu làm việc trong hoặc trước năm đó
            query = query.Where(p => (p.TaxAuthorityStartDate ?? p.StartDate) <= endOfYear);
            // Loại trừ nhân sự đã nghỉ hưu trước năm đó
            query = query.Where(p => !p.RetirementDate.HasValue || p.RetirementDate.Value.Year >= localYear);

            return query.ToList();
        }

        // Tải và hiển thị toàn bộ dữ liệu thống kê
        public void LoadStatistics()
        {
            try
            {
                using (var db = new AppDbContext())
                {
                    var list = GetFilteredPersonnel(db, out int filterYear, out _, out _);

                    // 1. Tổng số nhân sự
                    int total = list.Count;
                    txtTotalPersonnel.Text = total.ToString();

                    // 2. Phân bố giới tính
                    int male = list.Count(p => p.Gender == "Nam");
                    int female = total - male;

                    txtMaleCount.Text = $"{male} người";
                    txtFemaleCount.Text = $"{female} người";

                    double maleRatio = total > 0 ? (double)male / total * 100 : 0;
                    double femaleRatio = total > 0 ? (double)female / total * 100 : 0;

                    txtMaleRatio.Text = $"Tỷ lệ: {maleRatio:F1}%";
                    txtFemaleRatio.Text = $"Tỷ lệ: {femaleRatio:F1}%";

                    // 3. Trình độ học vấn
                    int sauDaiHoc = 0, daiHoc = 0, khacEdu = 0;
                    ClassifyEducation(list, ref sauDaiHoc, ref daiHoc, ref khacEdu);

                    txtEduCenterCount.Text = total.ToString();
                    txtPostGradCount.Text = sauDaiHoc.ToString();
                    txtGradCount.Text = daiHoc.ToString();
                    txtOtherEduCount.Text = khacEdu.ToString();

                    double sauDaiHocRatio = total > 0 ? (double)sauDaiHoc / total * 100 : 0;
                    double daiHocRatio = total > 0 ? (double)daiHoc / total * 100 : 0;
                    double khacEduRatio = total > 0 ? (double)khacEdu / total * 100 : 0;

                    txtPostGradRatio.Text = $" ({sauDaiHocRatio:F1}%)";
                    txtGradRatio.Text = $" ({daiHocRatio:F1}%)";
                    txtOtherEduRatio.Text = $" ({khacEduRatio:F1}%)";

                    int countHighEdu = sauDaiHoc + daiHoc;
                    double highEduRatio = total > 0 ? (double)countHighEdu / total * 100 : 0;
                    txtEducationSummary.Text = $"* Tổng số trình độ Đại học trở lên: {countHighEdu} người (Tỷ lệ: {highEduRatio:F1}%). Trong đó trình độ Sau Đại học (Thạc sĩ, Tiến sĩ) chiếm {sauDaiHocRatio:F1}%.";

                    // Vẽ biểu đồ tròn trình độ học vấn
                    RenderDonutChart(cvEducation, new List<PieSlice>
                    {
                        new PieSlice { Value = sauDaiHoc, ColorBrush = new SolidColorBrush(Color.FromRgb(123, 31, 162)) },
                        new PieSlice { Value = daiHoc,    ColorBrush = new SolidColorBrush(Color.FromRgb(21, 101, 192)) },
                        new PieSlice { Value = khacEdu,   ColorBrush = new SolidColorBrush(Color.FromRgb(120, 144, 156)) }
                    });

                    // 4. Đảng viên
                    int party = list.Count(p => p.PartyEntryDate.HasValue);
                    int nonParty = total - party;

                    txtPartyCenterCount.Text = total.ToString();
                    txtPartyMemberCount.Text = party.ToString();
                    txtNonPartyCount.Text = nonParty.ToString();

                    double partyRatio = total > 0 ? (double)party / total * 100 : 0;
                    double nonPartyRatio = total > 0 ? (double)nonParty / total * 100 : 0;

                    txtPartyMemberRatio.Text = $" ({partyRatio:F1}%)";
                    txtNonPartyRatio.Text = $" ({nonPartyRatio:F1}%)";
                    txtPartySummary.Text = $"* Số lượng Đảng viên trong đơn vị: {party} đồng chí (Tỷ lệ: {partyRatio:F1}%). Số quần chúng chưa vào Đảng: {nonParty} người ({nonPartyRatio:F1}%).";

                    // Vẽ biểu đồ tròn đảng viên
                    RenderDonutChart(cvParty, new List<PieSlice>
                    {
                        new PieSlice { Value = party,    ColorBrush = new SolidColorBrush(Color.FromRgb(211, 47, 47)) },
                        new PieSlice { Value = nonParty, ColorBrush = new SolidColorBrush(Color.FromRgb(255, 160, 0)) }
                    });

                    // 5. Phân bố độ tuổi
                    DateTime relativeDate = filterYear > 0 ? new DateTime(filterYear, 12, 31) : DateTime.Today;
                    int age1 = 0, age2 = 0, age3 = 0, age4 = 0;
                    ClassifyAge(list, relativeDate, ref age1, ref age2, ref age3, ref age4);

                    int ageTotal = age1 + age2 + age3 + age4;
                    double age1Ratio = ageTotal > 0 ? (double)age1 / ageTotal * 100 : 0;
                    double age2Ratio = ageTotal > 0 ? (double)age2 / ageTotal * 100 : 0;
                    double age3Ratio = ageTotal > 0 ? (double)age3 / ageTotal * 100 : 0;
                    double age4Ratio = ageTotal > 0 ? (double)age4 / ageTotal * 100 : 0;

                    txtAgeVal1.Text = $"{age1} ({age1Ratio:F1}%)";
                    txtAgeVal2.Text = $"{age2} ({age2Ratio:F1}%)";
                    txtAgeVal3.Text = $"{age3} ({age3Ratio:F1}%)";
                    txtAgeVal4.Text = $"{age4} ({age4Ratio:F1}%)";

                    // Vẽ biểu đồ cột độ tuổi (chiều cao tỷ lệ theo giá trị lớn nhất)
                    double maxAgePct = Math.Max(Math.Max(age1Ratio, age2Ratio), Math.Max(age3Ratio, age4Ratio));
                    barAge1.Height = maxAgePct > 0 ? (age1Ratio / maxAgePct) * 120 : 0;
                    barAge2.Height = maxAgePct > 0 ? (age2Ratio / maxAgePct) * 120 : 0;
                    barAge3.Height = maxAgePct > 0 ? (age3Ratio / maxAgePct) * 120 : 0;
                    barAge4.Height = maxAgePct > 0 ? (age4Ratio / maxAgePct) * 120 : 0;

                    // 6. Thâm niên công tác
                    int sen1 = 0, sen2 = 0, sen3 = 0, sen4 = 0;
                    ClassifySeniority(list, relativeDate, ref sen1, ref sen2, ref sen3, ref sen4);

                    int senTotal = sen1 + sen2 + sen3 + sen4;
                    double sen1Ratio = senTotal > 0 ? (double)sen1 / senTotal * 100 : 0;
                    double sen2Ratio = senTotal > 0 ? (double)sen2 / senTotal * 100 : 0;
                    double sen3Ratio = senTotal > 0 ? (double)sen3 / senTotal * 100 : 0;
                    double sen4Ratio = senTotal > 0 ? (double)sen4 / senTotal * 100 : 0;

                    txtSenVal1.Text = $"{sen1} ({sen1Ratio:F1}%)";
                    txtSenVal2.Text = $"{sen2} ({sen2Ratio:F1}%)";
                    txtSenVal3.Text = $"{sen3} ({sen3Ratio:F1}%)";
                    txtSenVal4.Text = $"{sen4} ({sen4Ratio:F1}%)";

                    // Vẽ biểu đồ cột thâm niên
                    double maxSenPct = Math.Max(Math.Max(sen1Ratio, sen2Ratio), Math.Max(sen3Ratio, sen4Ratio));
                    barSen1.Height = maxSenPct > 0 ? (sen1Ratio / maxSenPct) * 120 : 0;
                    barSen2.Height = maxSenPct > 0 ? (sen2Ratio / maxSenPct) * 120 : 0;
                    barSen3.Height = maxSenPct > 0 ? (sen3Ratio / maxSenPct) * 120 : 0;
                    barSen4.Height = maxSenPct > 0 ? (sen4Ratio / maxSenPct) * 120 : 0;

                    // 7. Thống kê các ngạch công chức
                    var rankRows = CalculateRankStatistics(list, db, total);
                    icRankStatistics.ItemsSource = rankRows;

                    int rankTotalMale = rankRows.Sum(r => r.MaleCount);
                    int rankTotalFemale = rankRows.Sum(r => r.FemaleCount);
                    int rankTotalAll = rankRows.Sum(r => r.TotalCount);

                    txtRankTotalMale.Text = rankTotalMale.ToString();
                    txtRankTotalFemale.Text = rankTotalFemale.ToString();
                    txtRankTotalAll.Text = rankTotalAll.ToString();

                    int distinctRanks = rankRows.Count(r => !r.IsUnassigned);
                    var topRank = rankRows.Where(r => !r.IsUnassigned).OrderByDescending(r => r.TotalCount).FirstOrDefault();
                    if (topRank != null && total > 0)
                    {
                        txtRankSummary.Text = $"* Đơn vị hiện có {distinctRanks} ngạch công chức. Ngạch có số lượng đông nhất là \"{topRank.RankName}\" ({topRank.RankCode}) với {topRank.TotalCount} công chức (chiếm {topRank.Ratio:F1}%).";
                    }
                    else if (rankRows.Any())
                    {
                        txtRankSummary.Text = $"* Đơn vị có {total} công chức nhưng chưa được phân mã/tên ngạch cụ thể.";
                    }
                    else
                    {
                        txtRankSummary.Text = "* Chưa có dữ liệu phân ngạch công chức.";
                    }

                    // 8. Thống kê lãnh đạo tổ theo từng bộ phận
                    var teamRows = CalculateTeamLeadershipRows(list);
                    icTeamLeadership.ItemsSource = teamRows;
                }
            }
            catch (Exception ex)
            {
                App.DebugLog("Lỗi tải dữ liệu thống kê: " + ex.Message);
            }
        }

        // Phân loại nhân sự theo trình độ học vấn
        private static void ClassifyEducation(List<Personnel> list, ref int sauDaiHoc, ref int daiHoc, ref int khacEdu)
        {
            foreach (var p in list)
            {
                string edu = (p.EducationLevel ?? "").ToLower();
                if (edu.Contains("thạc sĩ") || edu.Contains("tiến sĩ") || edu.Contains("sau đại học") ||
                    edu.Contains("sau đh") || edu.Contains("cao học"))
                    sauDaiHoc++;
                else if (edu.Contains("đại học") || edu.Contains("cử nhân") || edu.Contains("kỹ sư") || edu.Contains("đh"))
                    daiHoc++;
                else
                    khacEdu++;
            }
        }

        // Phân loại nhân sự theo độ tuổi tại thời điểm tính
        private static void ClassifyAge(List<Personnel> list, DateTime relativeDate,
            ref int age1, ref int age2, ref int age3, ref int age4)
        {
            foreach (var p in list)
            {
                if (!p.DateOfBirth.HasValue) continue;
                int age = relativeDate.Year - p.DateOfBirth.Value.Year;
                if (p.DateOfBirth.Value.Date > relativeDate.AddYears(-age)) age--;

                if (age < 30) age1++;
                else if (age <= 40) age2++;
                else if (age <= 50) age3++;
                else age4++;
            }
        }

        // Phân loại nhân sự theo thâm niên công tác tại thời điểm tính
        private static void ClassifySeniority(List<Personnel> list, DateTime relativeDate,
            ref int sen1, ref int sen2, ref int sen3, ref int sen4)
        {
            foreach (var p in list)
            {
                DateTime? startDate = p.TaxAuthorityStartDate ?? p.StartDate;
                if (!startDate.HasValue) continue;
                double years = (relativeDate - startDate.Value).TotalDays / 365.25;

                if (years < 5) sen1++;
                else if (years <= 10) sen2++;
                else if (years <= 20) sen3++;
                else sen4++;
            }
        }

        /// <summary>
        /// Tính toán danh sách thống kê theo các ngạch công chức.
        /// </summary>
        private static List<RankStatisticsRow> CalculateRankStatistics(List<Personnel> list, AppDbContext db, int total)
        {
            var allRanks = db.Ranks.AsNoTracking().ToList();
            var rankDictByCode = allRanks
                .Where(r => !string.IsNullOrEmpty(r.Code))
                .GroupBy(r => r.Code.Trim(), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.First().Name, StringComparer.OrdinalIgnoreCase);

            var rankDictByName = allRanks
                .Where(r => !string.IsNullOrEmpty(r.Name))
                .GroupBy(r => r.Name.Trim(), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.First().Code, StringComparer.OrdinalIgnoreCase);

            var rankGroups = list
                .GroupBy(p =>
                {
                    string code = (p.RankCode ?? "").Trim();
                    string name = (p.RankName ?? "").Trim();

                    if (string.IsNullOrEmpty(code) && !string.IsNullOrEmpty(name))
                    {
                        if (rankDictByName.TryGetValue(name, out var mappedCode))
                            code = mappedCode;
                    }
                    else if (!string.IsNullOrEmpty(code) && string.IsNullOrEmpty(name))
                    {
                        if (rankDictByCode.TryGetValue(code, out var mappedName))
                            name = mappedName;
                    }

                    if (string.IsNullOrEmpty(code) && string.IsNullOrEmpty(name))
                    {
                        return (Code: "", Name: "Chưa phân ngạch");
                    }

                    return (Code: code, Name: string.IsNullOrEmpty(name) ? code : name);
                })
                .OrderBy(g => string.IsNullOrEmpty(g.Key.Code) ? 1 : 0) // Chưa phân ngạch xuống cuối
                .ThenBy(g => g.Key.Code)
                .ThenBy(g => g.Key.Name)
                .ToList();

            int rankStt = 1;
            var rankRows = new List<RankStatisticsRow>();

            foreach (var g in rankGroups)
            {
                int m = g.Count(p => p.Gender == "Nam");
                int f = g.Count(p => p.Gender != "Nam");
                int t = g.Count();
                double ratio = total > 0 ? (double)t / total * 100 : 0;

                rankRows.Add(new RankStatisticsRow
                {
                    STT = rankStt++,
                    RankCode = string.IsNullOrEmpty(g.Key.Code) ? "---" : g.Key.Code,
                    RankName = g.Key.Name,
                    MaleCount = m,
                    FemaleCount = f,
                    TotalCount = t,
                    Ratio = ratio,
                    RatioText = $"{ratio:F1}%",
                    BarRatio = ratio,
                    IsUnassigned = string.IsNullOrEmpty(g.Key.Code)
                });
            }

            return rankRows;
        }

        /// <summary>
        /// Tính toán danh sách thống kê lãnh đạo tổ theo từng bộ phận.
        /// </summary>
        private static List<TeamLeadershipRow> CalculateTeamLeadershipRows(List<Personnel> list)
        {
            var leaderPositions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                { "Tổ trưởng", "Đội trưởng", "Trưởng phòng", "Trưởng Thuế cơ sở" };
            var deputyPositions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                { "Phó Tổ trưởng", "Phó Đội trưởng", "Phó Trưởng phòng", "Phó Trưởng Thuế cơ sở" };

            // Bộ phận bị loại khỏi bảng thống kê lãnh đạo tổ (không có khái niệm Tổ trưởng/Phó Tổ trưởng)
            var excludedDepts = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                { "Ban Lãnh đạo", "Ban lãnh đạo", "Lãnh đạo đơn vị", "Ban Giám đốc", "Ban giám đốc",
                  "HĐLĐ", "Hội đồng lao động", "Hợp đồng lao động" };

            return list
                .Where(p => !string.IsNullOrWhiteSpace(p.Department)
                         && !excludedDepts.Contains(p.Department!))
                .GroupBy(p => p.Department!)
                .OrderBy(g => g.Key)
                .Select(g =>
                {
                    int leaders  = g.Count(p => leaderPositions.Contains(p.Position ?? ""));
                    int deputies = g.Count(p => deputyPositions.Contains(p.Position ?? ""));
                    int staff    = g.Count(p =>
                        !leaderPositions.Contains(p.Position ?? "") &&
                        !deputyPositions.Contains(p.Position ?? ""));
                    return new TeamLeadershipRow
                    {
                        DepartmentName = g.Key,
                        LeaderCount    = leaders,
                        DeputyCount    = deputies,
                        StaffCount     = staff,
                        TotalCount     = g.Count()
                    };
                })
                .ToList();
        }

        // Vẽ biểu đồ tròn (Donut Chart) bằng WPF Path/ArcSegment
        private void RenderDonutChart(Canvas canvas, List<PieSlice> slices)
        {
            canvas.Children.Clear();
            double total = slices.Sum(s => s.Value);

            // Nếu không có dữ liệu, hiển thị vòng tròn xám trống
            if (total == 0)
            {
                var emptyRing = new Ellipse
                {
                    Width = 140, Height = 140,
                    Stroke = Brushes.LightGray,
                    StrokeThickness = 20
                };
                canvas.Children.Add(emptyRing);
                Canvas.SetLeft(emptyRing, 15);
                Canvas.SetTop(emptyRing, 15);
                return;
            }

            double cx = 85, cy = 85, r = 70, rInner = 42;
            double currentAngle = -90.0; // Bắt đầu từ đỉnh (12 giờ)

            foreach (var slice in slices)
            {
                if (slice.Value == 0) continue;

                double sweepAngle = (slice.Value / total) * 360.0;
                // Clamp để tránh lỗi vẽ vòng tròn 360° đầy
                if (Math.Abs(sweepAngle - 360.0) < 0.01) sweepAngle = 359.99;

                double nextAngle = currentAngle + sweepAngle;
                canvas.Children.Add(CreateDonutSlice(cx, cy, r, rInner, currentAngle, nextAngle, slice.ColorBrush));
                currentAngle = nextAngle;
            }
        }

        // Tạo một lát cắt của biểu đồ tròn dạng vành khuyên
        private static Path CreateDonutSlice(double cx, double cy, double r, double rInner,
            double startAngle, double endAngle, Brush fillBrush)
        {
            double startRad = startAngle * Math.PI / 180.0;
            double endRad = endAngle * Math.PI / 180.0;

            Point p1 = new Point(cx + r * Math.Cos(startRad), cy + r * Math.Sin(startRad));
            Point p2 = new Point(cx + r * Math.Cos(endRad), cy + r * Math.Sin(endRad));
            Point p3 = new Point(cx + rInner * Math.Cos(endRad), cy + rInner * Math.Sin(endRad));
            Point p4 = new Point(cx + rInner * Math.Cos(startRad), cy + rInner * Math.Sin(startRad));

            bool isLargeArc = (endAngle - startAngle) > 180.0;

            var figure = new PathFigure { StartPoint = p1, IsClosed = true };
            // Cung ngoài (chiều thuận)
            figure.Segments.Add(new ArcSegment { Point = p2, Size = new Size(r, r), SweepDirection = SweepDirection.Clockwise, IsLargeArc = isLargeArc });
            // Đường thẳng nối vào cung trong
            figure.Segments.Add(new LineSegment { Point = p3 });
            // Cung trong (chiều nghịch để tạo lỗ hổng)
            figure.Segments.Add(new ArcSegment { Point = p4, Size = new Size(rInner, rInner), SweepDirection = SweepDirection.Counterclockwise, IsLargeArc = isLargeArc });

            var pathGeometry = new PathGeometry();
            pathGeometry.Figures.Add(figure);

            return new Path { Data = pathGeometry, Fill = fillBrush, Stroke = Brushes.White, StrokeThickness = 1.5 };
        }

        // Xuất báo cáo thống kê ra file Excel
        private void btnExportExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using (var db = new AppDbContext())
                {
                    var list = GetFilteredPersonnel(db, out int filterYear, out string departmentStr, out string yearStr);

                    // Tính toán các chỉ số thống kê
                    int total = list.Count;
                    int male = list.Count(p => p.Gender == "Nam");
                    int female = total - male;
                    double maleRatio = total > 0 ? (double)male / total * 100 : 0;
                    double femaleRatio = total > 0 ? (double)female / total * 100 : 0;

                    int sauDaiHoc = 0, daiHoc = 0, khacEdu = 0;
                    ClassifyEducation(list, ref sauDaiHoc, ref daiHoc, ref khacEdu);
                    double sauDaiHocRatio = total > 0 ? (double)sauDaiHoc / total * 100 : 0;
                    double daiHocRatio = total > 0 ? (double)daiHoc / total * 100 : 0;
                    double khacEduRatio = total > 0 ? (double)khacEdu / total * 100 : 0;

                    int party = list.Count(p => p.PartyEntryDate.HasValue);
                    int nonParty = total - party;
                    double partyRatio = total > 0 ? (double)party / total * 100 : 0;
                    double nonPartyRatio = total > 0 ? (double)nonParty / total * 100 : 0;

                    DateTime relativeDate = filterYear > 0 ? new DateTime(filterYear, 12, 31) : DateTime.Today;

                    int age1 = 0, age2 = 0, age3 = 0, age4 = 0;
                    ClassifyAge(list, relativeDate, ref age1, ref age2, ref age3, ref age4);
                    double age1Ratio = total > 0 ? (double)age1 / total * 100 : 0;
                    double age2Ratio = total > 0 ? (double)age2 / total * 100 : 0;
                    double age3Ratio = total > 0 ? (double)age3 / total * 100 : 0;
                    double age4Ratio = total > 0 ? (double)age4 / total * 100 : 0;

                    int sen1 = 0, sen2 = 0, sen3 = 0, sen4 = 0;
                    ClassifySeniority(list, relativeDate, ref sen1, ref sen2, ref sen3, ref sen4);
                    double sen1Ratio = total > 0 ? (double)sen1 / total * 100 : 0;
                    double sen2Ratio = total > 0 ? (double)sen2 / total * 100 : 0;
                    double sen3Ratio = total > 0 ? (double)sen3 / total * 100 : 0;
                    double sen4Ratio = total > 0 ? (double)sen4 / total * 100 : 0;

                    // Hộp thoại lưu file
                    var saveFileDialog = new SaveFileDialog
                    {
                        Filter = "Excel Files|*.xlsx",
                        FileName = $"BaoCao_ThongKe_NhanSu_{DateTime.Now:yyyyMMdd}.xlsx"
                    };

                    if (saveFileDialog.ShowDialog() != true) return;

                    using (var workbook = new XLWorkbook())
                    {
                        // --- Sheet 1: Thống kê tổng hợp ---
                        var wsSummary = workbook.Worksheets.Add("Thống kê chung");

                        // Tiêu đề báo cáo
                        var titleRange = wsSummary.Range("A1:C1");
                        titleRange.Merge();
                        titleRange.Value = "BÁO CÁO THỐNG KÊ NHÂN SỰ";
                        titleRange.Style.Font.Bold = true;
                        titleRange.Style.Font.FontSize = 16;
                        titleRange.Style.Font.FontColor = XLColor.FromHtml("#1565C0");
                        titleRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        titleRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        wsSummary.Row(1).Height = 35;

                        // Thông tin bộ lọc
                        wsSummary.Cell(2, 1).Value = "Bộ phận:";
                        wsSummary.Cell(2, 1).Style.Font.Bold = true;
                        wsSummary.Cell(2, 2).Value = departmentStr;

                        wsSummary.Cell(3, 1).Value = "Năm thống kê:";
                        wsSummary.Cell(3, 1).Style.Font.Bold = true;
                        wsSummary.Cell(3, 2).Value = yearStr;



                        // Header bảng thống kê
                        int startRow = 6;
                        string[] summaryHeaders = { "Chỉ số thống kê", "Số lượng (người)", "Tỷ lệ (%)" };
                        for (int col = 1; col <= summaryHeaders.Length; col++)
                        {
                            var cell = wsSummary.Cell(startRow, col);
                            cell.Value = summaryHeaders[col - 1];
                            cell.Style.Font.Bold = true;
                            cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#1565C0");
                            cell.Style.Font.FontColor = XLColor.White;
                            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        }

                        // Dữ liệu thống kê theo nhóm
                        var rankRows = CalculateRankStatistics(list, db, total);
                        var teamRows = CalculateTeamLeadershipRows(list);

                        var rowData = new List<(string Name, double Count, double Ratio, bool IsGroup)>
                        {
                            ("TỔNG SỐ CÔNG CHỨC", total, 100.0, false),
                            ("Nam công chức", male, maleRatio, false),
                            ("Nữ công chức", female, femaleRatio, false),

                            ("TRÌNH ĐỘ HỌC VẤN", 0, 0, true),
                            ("Sau Đại học (Thạc sĩ, Tiến sĩ)", sauDaiHoc, sauDaiHocRatio, false),
                            ("Đại học (Cử nhân)", daiHoc, daiHocRatio, false),
                            ("Trình độ khác", khacEdu, khacEduRatio, false),

                            ("PHÂN BỐ ĐẢNG VIÊN", 0, 0, true),
                            ("Đảng viên", party, partyRatio, false),
                            ("Chưa là Đảng viên", nonParty, nonPartyRatio, false),

                            ("PHÂN BỐ ĐỘ TUỔI", 0, 0, true),
                            ("Dưới 30 tuổi", age1, age1Ratio, false),
                            ("Từ 30 - 40 tuổi", age2, age2Ratio, false),
                            ("Từ 41 - 50 tuổi", age3, age3Ratio, false),
                            ("Trên 50 tuổi", age4, age4Ratio, false),

                            ("THÂM NIÊN CÔNG TÁC", 0, 0, true),
                            ("Dưới 5 năm", sen1, sen1Ratio, false),
                            ("Từ 5 - 10 năm", sen2, sen2Ratio, false),
                            ("Từ 11 - 20 năm", sen3, sen3Ratio, false),
                            ("Trên 20 năm", sen4, sen4Ratio, false),

                            ("CÁC NGẠCH CÔNG CHỨC", 0, 0, true)
                        };

                        foreach (var r in rankRows)
                        {
                            string label = string.IsNullOrEmpty(r.RankCode) || r.RankCode == "---"
                                ? r.RankName
                                : $"{r.RankName} ({r.RankCode})";
                            rowData.Add((label, r.TotalCount, r.Ratio, false));
                        }

                        int curRow = startRow + 1;
                        foreach (var data in rowData)
                        {
                            if (data.IsGroup)
                            {
                                // Dòng tiêu đề nhóm
                                var range = wsSummary.Range(wsSummary.Cell(curRow, 1), wsSummary.Cell(curRow, 3));
                                range.Merge();
                                range.Value = data.Name;
                                range.Style.Font.Bold = true;
                                range.Style.Fill.BackgroundColor = XLColor.FromHtml("#E3F2FD");
                                range.Style.Font.FontColor = XLColor.FromHtml("#1565C0");
                                range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            }
                            else
                            {
                                // Dòng dữ liệu thông thường
                                wsSummary.Cell(curRow, 1).Value = data.Name;
                                wsSummary.Cell(curRow, 2).Value = data.Count;
                                wsSummary.Cell(curRow, 3).Value = data.Ratio;

                                wsSummary.Cell(curRow, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                                var cellCount = wsSummary.Cell(curRow, 2);
                                cellCount.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                cellCount.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                cellCount.Style.NumberFormat.Format = "#,##0";

                                var cellRatio = wsSummary.Cell(curRow, 3);
                                cellRatio.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                cellRatio.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                cellRatio.Style.NumberFormat.Format = "0.0";

                                // In đậm dòng tổng số
                                if (data.Name == "TỔNG SỐ CÔNG CHỨC")
                                {
                                    wsSummary.Cell(curRow, 1).Style.Font.Bold = true;
                                    cellCount.Style.Font.Bold = true;
                                    cellRatio.Style.Font.Bold = true;
                                }
                            }
                            curRow++;
                        }

                        wsSummary.Column(1).Width = 42;
                        wsSummary.Column(2).Width = 20;
                        wsSummary.Column(3).Width = 15;

                        // --- Sheet 2: Thống kê ngạch công chức ---
                        var wsRank = workbook.Worksheets.Add("Thống kê ngạch");

                        var rankTitleRange = wsRank.Range("A1:G1");
                        rankTitleRange.Merge();
                        rankTitleRange.Value = "BẢNG THỐNG KÊ CHI TIẾT CÁC NGẠCH CÔNG CHỨC";
                        rankTitleRange.Style.Font.Bold = true;
                        rankTitleRange.Style.Font.FontSize = 16;
                        rankTitleRange.Style.Font.FontColor = XLColor.FromHtml("#4338CA");
                        rankTitleRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        rankTitleRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        wsRank.Row(1).Height = 35;

                        wsRank.Cell(2, 1).Value = "Bộ phận:";
                        wsRank.Cell(2, 1).Style.Font.Bold = true;
                        wsRank.Cell(2, 2).Value = departmentStr;

                        wsRank.Cell(3, 1).Value = "Năm thống kê:";
                        wsRank.Cell(3, 1).Style.Font.Bold = true;
                        wsRank.Cell(3, 2).Value = yearStr;

                        string[] rankHeaders = { "STT", "Mã ngạch", "Tên ngạch công chức", "Nam", "Nữ", "Tổng số (người)", "Tỷ lệ (%)" };
                        for (int col = 1; col <= rankHeaders.Length; col++)
                        {
                            var cell = wsRank.Cell(5, col);
                            cell.Value = rankHeaders[col - 1];
                            cell.Style.Font.Bold = true;
                            cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#4338CA");
                            cell.Style.Font.FontColor = XLColor.White;
                            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        }
                        wsRank.Row(5).Height = 25;

                        int rRow = 6;
                        foreach (var r in rankRows)
                        {
                            wsRank.Cell(rRow, 1).Value = r.STT;
                            wsRank.Cell(rRow, 2).Value = r.RankCode;
                            wsRank.Cell(rRow, 3).Value = r.RankName;
                            wsRank.Cell(rRow, 4).Value = r.MaleCount;
                            wsRank.Cell(rRow, 5).Value = r.FemaleCount;
                            wsRank.Cell(rRow, 6).Value = r.TotalCount;
                            wsRank.Cell(rRow, 7).Value = r.Ratio;

                            for (int col = 1; col <= 7; col++)
                            {
                                var cell = wsRank.Cell(rRow, col);
                                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                                if (col != 3)
                                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                if (col == 4 || col == 5 || col == 6)
                                    cell.Style.NumberFormat.Format = "#,##0";
                                else if (col == 7)
                                    cell.Style.NumberFormat.Format = "0.0";
                            }
                            rRow++;
                        }

                        // Dòng tổng cộng bảng ngạch
                        wsRank.Cell(rRow, 1).Value = "";
                        wsRank.Cell(rRow, 2).Value = "";
                        wsRank.Cell(rRow, 3).Value = "TỔNG CỘNG";
                        wsRank.Cell(rRow, 3).Style.Font.Bold = true;
                        wsRank.Cell(rRow, 4).Value = rankRows.Sum(x => x.MaleCount);
                        wsRank.Cell(rRow, 4).Style.Font.Bold = true;
                        wsRank.Cell(rRow, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        wsRank.Cell(rRow, 4).Style.NumberFormat.Format = "#,##0";
                        wsRank.Cell(rRow, 5).Value = rankRows.Sum(x => x.FemaleCount);
                        wsRank.Cell(rRow, 5).Style.Font.Bold = true;
                        wsRank.Cell(rRow, 5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        wsRank.Cell(rRow, 5).Style.NumberFormat.Format = "#,##0";
                        wsRank.Cell(rRow, 6).Value = rankRows.Sum(x => x.TotalCount);
                        wsRank.Cell(rRow, 6).Style.Font.Bold = true;
                        wsRank.Cell(rRow, 6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        wsRank.Cell(rRow, 6).Style.NumberFormat.Format = "#,##0";
                        wsRank.Cell(rRow, 7).Value = total > 0 ? 100.0 : 0.0;
                        wsRank.Cell(rRow, 7).Style.Font.Bold = true;
                        wsRank.Cell(rRow, 7).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        wsRank.Cell(rRow, 7).Style.NumberFormat.Format = "0.0";

                        for (int col = 1; col <= 7; col++)
                        {
                            var cell = wsRank.Cell(rRow, col);
                            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#EEF2FF");
                        }

                        wsRank.Columns().AdjustToContents();

                        // --- Sheet 3: Thống kê lãnh đạo tổ ---
                        var wsTeam = workbook.Worksheets.Add("Lãnh đạo tổ");

                        var teamTitleRange = wsTeam.Range("A1:F1");
                        teamTitleRange.Merge();
                        teamTitleRange.Value = "BẢNG THỐNG KÊ LÃNH ĐẠO TỔ THEO BỘ PHẬN";
                        teamTitleRange.Style.Font.Bold = true;
                        teamTitleRange.Style.Font.FontSize = 16;
                        teamTitleRange.Style.Font.FontColor = XLColor.FromHtml("#0F766E");
                        teamTitleRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        teamTitleRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        wsTeam.Row(1).Height = 35;

                        wsTeam.Cell(2, 1).Value = "Bộ phận:";
                        wsTeam.Cell(2, 1).Style.Font.Bold = true;
                        wsTeam.Cell(2, 2).Value = departmentStr;

                        wsTeam.Cell(3, 1).Value = "Năm thống kê:";
                        wsTeam.Cell(3, 1).Style.Font.Bold = true;
                        wsTeam.Cell(3, 2).Value = yearStr;

                        string[] teamHeaders = { "STT", "Bộ phận / Tổ", "Tổ trưởng", "Phó Tổ trưởng", "Công chức", "Tổng cộng" };
                        for (int col = 1; col <= teamHeaders.Length; col++)
                        {
                            var cell = wsTeam.Cell(5, col);
                            cell.Value = teamHeaders[col - 1];
                            cell.Style.Font.Bold = true;
                            cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#0F766E");
                            cell.Style.Font.FontColor = XLColor.White;
                            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        }
                        wsTeam.Row(5).Height = 25;

                        int tRow = 6;
                        int teamStt = 1;
                        foreach (var t in teamRows)
                        {
                            wsTeam.Cell(tRow, 1).Value = teamStt++;
                            wsTeam.Cell(tRow, 2).Value = t.DepartmentName;
                            wsTeam.Cell(tRow, 3).Value = t.LeaderCount;
                            wsTeam.Cell(tRow, 4).Value = t.DeputyCount;
                            wsTeam.Cell(tRow, 5).Value = t.StaffCount;
                            wsTeam.Cell(tRow, 6).Value = t.TotalCount;

                            for (int col = 1; col <= 6; col++)
                            {
                                var cell = wsTeam.Cell(tRow, col);
                                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                                if (col != 2)
                                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                if (col >= 3)
                                    cell.Style.NumberFormat.Format = "#,##0";
                            }
                            tRow++;
                        }

                        // Dòng tổng cộng bảng lãnh đạo tổ
                        wsTeam.Cell(tRow, 1).Value = "";
                        wsTeam.Cell(tRow, 2).Value = "TỔNG CỘNG";
                        wsTeam.Cell(tRow, 2).Style.Font.Bold = true;
                        wsTeam.Cell(tRow, 3).Value = teamRows.Sum(x => x.LeaderCount);
                        wsTeam.Cell(tRow, 3).Style.Font.Bold = true;
                        wsTeam.Cell(tRow, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        wsTeam.Cell(tRow, 3).Style.NumberFormat.Format = "#,##0";
                        wsTeam.Cell(tRow, 4).Value = teamRows.Sum(x => x.DeputyCount);
                        wsTeam.Cell(tRow, 4).Style.Font.Bold = true;
                        wsTeam.Cell(tRow, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        wsTeam.Cell(tRow, 4).Style.NumberFormat.Format = "#,##0";
                        wsTeam.Cell(tRow, 5).Value = teamRows.Sum(x => x.StaffCount);
                        wsTeam.Cell(tRow, 5).Style.Font.Bold = true;
                        wsTeam.Cell(tRow, 5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        wsTeam.Cell(tRow, 5).Style.NumberFormat.Format = "#,##0";
                        wsTeam.Cell(tRow, 6).Value = teamRows.Sum(x => x.TotalCount);
                        wsTeam.Cell(tRow, 6).Style.Font.Bold = true;
                        wsTeam.Cell(tRow, 6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        wsTeam.Cell(tRow, 6).Style.NumberFormat.Format = "#,##0";

                        for (int col = 1; col <= 6; col++)
                        {
                            var cell = wsTeam.Cell(tRow, col);
                            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#F0FDFA");
                        }

                        wsTeam.Columns().AdjustToContents();

                        // --- Sheet 4: Danh sách chi tiết ---
                        var wsDetail = workbook.Worksheets.Add("Danh sách chi tiết");

                        var detailTitleRange = wsDetail.Range("A1:M1");
                        detailTitleRange.Merge();
                        detailTitleRange.Value = "DANH SÁCH CHI TIẾT CÁN BỘ NHÂN SỰ";
                        detailTitleRange.Style.Font.Bold = true;
                        detailTitleRange.Style.Font.FontSize = 16;
                        detailTitleRange.Style.Font.FontColor = XLColor.FromHtml("#1565C0");
                        detailTitleRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        detailTitleRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        wsDetail.Row(1).Height = 35;

                        // Header bảng chi tiết
                        string[] headers = {
                            "STT", "Mã cán bộ", "Họ và tên", "Ngày sinh", "Giới tính",
                            "Bộ phận công tác", "Chức vụ", "Mã ngạch", "Tên ngạch", "Trình độ học vấn",
                            "Đảng viên", "Ngày vào Đảng", "Ngày bắt đầu công tác"
                        };

                        for (int i = 0; i < headers.Length; i++)
                        {
                            var cell = wsDetail.Cell(3, i + 1);
                            cell.Value = headers[i];
                            cell.Style.Font.Bold = true;
                            cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#1565C0");
                            cell.Style.Font.FontColor = XLColor.White;
                            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        }
                        wsDetail.Row(3).Height = 25;

                        // Điền dữ liệu nhân sự (sắp xếp theo Bộ phận rồi Họ tên)
                        int idx = 1, rNum = 4;
                        foreach (var p in list.OrderBy(x => x.Department).ThenBy(x => x.FullName))
                        {
                            DateTime? start = p.TaxAuthorityStartDate ?? p.StartDate;

                            wsDetail.Cell(rNum, 1).Value = idx++;
                            wsDetail.Cell(rNum, 2).Value = "'" + (p.StaffId ?? "");
                            wsDetail.Cell(rNum, 3).Value = p.FullName;
                            wsDetail.Cell(rNum, 4).Value = p.DateOfBirth.HasValue ? p.DateOfBirth.Value.ToString("dd/MM/yyyy") : "";
                            wsDetail.Cell(rNum, 5).Value = p.Gender ?? "";
                            wsDetail.Cell(rNum, 6).Value = p.Department ?? "";
                            wsDetail.Cell(rNum, 7).Value = p.Position ?? "";
                            wsDetail.Cell(rNum, 8).Value = p.RankCode ?? "";
                            wsDetail.Cell(rNum, 9).Value = p.RankName ?? "";
                            wsDetail.Cell(rNum, 10).Value = p.EducationLevel ?? "";
                            wsDetail.Cell(rNum, 11).Value = p.PartyEntryDate.HasValue ? "Đảng viên" : "";
                            wsDetail.Cell(rNum, 12).Value = p.PartyEntryDate.HasValue ? p.PartyEntryDate.Value.ToString("dd/MM/yyyy") : "";
                            wsDetail.Cell(rNum, 13).Value = start.HasValue ? start.Value.ToString("dd/MM/yyyy") : "";

                            // Định dạng viền và căn lề
                            for (int col = 1; col <= headers.Length; col++)
                            {
                                var cell = wsDetail.Cell(rNum, col);
                                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                                // Căn giữa tất cả trừ các cột text dài
                                if (col != 3 && col != 6 && col != 7 && col != 9 && col != 10)
                                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            }
                            rNum++;
                        }

                        // Tự động điều chỉnh độ rộng cột
                        wsDetail.Columns().AdjustToContents();

                        workbook.SaveAs(saveFileDialog.FileName);

                        // Thông báo xuất thành công
                        var success = new SuccessWindow("Xuất báo cáo Excel thành công!", null, saveFileDialog.FileName, true);
                        if (Window.GetWindow(this) is Window parent)
                            success.Owner = parent;
                        success.ShowDialog();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi xuất Excel: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                App.DebugLog("Lỗi xuất Excel thống kê: " + ex.Message);
            }
        }
    }

    // Model đại diện cho một lát cắt của biểu đồ tròn
    public class PieSlice
    {
        public int Value { get; set; }
        public SolidColorBrush ColorBrush { get; set; } = new();
    }

    /// <summary>DTO cho một dòng trong bảng thống kê các ngạch công chức.</summary>
    public class RankStatisticsRow
    {
        public int STT { get; set; }
        public string RankCode { get; set; } = string.Empty;
        public string RankName { get; set; } = string.Empty;
        public int MaleCount { get; set; }
        public int FemaleCount { get; set; }
        public int TotalCount { get; set; }
        public double Ratio { get; set; }
        public string RatioText { get; set; } = string.Empty;
        public double BarRatio { get; set; }
        public bool IsUnassigned { get; set; }
    }

    /// <summary>DTO cho một dòng trong bảng thống kê lãnh đạo tổ.</summary>
    public class TeamLeadershipRow
    {
        public string DepartmentName { get; set; } = string.Empty;
        public int LeaderCount { get; set; }
        public int DeputyCount { get; set; }
        public int StaffCount { get; set; }
        public int TotalCount { get; set; }
    }
}
