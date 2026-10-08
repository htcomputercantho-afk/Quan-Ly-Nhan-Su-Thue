using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using MaterialDesignThemes.Wpf;
using TaxPersonnelManagement.Models;
using TaxPersonnelManagement.Services;

namespace TaxPersonnelManagement.Views
{
    /// <summary>
    /// ViewModel của từng ô tháng (T1..T12) trong ma trận phân công nhiệm vụ.
    /// </summary>
    public class DutyCellViewModel
    {
        public int Month { get; set; }
        public string Department { get; set; } = string.Empty;
        public bool HasFile { get; set; }
        public bool IsCarryOver { get; set; }
        public bool IsEmpty { get; set; }

        public Visibility HasFileVisibility => HasFile ? Visibility.Visible : Visibility.Collapsed;
        public Visibility CarryOverVisibility => IsCarryOver ? Visibility.Visible : Visibility.Collapsed;
        public Visibility EmptyVisibility => IsEmpty ? Visibility.Visible : Visibility.Collapsed;

        public string DisplayBadge { get; set; } = string.Empty;
        public string TooltipText { get; set; } = string.Empty;
        public int? PrimaryId { get; set; }

        public Brush CellBackground { get; set; } = Brushes.Transparent;
    }

    /// <summary>
    /// ViewModel của một dòng Tổ / bộ phận trong ma trận.
    /// </summary>
    public class DutyRowViewModel
    {
        public string Department { get; set; } = string.Empty;
        public string Subtitle { get; set; } = string.Empty;
        public string Abbreviation { get; set; } = "TỔ";
        public Brush AvatarBackground { get; set; } = Brushes.LightGray;
        public Brush AvatarForeground { get; set; } = Brushes.Black;

        public List<DutyCellViewModel> Cells { get; set; } = new();

        public int ChangeCountInYear { get; set; }
        public int TotalRecords { get; set; }
        public string StatusCode { get; set; } = "All"; // "Stable", "Multiple", "NoFile"

        public string StatusText { get; set; } = string.Empty;
        public PackIconKind StatusIcon { get; set; } = PackIconKind.CheckCircleOutline;
        public Brush StatusForeground { get; set; } = Brushes.Gray;
        public Brush StatusBackground { get; set; } = Brushes.Transparent;
        public Brush StatusBorder { get; set; } = Brushes.Transparent;

        public Visibility CanEditVisibility { get; set; } = Visibility.Visible;
    }

    /// <summary>
    /// Trang quản lý và theo dõi Phân công nhiệm vụ theo ma trận Tổ x Tháng.
    /// </summary>
    public partial class DutyAssignmentView : Page
    {
        private readonly bool _canEdit;
        private bool _suppressEvents = true;
        private List<DutyRowViewModel> _allRows = new();
        private string _currentStatusFilter = "All";

        private static readonly Brush ActiveBlueBorder = new SolidColorBrush(Color.FromRgb(37, 99, 235)); // #2563EB
        private static readonly Brush ActiveBlueBg = new SolidColorBrush(Color.FromArgb(18, 37, 99, 235));
        private static readonly Brush ActiveEmeraldBorder = new SolidColorBrush(Color.FromRgb(5, 150, 105)); // #059669
        private static readonly Brush ActiveEmeraldBg = new SolidColorBrush(Color.FromArgb(18, 5, 150, 105));
        private static readonly Brush ActiveAmberBorder = new SolidColorBrush(Color.FromRgb(217, 119, 6)); // #D97706
        private static readonly Brush ActiveAmberBg = new SolidColorBrush(Color.FromArgb(18, 217, 119, 6));
        private static readonly Brush ActiveSlateBorder = new SolidColorBrush(Color.FromRgb(71, 85, 105)); // #475569
        private static readonly Brush ActiveSlateBg = new SolidColorBrush(Color.FromArgb(18, 71, 85, 105));

        private static readonly Brush InactiveBorder = new SolidColorBrush(Color.FromRgb(226, 232, 240)); // #E2E8F0
        private static readonly Brush InactiveBg = Brushes.White;

        public DutyAssignmentView()
        {
            InitializeComponent();
            _canEdit = App.CurrentUser?.Role != UserRole.Staff;
            if (!_canEdit)
            {
                btnUploadTop.Visibility = Visibility.Collapsed;
            }

            InitYears();
            _suppressEvents = false;
            LoadData();
        }

        private void InitYears()
        {
            int currentYear = DateTime.Today.Year;
            cboYear.Items.Clear();
            for (int y = currentYear - 3; y <= currentYear + 2; y++)
            {
                cboYear.Items.Add(y);
            }
            cboYear.SelectedItem = currentYear;
        }

        private int GetSelectedYear()
        {
            if (cboYear.SelectedItem is int y) return y;
            return DateTime.Today.Year;
        }

        private static (string Abbr, Brush Bg, Brush Fg) GetDepartmentAvatar(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return ("TỔ", new SolidColorBrush(Color.FromRgb(241, 245, 249)), new SolidColorBrush(Color.FromRgb(100, 116, 139)));

            string lower = name.ToLowerInvariant();
            if (lower.Contains("lãnh đạo"))
                return ("BLĐ", new SolidColorBrush(Color.FromRgb(219, 234, 254)), new SolidColorBrush(Color.FromRgb(29, 78, 216)));
            if (lower.Contains("hành chính") || lower.Contains("tổng hợp"))
                return ("HC", new SolidColorBrush(Color.FromRgb(204, 251, 241)), new SolidColorBrush(Color.FromRgb(15, 118, 110)));
            if (lower.Contains("kiểm tra 1") || lower.Contains("số 1"))
                return ("KT1", new SolidColorBrush(Color.FromRgb(224, 231, 255)), new SolidColorBrush(Color.FromRgb(67, 56, 202)));
            if (lower.Contains("kiểm tra 2") || lower.Contains("số 2"))
                return ("KT2", new SolidColorBrush(Color.FromRgb(224, 231, 255)), new SolidColorBrush(Color.FromRgb(67, 56, 202)));
            if (lower.Contains("kiểm tra 3") || lower.Contains("số 3"))
                return ("KT3", new SolidColorBrush(Color.FromRgb(224, 231, 255)), new SolidColorBrush(Color.FromRgb(67, 56, 202)));
            if (lower.Contains("kiểm tra 4") || lower.Contains("số 4"))
                return ("KT4", new SolidColorBrush(Color.FromRgb(224, 231, 255)), new SolidColorBrush(Color.FromRgb(67, 56, 202)));
            if (lower.Contains("nghiệp vụ") || lower.Contains("pháp chế") || lower.Contains("dự toán"))
                return ("NV", new SolidColorBrush(Color.FromRgb(243, 232, 255)), new SolidColorBrush(Color.FromRgb(126, 34, 206)));
            if (lower.Contains("khoản thu khác"))
                return ("TN", new SolidColorBrush(Color.FromRgb(255, 228, 230)), new SolidColorBrush(Color.FromRgb(190, 18, 60)));
            if (lower.Contains("cá nhân") || lower.Contains("hộ kinh"))
                return ("CN", new SolidColorBrush(Color.FromRgb(254, 243, 199)), new SolidColorBrush(Color.FromRgb(180, 83, 9)));
            if (lower.Contains("doanh nghiệp"))
                return ("DN", new SolidColorBrush(Color.FromRgb(254, 215, 170)), new SolidColorBrush(Color.FromRgb(194, 65, 12)));

            // Rút gọn từ các chữ cái đầu
            var words = name.Split(new[] { ' ', ',', '-' }, StringSplitOptions.RemoveEmptyEntries);
            string initials = words.Length >= 2
                ? (words[0][0].ToString() + words[1][0].ToString()).ToUpperInvariant()
                : (words.Length == 1 ? words[0].Substring(0, Math.Min(2, words[0].Length)).ToUpperInvariant() : "TỔ");

            return (initials, new SolidColorBrush(Color.FromRgb(241, 245, 249)), new SolidColorBrush(Color.FromRgb(71, 85, 105)));
        }

        /// <summary>
        /// Nạp lại toàn bộ dữ liệu phân công nhiệm vụ và xây dựng ma trận.
        /// </summary>
        public void LoadData()
        {
            int selectedYear = GetSelectedYear();
            int currentYear = DateTime.Today.Year;
            int currentMonth = DateTime.Today.Month;

            var departments = DutyAssignmentService.GetDepartments();
            var allVersions = DutyAssignmentService.GetVersions(); // Toàn bộ lịch sử sắp xếp mới nhất trước

            // Nhóm theo Tổ
            var byDept = allVersions.GroupBy(v => v.Department, StringComparer.OrdinalIgnoreCase)
                                    .ToDictionary(g => g.Key, g => g.ToList(), StringComparer.OrdinalIgnoreCase);

            var rows = new List<DutyRowViewModel>();


            foreach (var dept in departments)
            {
                byDept.TryGetValue(dept, out var deptVersions);
                deptVersions ??= new List<DutyVersionItem>();

                // Các record trong năm đang xem
                var inYear = deptVersions.Where(v => v.Year == selectedYear).ToList();
                // Toàn bộ các record trước năm đang xem
                var beforeYear = deptVersions.Where(v => v.Year < selectedYear)
                                             .OrderByDescending(v => v.EffectiveDate)
                                             .ToList();

                var (abbr, bg, fg) = GetDepartmentAvatar(dept);

                var row = new DutyRowViewModel
                {
                    Department = dept,
                    Abbreviation = abbr,
                    AvatarBackground = bg,
                    AvatarForeground = fg,
                    ChangeCountInYear = inYear.Count,
                    TotalRecords = deptVersions.Count,
                    CanEditVisibility = _canEdit ? Visibility.Visible : Visibility.Collapsed
                };

                // Xác định phiên bản đang có hiệu lực / gần nhất cho Subtitle
                var activeNow = deptVersions.FirstOrDefault(v => v.IsCurrent);
                var latest = deptVersions.FirstOrDefault();
                if (activeNow != null)
                {
                    row.Subtitle = $"Hiện hành: {activeNow.DisplayTitle} ({activeNow.EffectiveDate:dd/MM/yyyy})";
                }
                else if (latest != null)
                {
                    row.Subtitle = $"Văn bản gần nhất: {latest.DisplayTitle} ({latest.EffectiveDate:dd/MM/yyyy})";
                }
                else
                {
                    row.Subtitle = "Chưa cập nhật văn bản phân công";
                }

                // Xây dựng 12 ô tháng (T1..T12)
                for (int m = 1; m <= 12; m++)
                {
                    var cell = new DutyCellViewModel
                    {
                        Month = m,
                        Department = dept
                    };

                    // Highlight nhẹ tháng hiện tại nếu đang xem năm nay
                    if (selectedYear == currentYear && m == currentMonth)
                    {
                        cell.CellBackground = new SolidColorBrush(Color.FromArgb(28, 59, 130, 246)); // Blue tint 11%
                    }

                    // Tìm các record ban hành trong tháng này của năm đang xem
                    var inMonth = inYear.Where(v => v.Month == m)
                                        .OrderByDescending(v => v.EffectiveDate)
                                        .ToList();

                    if (inMonth.Count > 0)
                    {
                        // THÁNG CÓ FILE TẢI LÊN
                        cell.HasFile = true;
                        var primary = inMonth[0];
                        cell.PrimaryId = primary.Id;
                        cell.DisplayBadge = inMonth.Count == 1 ? primary.EffectiveDate.ToString("dd/MM") : $"x{inMonth.Count}";

                        var details = string.Join("\n", inMonth.Select(r => $"• {r.DisplayTitle} (Hiệu lực: {r.EffectiveDate:dd/MM/yyyy}) - {r.FileName}"));
                        cell.TooltipText = $"[Tháng {m}/{selectedYear}] Có {inMonth.Count} phân công nhiệm vụ:\n{details}\n\n👉 Bấm để xem PDF trực tiếp hoặc tải về";
                    }
                    else
                    {
                        // Không có file tải lên ở tháng này -> Kiểm tra xem có văn bản nào có hiệu lực trước đó không
                        var priorInYear = inYear.Where(v => v.Month < m)
                                                .OrderByDescending(v => v.Month)
                                                .ThenByDescending(v => v.EffectiveDate)
                                                .FirstOrDefault();

                        var effectivePrior = priorInYear ?? beforeYear.FirstOrDefault();

                        if (effectivePrior != null)
                        {
                            // GIỮ NGUYÊN (Kế thừa từ văn bản trước)
                            cell.IsCarryOver = true;
                            cell.PrimaryId = effectivePrior.Id;
                            cell.TooltipText = $"[Tháng {m}/{selectedYear}] Giữ nguyên nhiệm vụ theo văn bản trước:\n• {effectivePrior.DisplayTitle} (Hiệu lực: {effectivePrior.EffectiveDate:dd/MM/yyyy})\n\n👉 Bấm để xem văn bản đang áp dụng";
                        }
                        else
                        {
                            // CHƯA TỪNG CÓ FILE NÀO TRƯỚC ĐÓ
                            cell.IsEmpty = true;
                            cell.TooltipText = $"[Tháng {m}/{selectedYear}] Chưa có văn bản phân công nào";
                        }
                    }

                    row.Cells.Add(cell);
                }

                // Tính toán Trạng thái & Màu sắc của dòng dựa trên năm đang xem
                bool hasInYear = inYear.Count > 0;
                bool hasPrior = beforeYear.Count > 0;

                if (!hasInYear && !hasPrior)
                {
                    // Chưa từng có văn bản nào tính đến năm đang xem
                    row.StatusCode = "NoFile";
                    row.StatusText = "Chưa có file";
                    row.StatusIcon = PackIconKind.FileQuestionOutline;
                    row.StatusForeground = new SolidColorBrush(Color.FromRgb(71, 85, 105)); // #475569
                    row.StatusBackground = new SolidColorBrush(Color.FromRgb(241, 245, 249)); // #F1F5F9
                    row.StatusBorder = new SolidColorBrush(Color.FromRgb(226, 232, 240)); // #E2E8F0
                }
                else if (!hasInYear && hasPrior)
                {
                    // Năm nay không có văn bản mới nhưng có văn bản từ năm trước -> Ổn định (kế thừa/giữ nguyên)
                    row.StatusCode = "Stable";
                    row.StatusText = "Ổn định (giữ nguyên)";
                    row.StatusIcon = PackIconKind.BookmarkCheckOutline;
                    row.StatusForeground = new SolidColorBrush(Color.FromRgb(3, 105, 161)); // Sky 700
                    row.StatusBackground = new SolidColorBrush(Color.FromRgb(224, 242, 254)); // Sky 100
                    row.StatusBorder = new SolidColorBrush(Color.FromRgb(186, 230, 253)); // Sky 200
                }
                else if (inYear.Count == 1)
                {
                    // Có 1 lần ban hành trong năm -> Ổn định
                    row.StatusCode = "Stable";
                    row.StatusText = "Ổn định · 1 lần";
                    row.StatusIcon = PackIconKind.CheckCircleOutline;
                    row.StatusForeground = new SolidColorBrush(Color.FromRgb(6, 95, 70)); // Emerald 800
                    row.StatusBackground = new SolidColorBrush(Color.FromRgb(209, 250, 229)); // Emerald 100
                    row.StatusBorder = new SolidColorBrush(Color.FromRgb(167, 243, 208)); // Emerald 200
                }
                else
                {
                    // Thay đổi từ 2 lần trở lên -> Nhiều thay đổi
                    row.StatusCode = "Multiple";
                    row.StatusText = $"{inYear.Count} lần thay đổi";
                    row.StatusIcon = PackIconKind.AlertCircleOutline;
                    row.StatusForeground = new SolidColorBrush(Color.FromRgb(146, 64, 14)); // Amber 800
                    row.StatusBackground = new SolidColorBrush(Color.FromRgb(254, 243, 199)); // Amber 100
                    row.StatusBorder = new SolidColorBrush(Color.FromRgb(253, 230, 138)); // Amber 200
                }

                rows.Add(row);
            }

            _allRows = rows;
            UpdateKpiCardVisuals();
            ApplyFilter();
        }

        /// <summary>
        /// Cập nhật giá trị số và tỷ lệ phần trăm trên 4 thẻ KPI dựa trên phạm vi dữ liệu đang lọc.
        /// </summary>
        private void UpdateKpiCards(List<DutyRowViewModel> scopeRows, bool isFilteredBySearch)
        {
            int total = scopeRows.Count;
            int stable = scopeRows.Count(r => r.StatusCode == "Stable");
            int multiple = scopeRows.Count(r => r.StatusCode == "Multiple");
            int noFile = scopeRows.Count(r => r.StatusCode == "NoFile");

            txtKpiTotal.Text = total.ToString();
            txtKpiStable.Text = stable.ToString();
            txtKpiMultiple.Text = multiple.ToString();
            txtKpiNoFile.Text = noFile.ToString();

            if (isFilteredBySearch)
            {
                txtKpiTotalPercent.Text = "Khớp";
                txtKpiTotalDesc.Text = "Theo từ khóa";
            }
            else
            {
                txtKpiTotalPercent.Text = "100%";
                txtKpiTotalDesc.Text = "Đang theo dõi";
            }

            txtKpiStablePercent.Text = total > 0 ? $"{Math.Round((double)stable / total * 100)}%" : "0%";
            txtKpiMultiplePercent.Text = total > 0 ? $"{Math.Round((double)multiple / total * 100)}%" : "0%";
            txtKpiNoFilePercent.Text = total > 0 ? $"{Math.Round((double)noFile / total * 100)}%" : "0%";
        }

        /// <summary>
        /// Cập nhật viền, màu nền và badge trạng thái đang lọc trên 4 thẻ KPI để người dùng dễ nhận biết.
        /// </summary>
        private void UpdateKpiCardVisuals()
        {
            bool isAll = _currentStatusFilter == "All";
            bool isStable = _currentStatusFilter == "Stable";
            bool isMultiple = _currentStatusFilter == "Multiple";
            bool isNoFile = _currentStatusFilter == "NoFile";

            // Card 1: Tổng số Tổ
            cardKpiTotal.BorderBrush = isAll ? ActiveBlueBorder : InactiveBorder;
            cardKpiTotal.BorderThickness = new Thickness(isAll ? 2 : 1);
            cardKpiTotal.Background = isAll ? ActiveBlueBg : InactiveBg;
            badgeKpiTotalActive.Visibility = isAll ? Visibility.Visible : Visibility.Collapsed;

            // Card 2: Ổn định
            cardKpiStable.BorderBrush = isStable ? ActiveEmeraldBorder : InactiveBorder;
            cardKpiStable.BorderThickness = new Thickness(isStable ? 2 : 1);
            cardKpiStable.Background = isStable ? ActiveEmeraldBg : InactiveBg;
            badgeKpiStableActive.Visibility = isStable ? Visibility.Visible : Visibility.Collapsed;

            // Card 3: Nhiều thay đổi
            cardKpiMultiple.BorderBrush = isMultiple ? ActiveAmberBorder : InactiveBorder;
            cardKpiMultiple.BorderThickness = new Thickness(isMultiple ? 2 : 1);
            cardKpiMultiple.Background = isMultiple ? ActiveAmberBg : InactiveBg;
            badgeKpiMultipleActive.Visibility = isMultiple ? Visibility.Visible : Visibility.Collapsed;

            // Card 4: Chưa có văn bản
            cardKpiNoFile.BorderBrush = isNoFile ? ActiveSlateBorder : InactiveBorder;
            cardKpiNoFile.BorderThickness = new Thickness(isNoFile ? 2 : 1);
            cardKpiNoFile.Background = isNoFile ? ActiveSlateBg : InactiveBg;
            badgeKpiNoFileActive.Visibility = isNoFile ? Visibility.Visible : Visibility.Collapsed;
        }

        /// <summary>
        /// Chuyển đổi bộ lọc trạng thái và đồng bộ cả 4 thẻ KPI lẫn thanh tab RadioButton.
        /// </summary>
        private void SetStatusFilter(string filter)
        {
            _currentStatusFilter = filter;

            _suppressEvents = true;
            rbFilterAll.IsChecked = filter == "All";
            rbFilterStable.IsChecked = filter == "Stable";
            rbFilterMultiple.IsChecked = filter == "Multiple";
            rbFilterNoFile.IsChecked = filter == "NoFile";
            _suppressEvents = false;

            UpdateKpiCardVisuals();
            ApplyFilter();
        }

        private void ApplyFilter()
        {
            string keyword = txtSearch.Text?.Trim() ?? string.Empty;
            btnClearSearch.Visibility = string.IsNullOrEmpty(keyword) ? Visibility.Collapsed : Visibility.Visible;

            // 1. Lọc theo từ khóa tìm kiếm Tổ
            var searchScope = _allRows.AsEnumerable();
            if (!string.IsNullOrEmpty(keyword))
            {
                searchScope = searchScope.Where(r => r.Department.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0);
            }
            var scopeList = searchScope.ToList();

            // Cập nhật số liệu trên 4 thẻ KPI theo phạm vi tìm kiếm (hoặc toàn bộ năm nếu không tìm kiếm)
            UpdateKpiCards(scopeList, !string.IsNullOrEmpty(keyword));

            // 2. Lọc tiếp theo trạng thái đã chọn
            var filtered = scopeList.AsEnumerable();
            switch (_currentStatusFilter)
            {
                case "Stable":
                    filtered = filtered.Where(r => r.StatusCode == "Stable");
                    break;
                case "Multiple":
                    filtered = filtered.Where(r => r.StatusCode == "Multiple");
                    break;
                case "NoFile":
                    filtered = filtered.Where(r => r.StatusCode == "NoFile");
                    break;
                case "All":
                default:
                    break;
            }

            var list = filtered.ToList();
            icDepartmentRows.ItemsSource = null;
            icDepartmentRows.ItemsSource = list;
            pnlEmptyState.Visibility = list.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
            txtFilterCount.Text = $"{list.Count} / {_allRows.Count} Tổ";
        }

        private void CboYear_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            // Bấm vào bất kỳ đâu trên ô Năm (icon, chữ 'Năm:', giá trị số, khoảng trống, mũi tên)
            // đều đóng / mở danh sách chọn một cách mượt mà và trực quan
            if (e.OriginalSource is DependencyObject dep)
            {
                // Nếu click xảy ra bên trong popup items (danh sách năm) thì để ComboBox xử lý chọn năm
                DependencyObject? cur = dep;
                while (cur != null)
                {
                    if (cur is ComboBoxItem)
                        return;
                    cur = System.Windows.Media.VisualTreeHelper.GetParent(cur);
                }
            }

            cboYear.IsDropDownOpen = !cboYear.IsDropDownOpen;
            e.Handled = true;
        }

        private void CboYear_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_suppressEvents) return;
            LoadData();
        }

        private void FilterStatus_Checked(object sender, RoutedEventArgs e)
        {
            if (_suppressEvents) return;
            if (sender is RadioButton rb && rb.Tag is string tag)
            {
                SetStatusFilter(tag);
            }
        }

        private void FilterStatus_Click(object sender, RoutedEventArgs e)
        {
            if (_suppressEvents) return;
            if (sender is RadioButton rb && rb.Tag is string tag)
            {
                SetStatusFilter(tag);
            }
        }

        // Bấm vào các thẻ KPI để lọc nhanh danh sách (bấm lại thẻ đang chọn để bỏ lọc về Tất cả)
        private void KpiCardTotal_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            SetStatusFilter("All");
            e.Handled = true;
        }

        private void KpiCardStable_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            SetStatusFilter(_currentStatusFilter == "Stable" ? "All" : "Stable");
            e.Handled = true;
        }

        private void KpiCardMultiple_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            SetStatusFilter(_currentStatusFilter == "Multiple" ? "All" : "Multiple");
            e.Handled = true;
        }

        private void KpiCardNoFile_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            SetStatusFilter(_currentStatusFilter == "NoFile" ? "All" : "NoFile");
            e.Handled = true;
        }

        private void TxtSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            ApplyFilter();
        }

        private void BtnClearSearch_Click(object sender, RoutedEventArgs e)
        {
            txtSearch.Text = string.Empty;
        }

        private void BtnRefresh_Click(object sender, RoutedEventArgs e)
        {
            LoadData();
        }

        /// <summary>
        /// Nút "Tải PDF lên" ở thanh trên cùng: mở dialog tải lên cho năm hiện tại.
        /// </summary>
        private void BtnUploadTop_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new DutyAssignmentUploadDialog(null, GetSelectedYear())
            {
                Owner = Window.GetWindow(this)
            };
            if (dlg.ShowDialog() == true)
            {
                DutyAssignmentNotificationDialog.ShowSuccess(Window.GetWindow(this), "Đã lưu văn bản phân công nhiệm vụ mới thành công!", "Tải lên thành công");
                LoadData();
            }
        }

        /// <summary>
        /// Nút "+" trên từng dòng Tổ: mở dialog tải lên cho đúng Tổ đó.
        /// </summary>
        private void BtnRowAdd_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is string dept)
            {
                var dlg = new DutyAssignmentUploadDialog(dept, GetSelectedYear())
                {
                    Owner = Window.GetWindow(this)
                };
                if (dlg.ShowDialog() == true)
                {
                    DutyAssignmentNotificationDialog.ShowSuccess(Window.GetWindow(this), $"Đã lưu văn bản phân công nhiệm vụ cho {dept} thành công!", "Tải lên thành công");
                    LoadData();
                }
            }
        }

        /// <summary>
        /// Nút xem lịch sử trên dòng Tổ: mở cửa sổ xem toàn bộ lịch sử + PDF của Tổ.
        /// </summary>
        private void BtnViewHistory_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is string dept)
            {
                var dlg = new DutyAssignmentViewerDialog(dept)
                {
                    Owner = Window.GetWindow(this)
                };
                dlg.ShowDialog();
                if (dlg.Changed)
                {
                    LoadData();
                }
            }
        }

        /// <summary>
        /// Bấm vào chip PDF trong từng ô tháng: mở trực tiếp cửa sổ xem PDF.
        /// </summary>
        private void BtnCellViewPdf_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is DutyCellViewModel cell)
            {
                var dlg = new DutyAssignmentViewerDialog(cell.Department, cell.PrimaryId)
                {
                    Owner = Window.GetWindow(this)
                };
                dlg.ShowDialog();
                if (dlg.Changed)
                {
                    LoadData();
                }
            }
        }

        /// <summary>
        /// Bấm vào ô giữ nguyên hiệu lực: mở xem văn bản đang áp dụng.
        /// </summary>
        private void CarryOverCell_Click(object sender, MouseButtonEventArgs e)
        {
            if (sender is FrameworkElement el && el.Tag is DutyCellViewModel cell && cell.PrimaryId.HasValue)
            {
                var dlg = new DutyAssignmentViewerDialog(cell.Department, cell.PrimaryId)
                {
                    Owner = Window.GetWindow(this)
                };
                dlg.ShowDialog();
                if (dlg.Changed)
                {
                    LoadData();
                }
            }
        }
    }
}
