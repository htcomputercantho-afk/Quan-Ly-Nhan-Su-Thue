using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using TaxPersonnelManagement.Models;
using TaxPersonnelManagement.Data;
using Microsoft.EntityFrameworkCore;
using ClosedXML.Excel;
using Microsoft.Win32;
using TaxPersonnelManagement.Helpers;

namespace TaxPersonnelManagement.Views
{
    public partial class LeaveDetailView : Page
    {
        private int _activeTab = 0; // 0: Tổng hợp, 1: Nghỉ phép, 2: Nghỉ ốm, 3: Dưỡng sức, 4: Không lương
        private bool _isUpdating = false;
        private bool _isInitialized = false;
        private bool _userHasManuallyToggledFilter = false;

        // Data for Tab 0: Tổng hợp
        private List<LeaveGeneralSummaryItem> _allGeneralSummaries = new List<LeaveGeneralSummaryItem>();
        private List<LeaveGeneralSummaryItem> _filteredGeneralSummaries = new List<LeaveGeneralSummaryItem>();

        // Data for Tab 1: Nghỉ phép
        private List<LeaveSummaryItem> _allAnnualSummaries = new List<LeaveSummaryItem>();
        private List<LeaveSummaryItem> _filteredAnnualSummaries = new List<LeaveSummaryItem>();

        // Data for Tab 2: Nghỉ ốm
        private List<SickLeaveItem> _allSickSummaries = new List<SickLeaveItem>();
        private List<SickLeaveItem> _filteredSickSummaries = new List<SickLeaveItem>();

        // Data for Tab 3: Dưỡng sức
        private List<ConvalescenceLeaveItem> _allConvalescenceSummaries = new List<ConvalescenceLeaveItem>();
        private List<ConvalescenceLeaveItem> _filteredConvalescenceSummaries = new List<ConvalescenceLeaveItem>();

        // Data for Tab 4: Không lương
        private List<UnpaidLeaveItem> _allUnpaidSummaries = new List<UnpaidLeaveItem>();
        private List<UnpaidLeaveItem> _filteredUnpaidSummaries = new List<UnpaidLeaveItem>();

        // Pagination
        private int _currentPage = 1;
        private const int PageSize = 20;
        private int _totalPages = 1;

        public LeaveDetailView()
        {
            _isUpdating = true;
            InitializeComponent();
            InitializeFilters();
            _isInitialized = true;
            _isUpdating = false;

            if (rbTabSummary != null)
            {
                rbTabSummary.IsChecked = true;
            }
            UpdateTabVisibility();
            LoadAllLeaveSummaries();
        }

        private void InitializeFilters()
        {
            _isUpdating = true;

            try
            {
                // 1. Month Filter
                cbMonthFilter.Items.Clear();
                cbMonthFilter.Items.Add("Cả năm");
                for (int i = 1; i <= 12; i++)
                {
                    cbMonthFilter.Items.Add($"Tháng {i}");
                }
                cbMonthFilter.SelectedIndex = 0;

                // 2. Year Filter
                cbYearFilter.Items.Clear();
                int currentYear = DateTime.Now.Year;
                for (int y = 2025; y <= currentYear + 3; y++)
                {
                    cbYearFilter.Items.Add(y);
                }
                cbYearFilter.SelectedItem = currentYear;

                // 3. Department Filter
                cbDepartmentFilter.Items.Clear();
                cbDepartmentFilter.Items.Add("Tất cả bộ phận");
                
                using (var db = new AppDbContext())
                {
                    var departments = db.Departments.Select(d => d.Name).OrderBy(n => n).ToList();
                    foreach (var dept in departments)
                    {
                        cbDepartmentFilter.Items.Add(dept);
                    }
                }
                cbDepartmentFilter.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                App.DebugLog("Error initializing filters: " + ex.Message);
            }
            finally
            {
                _isUpdating = false;
            }
        }

        private void Tab_Checked(object sender, RoutedEventArgs e)
        {
            if (!_isInitialized || _isUpdating) return;

            if (rbTabSummary?.IsChecked == true) _activeTab = 0;
            else if (rbTabAnnual?.IsChecked == true) _activeTab = 1;
            else if (rbTabSick?.IsChecked == true) _activeTab = 2;
            else if (rbTabConvalescence?.IsChecked == true) _activeTab = 3;
            else if (rbTabUnpaid?.IsChecked == true) _activeTab = 4;

            _userHasManuallyToggledFilter = false;
            UpdateTabVisibility();
            UpdateTitle();
            _currentPage = 1;
            ApplyFilter();
        }

        private void UpdateTabVisibility()
        {
            if (dgGeneralSummary == null || dgLeaveSummary == null || dgSickLeave == null || dgConvalescenceLeave == null || dgUnpaidLeave == null) return;

            dgGeneralSummary.Visibility = _activeTab == 0 ? Visibility.Visible : Visibility.Collapsed;
            dgLeaveSummary.Visibility = _activeTab == 1 ? Visibility.Visible : Visibility.Collapsed;
            dgSickLeave.Visibility = _activeTab == 2 ? Visibility.Visible : Visibility.Collapsed;
            dgConvalescenceLeave.Visibility = _activeTab == 3 ? Visibility.Visible : Visibility.Collapsed;
            dgUnpaidLeave.Visibility = _activeTab == 4 ? Visibility.Visible : Visibility.Collapsed;

            if (chkOnlyWithLeaves != null)
            {
                if (_activeTab == 1)
                {
                    chkOnlyWithLeaves.Visibility = Visibility.Collapsed;
                }
                else
                {
                    chkOnlyWithLeaves.Visibility = Visibility.Visible;
                    if (!_userHasManuallyToggledFilter)
                    {
                        // Tab Tổng hợp: mặc định false (hiện tất cả), Tab Nghỉ ốm / Dưỡng sức / Không lương: mặc định true (chỉ hiện người có nghỉ)
                        chkOnlyWithLeaves.IsChecked = (_activeTab == 2 || _activeTab == 3 || _activeTab == 4);
                    }
                }
            }
        }

        private void UpdateTitle()
        {
            if (txtHeaderTitle == null) return;

            int selectedYear = DateTime.Now.Year;
            if (cbYearFilter?.SelectedItem is int yr) selectedYear = yr;

            int? selectedMonth = null;
            if (cbMonthFilter != null && cbMonthFilter.SelectedIndex > 0)
                selectedMonth = cbMonthFilter.SelectedIndex;

            string timeSuffix = selectedMonth.HasValue
                ? $"THÁNG {selectedMonth.Value} NĂM {selectedYear}"
                : $"NĂM {selectedYear}";

            switch (_activeTab)
            {
                case 0:
                    txtHeaderTitle.Text = $"BẢNG TỔNG HỢP NGÀY NGHỈ {timeSuffix}";
                    break;
                case 1:
                    txtHeaderTitle.Text = $"BẢNG TỔNG HỢP SỐ NGÀY NGHỈ PHÉP THỰC TẾ {timeSuffix}";
                    if (colOldYear != null) colOldYear.Header = $"Số ngày phép nghỉ theo chế độ năm {selectedYear - 1}";
                    if (colCurrentYear != null) colCurrentYear.Header = $"Số ngày phép nghỉ theo chế độ năm {selectedYear}";
                    break;
                case 2:
                    txtHeaderTitle.Text = $"BẢNG CHI TIẾT NGHỈ ỐM {timeSuffix}";
                    break;
                case 3:
                    txtHeaderTitle.Text = $"BẢNG CHI TIẾT NGHỈ DƯỠNG SỨC {timeSuffix}";
                    break;
                case 4:
                    txtHeaderTitle.Text = $"BẢNG CHI TIẾT NGHỈ KHÔNG LƯƠNG {timeSuffix}";
                    break;
            }
        }

        private void FilterComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isUpdating) return;
            LoadAllLeaveSummaries();
        }

        private void Filter_Changed(object sender, RoutedEventArgs e)
        {
            if (_isUpdating) return;
            _userHasManuallyToggledFilter = true;
            ApplyFilter();
        }

        private void TxtSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (btnClearSearch != null)
            {
                btnClearSearch.Visibility = string.IsNullOrEmpty(txtSearch.Text) ? Visibility.Collapsed : Visibility.Visible;
            }
            ApplyFilter();
        }

        private void BtnClearSearch_Click(object sender, RoutedEventArgs e)
        {
            txtSearch.Clear();
            txtSearch.Focus();
        }

        private void LoadAllLeaveSummaries()
        {
            if (_isUpdating) return;
            _isUpdating = true;

            try
            {
                int selectedYear = DateTime.Now.Year;
                if (cbYearFilter != null && cbYearFilter.SelectedItem is int yr)
                {
                    selectedYear = yr;
                }

                int? selectedMonth = null;
                if (cbMonthFilter != null && cbMonthFilter.SelectedIndex > 0)
                {
                    selectedMonth = cbMonthFilter.SelectedIndex;
                }

                string? selectedDept = null;
                if (cbDepartmentFilter != null && cbDepartmentFilter.SelectedIndex > 0)
                {
                    selectedDept = cbDepartmentFilter.SelectedItem as string;
                }

                UpdateTitle();

                DateTime rangeStart;
                DateTime rangeEnd;
                if (selectedMonth.HasValue)
                {
                    rangeStart = new DateTime(selectedYear, selectedMonth.Value, 1);
                    rangeEnd = rangeStart.AddMonths(1).AddDays(-1);
                }
                else
                {
                    rangeStart = new DateTime(selectedYear, 1, 1);
                    rangeEnd = new DateTime(selectedYear, 12, 31);
                }

                using (var db = new AppDbContext())
                {
                    var deptOrder = new List<string> {
                        "Ban lãnh đạo",
                        "Tổ Hành chính, tổng hợp",
                        "Tổ Kiểm tra số 1",
                        "Tổ Kiểm tra số 2",
                        "Tổ Kiểm tra số 3",
                        "Tổ Nghiệp vụ, dự toán, pháp chế",
                        "Tổ Quản lý các khoản thu khác",
                        "Tổ Quản lý, hỗ trợ cá nhân, hộ kinh doanh số 1",
                        "Tổ Quản lý, hỗ trợ cá nhân, hộ kinh doanh số 2",
                        "Tổ Quản lý, hỗ trợ doanh nghiệp số 1",
                        "Tổ Quản lý, hỗ trợ doanh nghiệp số 2"
                    };

                    var query = db.Personnel
                                  .AsNoTracking()
                                  .Include(p => p.LeaveHistories)
                                  .Where(p => string.IsNullOrEmpty(p.Status) || p.Status == "Đang công tác")
                                  .AsQueryable();
                    if (!string.IsNullOrEmpty(selectedDept))
                    {
                        query = query.Where(p => p.Department == selectedDept);
                    }

                    var personnelList = query.AsEnumerable()
                                          .OrderBy(p =>
                                          {
                                              string dept = (p.Department ?? "").Trim();
                                              int index = deptOrder.FindIndex(d => d.Equals(dept, StringComparison.OrdinalIgnoreCase));
                                              return index == -1 ? 999 : index;
                                          })
                                          .ThenBy(p =>
                                          {
                                              string pos = p.Position?.ToLower() ?? "";
                                              string dept3 = (p.Department ?? "").ToLower();

                                              if (dept3.Contains("lãnh đạo"))
                                              {
                                                  if (pos.Contains("trưởng") && !pos.Contains("phó") && !pos.Contains("quyền")) return 1;
                                                  if (pos.Contains("quyền")) return 2;
                                                  if (pos.Contains("phó")) return 3;
                                              }
                                              else
                                              {
                                                  if (pos.Contains("tổ trưởng") && !pos.Contains("phó")) return 1;
                                                  if (pos.Contains("phó")) return 2;
                                                  if (pos.Contains("công chức")) return 3;
                                              }
                                              return 99;
                                          })
                                          .ThenBy(p => p.FullName)
                                          .ToList();

                    var generalResults = new List<LeaveGeneralSummaryItem>();
                    var annualResults = new List<LeaveSummaryItem>();
                    var sickResults = new List<SickLeaveItem>();
                    var convalescenceResults = new List<ConvalescenceLeaveItem>();
                    var unpaidResults = new List<UnpaidLeaveItem>();

                    int prevYear = selectedYear - 1;

                    foreach (var p in personnelList)
                    {
                        var leaves = p.LeaveHistories ?? new List<LeaveHistory>();

                        // Filter by date range for the current view
                        var rangeLeaves = leaves.Where(h => h.StartDate <= rangeEnd && (h.EndDate ?? h.StartDate) >= rangeStart).ToList();

                        // 1. Phép năm
                        int totalAnnualTarget = CalculateTotalAnnualLeave(p, selectedYear);
                        double totalCurrentYearTakenAnnual = leaves
                            .Where(h => (h.LeaveType == "Phép năm" || h.LeaveType == "Nghỉ phép") &&
                                        h.StartDate.Year == selectedYear && (h.LeaveYear == selectedYear || h.LeaveYear == null))
                            .Sum(h => h.DurationDays);

                        double remainingAnnual = totalAnnualTarget - totalCurrentYearTakenAnnual;

                        var annualInPeriod = rangeLeaves
                            .Where(h => h.LeaveType == "Phép năm" || h.LeaveType == "Nghỉ phép")
                            .OrderBy(h => h.StartDate)
                            .ToList();

                        double takenFromOldYear = annualInPeriod
                            .Where(h => h.LeaveYear == prevYear)
                            .Sum(h => h.DurationDays);

                        double takenFromCurrentYear = annualInPeriod
                            .Where(h => h.LeaveYear == selectedYear || h.LeaveYear == null)
                            .Sum(h => h.DurationDays);

                        double annualTakenInPeriod = takenFromOldYear + takenFromCurrentYear;

                        var annualDetailLines = new List<LeaveDetailLine>();
                        for (int i = 0; i < annualInPeriod.Count; i++)
                        {
                            var h = annualInPeriod[i];
                            string reason = h.UserReasonDisplay;
                            string note = string.IsNullOrWhiteSpace(reason) ? "" : $": {reason}";

                            annualDetailLines.Add(new LeaveDetailLine
                            {
                                MainContent = $"Lần {i + 1}: {h.DurationDays:0.#} ngày từ ngày {DatePickerHelper.FormatDateForDisplay(h.StartDate)} đến ngày {(h.EndDate.HasValue ? DatePickerHelper.FormatDateForDisplay(h.EndDate.Value) : DatePickerHelper.FormatDateForDisplay(h.StartDate))}{note}",
                                Suffix = $" - [NGHỈ PHÉP NĂM {h.LeaveYear ?? h.StartDate.Year}]"
                            });
                        }
                        string annualDetailedContent = string.Join("\n", annualDetailLines.Select(d => d.MainContent + d.Suffix));

                        // 2. Nghỉ ốm
                        var sickInPeriod = rangeLeaves
                            .Where(h => h.LeaveType == "Nghỉ ốm" || h.LeaveType == "Ốm")
                            .OrderBy(h => h.StartDate)
                            .ToList();

                        double sickDaysInPeriod = sickInPeriod.Sum(h => h.DurationDays);

                        var sickDetailLines = new List<LeaveDetailLine>();
                        for (int i = 0; i < sickInPeriod.Count; i++)
                        {
                            var h = sickInPeriod[i];
                            string reason = h.UserReasonDisplay;
                            string note = string.IsNullOrWhiteSpace(reason) ? "" : $": {reason}";
                            string dateRange = h.EndDate.HasValue && h.EndDate.Value.Date != h.StartDate.Date
                                ? $"từ ngày {DatePickerHelper.FormatDateForDisplay(h.StartDate)} đến ngày {DatePickerHelper.FormatDateForDisplay(h.EndDate.Value)}"
                                : $"ngày {DatePickerHelper.FormatDateForDisplay(h.StartDate)}";

                            sickDetailLines.Add(new LeaveDetailLine
                            {
                                MainContent = $"Lần {i + 1}: {h.DurationDays:0.#} ngày {dateRange}{note}",
                                Suffix = ""
                            });
                        }
                        string sickDetailedContent = string.Join("\n", sickDetailLines.Select(d => d.MainContent));

                        // 3. Dưỡng sức
                        var convalescenceInPeriod = rangeLeaves
                            .Where(h => !string.IsNullOrEmpty(h.LeaveType) &&
                                        (h.LeaveType == "Dưỡng sức" || h.LeaveType == "Nghỉ dưỡng sức" ||
                                         h.LeaveType.Contains("dưỡng sức", StringComparison.OrdinalIgnoreCase) ||
                                         h.LeaveType.Contains("duong suc", StringComparison.OrdinalIgnoreCase)))
                            .OrderBy(h => h.StartDate)
                            .ToList();

                        double convalescenceDaysInPeriod = convalescenceInPeriod.Sum(h => h.DurationDays);

                        var convalescenceDetailLines = new List<LeaveDetailLine>();
                        for (int i = 0; i < convalescenceInPeriod.Count; i++)
                        {
                            var h = convalescenceInPeriod[i];
                            string reason = h.UserReasonDisplay;
                            string note = string.IsNullOrWhiteSpace(reason) ? "" : $": {reason}";
                            string dateRange = h.EndDate.HasValue && h.EndDate.Value.Date != h.StartDate.Date
                                ? $"từ ngày {DatePickerHelper.FormatDateForDisplay(h.StartDate)} đến ngày {DatePickerHelper.FormatDateForDisplay(h.EndDate.Value)}"
                                : $"ngày {DatePickerHelper.FormatDateForDisplay(h.StartDate)}";

                            convalescenceDetailLines.Add(new LeaveDetailLine
                            {
                                MainContent = $"Lần {i + 1}: {h.DurationDays:0.#} ngày {dateRange}{note}",
                                Suffix = ""
                            });
                        }
                        string convalescenceDetailedContent = string.Join("\n", convalescenceDetailLines.Select(d => d.MainContent));

                        // 4. Không lương
                        var unpaidInPeriod = rangeLeaves
                            .Where(h => h.LeaveType == "Không lương" || h.LeaveType == "Nghỉ không lương")
                            .OrderBy(h => h.StartDate)
                            .ToList();

                        double unpaidDaysInPeriod = unpaidInPeriod.Sum(h => h.DurationDays);

                        var unpaidDetailLines = new List<LeaveDetailLine>();
                        for (int i = 0; i < unpaidInPeriod.Count; i++)
                        {
                            var h = unpaidInPeriod[i];
                            string reason = h.UserReasonDisplay;
                            string note = string.IsNullOrWhiteSpace(reason) ? "" : $": {reason}";
                            string dateRange = h.EndDate.HasValue && h.EndDate.Value.Date != h.StartDate.Date
                                ? $"từ ngày {DatePickerHelper.FormatDateForDisplay(h.StartDate)} đến ngày {DatePickerHelper.FormatDateForDisplay(h.EndDate.Value)}"
                                : $"ngày {DatePickerHelper.FormatDateForDisplay(h.StartDate)}";

                            unpaidDetailLines.Add(new LeaveDetailLine
                            {
                                MainContent = $"Lần {i + 1}: {h.DurationDays:0.#} ngày {dateRange}{note}",
                                Suffix = ""
                            });
                        }
                        string unpaidDetailedContent = string.Join("\n", unpaidDetailLines.Select(d => d.MainContent));

                        // Tab 0: General Summary Item
                        generalResults.Add(new LeaveGeneralSummaryItem
                        {
                            PersonnelId = p.Id,
                            StaffId = p.StaffId ?? string.Empty,
                            IdentityCardNumber = p.IdentityCardNumber ?? string.Empty,
                            FullName = p.FullName ?? string.Empty,
                            Department = p.Department ?? string.Empty,
                            AnnualLeaveDays = annualTakenInPeriod,
                            SickLeaveDays = sickDaysInPeriod,
                            ConvalescenceLeaveDays = convalescenceDaysInPeriod,
                            UnpaidLeaveDays = unpaidDaysInPeriod,
                            TotalLeaveDays = annualTakenInPeriod + sickDaysInPeriod + convalescenceDaysInPeriod + unpaidDaysInPeriod,
                            AvatarBase64 = p.AvatarBase64
                        });

                        // Tab 1: Annual Leave Item (Skip if month selected and 0 days taken, matching original logic)
                        if (!selectedMonth.HasValue || annualTakenInPeriod > 0)
                        {
                            annualResults.Add(new LeaveSummaryItem
                            {
                                PersonnelId = p.Id,
                                StaffId = p.StaffId ?? string.Empty,
                                IdentityCardNumber = p.IdentityCardNumber ?? string.Empty,
                                FullName = p.FullName ?? string.Empty,
                                Department = p.Department ?? string.Empty,
                                TotalTarget = totalAnnualTarget,
                                TakenFromOldYear = takenFromOldYear,
                                TakenFromCurrentYear = takenFromCurrentYear,
                                ActualTaken = annualTakenInPeriod,
                                DetailedContent = annualDetailedContent,
                                DetailLines = annualDetailLines,
                                Remaining = remainingAnnual,
                                AvatarBase64 = p.AvatarBase64
                            });
                        }

                        // Tab 2: Sick Leave Item
                        sickResults.Add(new SickLeaveItem
                        {
                            PersonnelId = p.Id,
                            StaffId = p.StaffId ?? string.Empty,
                            IdentityCardNumber = p.IdentityCardNumber ?? string.Empty,
                            FullName = p.FullName ?? string.Empty,
                            Department = p.Department ?? string.Empty,
                            SickDays = sickDaysInPeriod,
                            DetailedContent = sickDetailedContent,
                            DetailLines = sickDetailLines,
                            AvatarBase64 = p.AvatarBase64
                        });

                        // Tab 3: Convalescence Leave Item
                        convalescenceResults.Add(new ConvalescenceLeaveItem
                        {
                            PersonnelId = p.Id,
                            StaffId = p.StaffId ?? string.Empty,
                            IdentityCardNumber = p.IdentityCardNumber ?? string.Empty,
                            FullName = p.FullName ?? string.Empty,
                            Department = p.Department ?? string.Empty,
                            ConvalescenceDays = convalescenceDaysInPeriod,
                            DetailedContent = convalescenceDetailedContent,
                            DetailLines = convalescenceDetailLines,
                            AvatarBase64 = p.AvatarBase64
                        });

                        // Tab 4: Unpaid Leave Item
                        unpaidResults.Add(new UnpaidLeaveItem
                        {
                            PersonnelId = p.Id,
                            StaffId = p.StaffId ?? string.Empty,
                            IdentityCardNumber = p.IdentityCardNumber ?? string.Empty,
                            FullName = p.FullName ?? string.Empty,
                            Department = p.Department ?? string.Empty,
                            UnpaidDays = unpaidDaysInPeriod,
                            DetailedContent = unpaidDetailedContent,
                            DetailLines = unpaidDetailLines,
                            AvatarBase64 = p.AvatarBase64
                        });
                    }

                    _allGeneralSummaries = generalResults;
                    _allAnnualSummaries = annualResults;
                    _allSickSummaries = sickResults;
                    _allConvalescenceSummaries = convalescenceResults;
                    _allUnpaidSummaries = unpaidResults;

                    ApplyFilter();
                }
            }
            catch (Exception ex)
            {
                App.DebugLog("Error loading all leave summaries: " + ex.Message);
            }
            finally
            {
                _isUpdating = false;
            }
        }

        private void ApplyFilter()
        {
            if (!_isInitialized || txtSearch == null) return;

            string rawKeyword = txtSearch.Text?.Trim() ?? string.Empty;
            string[] keywords = string.IsNullOrEmpty(rawKeyword)
                ? Array.Empty<string>()
                : RemoveSign(rawKeyword.ToLower()).Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            bool onlyWithLeaves = chkOnlyWithLeaves?.IsChecked == true;

            switch (_activeTab)
            {
                case 0:
                    // Tab Tổng hợp
                    var genList = _allGeneralSummaries.AsEnumerable();
                    if (onlyWithLeaves)
                    {
                        genList = genList.Where(s => s.TotalLeaveDays > 0);
                    }
                    if (keywords.Length > 0)
                    {
                        genList = genList.Where(s =>
                        {
                            string target = RemoveSign($"{s.FullName} {s.StaffId} {s.IdentityCardNumber} {s.Department}".ToLower());
                            return keywords.All(k => target.Contains(k));
                        });
                    }
                    _filteredGeneralSummaries = genList.ToList();
                    for (int i = 0; i < _filteredGeneralSummaries.Count; i++)
                    {
                        _filteredGeneralSummaries[i].STT = i + 1;
                    }
                    break;

                case 1:
                    // Tab Nghỉ phép
                    var annList = _allAnnualSummaries.AsEnumerable();
                    if (keywords.Length > 0)
                    {
                        annList = annList.Where(s =>
                        {
                            string target = RemoveSign($"{s.FullName} {s.StaffId} {s.IdentityCardNumber} {s.DetailedContent} {s.Department}".ToLower());
                            return keywords.All(k => target.Contains(k));
                        });
                    }
                    _filteredAnnualSummaries = annList.ToList();
                    for (int i = 0; i < _filteredAnnualSummaries.Count; i++)
                    {
                        _filteredAnnualSummaries[i].STT = i + 1;
                    }
                    break;

                case 2:
                    // Tab Nghỉ ốm
                    var sickList = _allSickSummaries.AsEnumerable();
                    if (onlyWithLeaves)
                    {
                        sickList = sickList.Where(s => s.SickDays > 0);
                    }
                    if (keywords.Length > 0)
                    {
                        sickList = sickList.Where(s =>
                        {
                            string target = RemoveSign($"{s.FullName} {s.StaffId} {s.IdentityCardNumber} {s.DetailedContent} {s.Department}".ToLower());
                            return keywords.All(k => target.Contains(k));
                        });
                    }
                    _filteredSickSummaries = sickList.ToList();
                    for (int i = 0; i < _filteredSickSummaries.Count; i++)
                    {
                        _filteredSickSummaries[i].STT = i + 1;
                    }
                    break;

                case 3:
                    // Tab Dưỡng sức
                    var convList = _allConvalescenceSummaries.AsEnumerable();
                    if (onlyWithLeaves)
                    {
                        convList = convList.Where(s => s.ConvalescenceDays > 0);
                    }
                    if (keywords.Length > 0)
                    {
                        convList = convList.Where(s =>
                        {
                            string target = RemoveSign($"{s.FullName} {s.StaffId} {s.IdentityCardNumber} {s.DetailedContent} {s.Department}".ToLower());
                            return keywords.All(k => target.Contains(k));
                        });
                    }
                    _filteredConvalescenceSummaries = convList.ToList();
                    for (int i = 0; i < _filteredConvalescenceSummaries.Count; i++)
                    {
                        _filteredConvalescenceSummaries[i].STT = i + 1;
                    }
                    break;

                case 4:
                    // Tab Không lương
                    var unpaidList = _allUnpaidSummaries.AsEnumerable();
                    if (onlyWithLeaves)
                    {
                        unpaidList = unpaidList.Where(s => s.UnpaidDays > 0);
                    }
                    if (keywords.Length > 0)
                    {
                        unpaidList = unpaidList.Where(s =>
                        {
                            string target = RemoveSign($"{s.FullName} {s.StaffId} {s.IdentityCardNumber} {s.DetailedContent} {s.Department}".ToLower());
                            return keywords.All(k => target.Contains(k));
                        });
                    }
                    _filteredUnpaidSummaries = unpaidList.ToList();
                    for (int i = 0; i < _filteredUnpaidSummaries.Count; i++)
                    {
                        _filteredUnpaidSummaries[i].STT = i + 1;
                    }
                    break;
            }

            _currentPage = 1;
            ApplyPagination();
        }

        private void ApplyPagination()
        {
            if (!_isInitialized) return;

            int totalItems = 0;
            switch (_activeTab)
            {
                case 0: totalItems = _filteredGeneralSummaries.Count; break;
                case 1: totalItems = _filteredAnnualSummaries.Count; break;
                case 2: totalItems = _filteredSickSummaries.Count; break;
                case 3: totalItems = _filteredConvalescenceSummaries.Count; break;
                case 4: totalItems = _filteredUnpaidSummaries.Count; break;
            }

            _totalPages = Math.Max(1, (int)Math.Ceiling((double)totalItems / PageSize));
            if (_currentPage > _totalPages) _currentPage = _totalPages;
            if (_currentPage < 1) _currentPage = 1;

            int startIndex = (_currentPage - 1) * PageSize;

            switch (_activeTab)
            {
                case 0:
                    var genPage = _filteredGeneralSummaries.Skip(startIndex).Take(PageSize).ToList();
                    for (int i = 0; i < genPage.Count; i++) genPage[i].STT = startIndex + i + 1;
                    dgGeneralSummary.ItemsSource = genPage;
                    if (genPage.Count > 0) dgGeneralSummary.ScrollIntoView(genPage[0]);
                    break;

                case 1:
                    var annPage = _filteredAnnualSummaries.Skip(startIndex).Take(PageSize).ToList();
                    for (int i = 0; i < annPage.Count; i++) annPage[i].STT = startIndex + i + 1;
                    dgLeaveSummary.ItemsSource = annPage;
                    if (annPage.Count > 0) dgLeaveSummary.ScrollIntoView(annPage[0]);
                    break;

                case 2:
                    var sickPage = _filteredSickSummaries.Skip(startIndex).Take(PageSize).ToList();
                    for (int i = 0; i < sickPage.Count; i++) sickPage[i].STT = startIndex + i + 1;
                    dgSickLeave.ItemsSource = sickPage;
                    if (sickPage.Count > 0) dgSickLeave.ScrollIntoView(sickPage[0]);
                    break;

                case 3:
                    var convPage = _filteredConvalescenceSummaries.Skip(startIndex).Take(PageSize).ToList();
                    for (int i = 0; i < convPage.Count; i++) convPage[i].STT = startIndex + i + 1;
                    dgConvalescenceLeave.ItemsSource = convPage;
                    if (convPage.Count > 0) dgConvalescenceLeave.ScrollIntoView(convPage[0]);
                    break;

                case 4:
                    var unpaidPage = _filteredUnpaidSummaries.Skip(startIndex).Take(PageSize).ToList();
                    for (int i = 0; i < unpaidPage.Count; i++) unpaidPage[i].STT = startIndex + i + 1;
                    dgUnpaidLeave.ItemsSource = unpaidPage;
                    if (unpaidPage.Count > 0) dgUnpaidLeave.ScrollIntoView(unpaidPage[0]);
                    break;
            }

            // Update footer
            if (totalItems == 0)
            {
                txtPagingInfo.Text = "Không có dữ liệu";
                txtPageInfo.Text = "0 / 0";
            }
            else
            {
                int countOnPage = 0;
                switch (_activeTab)
                {
                    case 0: countOnPage = ((List<LeaveGeneralSummaryItem>)dgGeneralSummary.ItemsSource)?.Count ?? 0; break;
                    case 1: countOnPage = ((List<LeaveSummaryItem>)dgLeaveSummary.ItemsSource)?.Count ?? 0; break;
                    case 2: countOnPage = ((List<SickLeaveItem>)dgSickLeave.ItemsSource)?.Count ?? 0; break;
                    case 3: countOnPage = ((List<ConvalescenceLeaveItem>)dgConvalescenceLeave.ItemsSource)?.Count ?? 0; break;
                    case 4: countOnPage = ((List<UnpaidLeaveItem>)dgUnpaidLeave.ItemsSource)?.Count ?? 0; break;
                }

                int from = startIndex + 1;
                int to = startIndex + countOnPage;
                txtPagingInfo.Text = $"Hiển thị {from} - {to} trên {totalItems} nhân viên";
                txtPageInfo.Text = $"{_currentPage} / {_totalPages}";
            }

            btnPrevPage.IsEnabled = _currentPage > 1;
            btnNextPage.IsEnabled = _currentPage < _totalPages;
        }

        private void BtnPrevPage_Click(object sender, RoutedEventArgs e)
        {
            if (_currentPage > 1)
            {
                _currentPage--;
                ApplyPagination();
            }
        }

        private void BtnNextPage_Click(object sender, RoutedEventArgs e)
        {
            if (_currentPage < _totalPages)
            {
                _currentPage++;
                ApplyPagination();
            }
        }

        private int CalculateTotalAnnualLeave(Personnel p, int selectedYear)
        {
            int baseDays = 12;
            DateTime? targetStart = p.LeaveCalculationDate ?? p.TaxAuthorityStartDate ?? p.StartDate;

            if (targetStart.HasValue)
            {
                DateTime start = targetStart.Value;
                int years = selectedYear - start.Year + 1;
                if (years < 0) years = 0;
                baseDays += (years / 5);
            }
            return baseDays;
        }

        #region Link to Personnel Detail

        private void PersonnelName_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                if (sender is FrameworkElement elem)
                {
                    OpenPersonnelDetailFromItem(elem.DataContext);
                    e.Handled = true;
                }
            }
        }

        private void DataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (sender is DataGrid dg && dg.SelectedItem != null)
            {
                OpenPersonnelDetailFromItem(dg.SelectedItem);
            }
        }

        private void OpenPersonnelDetailFromItem(object? item)
        {
            int personnelId = 0;
            if (item is LeaveGeneralSummaryItem gen) personnelId = gen.PersonnelId;
            else if (item is LeaveSummaryItem ann) personnelId = ann.PersonnelId;
            else if (item is SickLeaveItem sick) personnelId = sick.PersonnelId;
            else if (item is ConvalescenceLeaveItem conv) personnelId = conv.PersonnelId;
            else if (item is UnpaidLeaveItem unp) personnelId = unp.PersonnelId;

            if (personnelId > 0 && Application.Current.MainWindow is MainWindow mw)
            {
                using var context = new AppDbContext();
                var p = context.Personnel
                               .Include(x => x.SalaryRecords)
                               .Include(x => x.LeaveHistories)
                               .Include(x => x.EvaluationRecords)
                               .Include(x => x.PersonnelDegrees)
                               .FirstOrDefault(x => x.Id == personnelId);
                if (p != null)
                {
                    // Chuyển tới Tab 5 (Index 4: THÔNG TIN NGHỈ PHÉP) của cán bộ
                    mw.NavigateToPersonnelDetail(p, 4);
                }
            }
        }

        #endregion

        #region Export Excel

        private void BtnExportExcel_Click(object sender, RoutedEventArgs e)
        {
            switch (_activeTab)
            {
                case 0: ExportGeneralSummaryExcel(); break;
                case 1: ExportAnnualLeaveExcel(); break;
                case 2: ExportSickLeaveExcel(); break;
                case 3: ExportConvalescenceLeaveExcel(); break;
                case 4: ExportUnpaidLeaveExcel(); break;
            }
        }

        private void ShowWarning(string message, string title = "Thông báo")
        {
            var warning = new WarningWindow(message, title);
            if (Window.GetWindow(this) is Window parent) warning.Owner = parent;
            warning.ShowDialog();
        }

        private void ExportGeneralSummaryExcel()
        {
            if (_filteredGeneralSummaries == null || !_filteredGeneralSummaries.Any())
            {
                ShowWarning("Không có dữ liệu để xuất!", "Thông báo");
                return;
            }

            int selectedYear = DateTime.Now.Year;
            if (cbYearFilter?.SelectedItem is int yr) selectedYear = yr;

            int? selectedMonth = null;
            if (cbMonthFilter != null && cbMonthFilter.SelectedIndex > 0) selectedMonth = cbMonthFilter.SelectedIndex;

            string monthFileSuffix = selectedMonth.HasValue ? $"Thang{selectedMonth.Value}_" : "";
            var saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = $"BaoCao_TongHopNgayNghi_{monthFileSuffix}{selectedYear}_{DateTime.Now:yyyyMMdd}.xlsx"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("Tổng hợp ngày nghỉ");

                        string title = selectedMonth.HasValue
                            ? $"BẢNG TỔNG HỢP NGÀY NGHỈ THÁNG {selectedMonth.Value} NĂM {selectedYear}"
                            : $"BẢNG TỔNG HỢP NGÀY NGHỈ NĂM {selectedYear}";

                        worksheet.Cell(1, 1).Value = title;
                        worksheet.Range(1, 1, 1, 9).Merge();
                        worksheet.Cell(1, 1).Style.Font.Bold = true;
                        worksheet.Cell(1, 1).Style.Font.FontSize = 16;
                        worksheet.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml("#1976D2");
                        worksheet.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(1, 1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Row(1).Height = 40;

                        // Headers
                        string[] headers = { "STT", "HỌ VÀ TÊN", "BỘ PHẬN", "SỐ CCCD", "PHÉP (NGÀY)", "ỐM (NGÀY)", "DƯỠNG SỨC (NGÀY)", "KHÔNG LƯƠNG (NGÀY)", "TỔNG SỐ NGÀY" };
                        for (int col = 0; col < headers.Length; col++)
                        {
                            var cell = worksheet.Cell(3, col + 1);
                            cell.Value = headers[col];
                            cell.Style.Font.Bold = true;
                            cell.Style.Font.FontSize = 11;
                            cell.Style.Font.FontColor = XLColor.White;
                            cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#1976D2");
                            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            cell.Style.Border.OutsideBorderColor = XLColor.FromHtml("#CCCCCC");
                        }
                        worksheet.Row(3).Height = 30;

                        int currentRow = 4;
                        for (int i = 0; i < _filteredGeneralSummaries.Count; i++)
                        {
                            var item = _filteredGeneralSummaries[i];
                            worksheet.Cell(currentRow, 1).Value = i + 1;
                            worksheet.Cell(currentRow, 2).Value = item.FullName;
                            worksheet.Cell(currentRow, 3).Value = item.Department;
                            worksheet.Cell(currentRow, 4).Value = item.IdentityCardNumber;
                            worksheet.Cell(currentRow, 5).Value = item.AnnualLeaveDays;
                            worksheet.Cell(currentRow, 6).Value = item.SickLeaveDays;
                            worksheet.Cell(currentRow, 7).Value = item.ConvalescenceLeaveDays;
                            worksheet.Cell(currentRow, 8).Value = item.UnpaidLeaveDays;
                            worksheet.Cell(currentRow, 9).Value = item.TotalLeaveDays;

                            worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                            worksheet.Cell(currentRow, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                            worksheet.Cell(currentRow, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(currentRow, 5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(currentRow, 6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(currentRow, 7).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(currentRow, 8).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(currentRow, 9).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(currentRow, 9).Style.Font.Bold = true;

                            for (int col = 1; col <= 9; col++)
                            {
                                worksheet.Cell(currentRow, col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                worksheet.Cell(currentRow, col).Style.Border.OutsideBorderColor = XLColor.FromHtml("#E0E0E0");
                            }

                            if (currentRow % 2 == 1)
                            {
                                worksheet.Range(currentRow, 1, currentRow, 9).Style.Fill.BackgroundColor = XLColor.FromHtml("#F8FAFC");
                            }

                            worksheet.Row(currentRow).Height = 24;
                            currentRow++;
                        }

                        worksheet.Column(1).Width = 8;
                        worksheet.Column(2).Width = 26;
                        worksheet.Column(3).Width = 25;
                        worksheet.Column(4).Width = 18;
                        worksheet.Column(5).Width = 16;
                        worksheet.Column(6).Width = 16;
                        worksheet.Column(7).Width = 18;
                        worksheet.Column(8).Width = 20;
                        worksheet.Column(9).Width = 18;

                        workbook.SaveAs(saveFileDialog.FileName);
                        var success = new SuccessWindow("Xuất báo cáo Excel thành công!", null, saveFileDialog.FileName, true);
                        if (Window.GetWindow(this) is Window parent) success.Owner = parent;
                        success.ShowDialog();
                    }
                }
                catch (Exception ex)
                {
                    ShowWarning($"Lỗi khi xuất Excel: {ex.Message}", "Lỗi xuất file");
                }
            }
        }

        private void ExportAnnualLeaveExcel()
        {
            if (_filteredAnnualSummaries == null || !_filteredAnnualSummaries.Any())
            {
                ShowWarning("Không có dữ liệu để xuất!", "Thông báo");
                return;
            }

            int selectedYear = DateTime.Now.Year;
            if (cbYearFilter != null && cbYearFilter.SelectedItem is int yr)
            {
                selectedYear = yr;
            }

            int? selectedMonth = null;
            if (cbMonthFilter != null && cbMonthFilter.SelectedIndex > 0)
            {
                selectedMonth = cbMonthFilter.SelectedIndex;
            }

            string monthFileSuffix = selectedMonth.HasValue ? $"Thang{selectedMonth.Value}_" : "";
            var saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = $"BaoCao_PhepNam_{monthFileSuffix}{selectedYear}_{DateTime.Now:yyyyMMdd}.xlsx"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("Chi tiết phép năm");

                        string title = $"BẢNG TỔNG HỢP SỐ NGÀY NGHỈ PHÉP THỰC TẾ NĂM {selectedYear}";
                        if (selectedMonth.HasValue)
                        {
                            title = $"BẢNG TỔNG HỢP SỐ NGÀY NGHỈ PHÉP THỰC TẾ THÁNG {selectedMonth.Value} NĂM {selectedYear}";
                        }

                        worksheet.Cell(1, 1).Value = title;
                        worksheet.Range(1, 1, 1, 9).Merge();
                        worksheet.Cell(1, 1).Style.Font.Bold = true;
                        worksheet.Cell(1, 1).Style.Font.FontSize = 16;
                        worksheet.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml("#1976D2");
                        worksheet.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(1, 1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Row(1).Height = 40;

                        // Row 2 & 3 Headers
                        worksheet.Cell(2, 1).Value = "STT";
                        worksheet.Range(2, 1, 3, 1).Merge();

                        worksheet.Cell(2, 2).Value = "HỌ VÀ TÊN";
                        worksheet.Range(2, 2, 3, 2).Merge();

                        worksheet.Cell(2, 3).Value = "SỐ CCCD";
                        worksheet.Range(2, 3, 3, 3).Merge();

                        worksheet.Cell(2, 4).Value = "PHÉP TIÊU CHUẨN";
                        worksheet.Range(2, 4, 3, 4).Merge();

                        worksheet.Cell(2, 5).Value = "THỰC TẾ ĐÃ NGHỈ";
                        worksheet.Range(2, 5, 2, 7).Merge();

                        worksheet.Cell(3, 5).Value = $"Số ngày phép nghỉ theo chế độ năm {selectedYear - 1}";
                        worksheet.Cell(3, 6).Value = $"Số ngày phép nghỉ theo chế độ năm {selectedYear}";
                        worksheet.Cell(3, 7).Value = "Tổng";

                        worksheet.Cell(2, 8).Value = "SỐ PHÉP CÒN LẠI";
                        worksheet.Range(2, 8, 3, 8).Merge();

                        worksheet.Cell(2, 9).Value = "NỘI DUNG CHI TIẾT";
                        worksheet.Range(2, 9, 3, 9).Merge();

                        var headerRange = worksheet.Range(2, 1, 3, 9);
                        headerRange.Style.Font.Bold = true;
                        headerRange.Style.Font.FontSize = 11;
                        headerRange.Style.Font.FontColor = XLColor.White;
                        headerRange.Style.Fill.BackgroundColor = XLColor.FromHtml("#1976D2");
                        headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        headerRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        headerRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        headerRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                        headerRange.Style.Border.OutsideBorderColor = XLColor.FromHtml("#CCCCCC");
                        headerRange.Style.Border.InsideBorderColor = XLColor.FromHtml("#CCCCCC");

                        int currentRow = 4;
                        for (int i = 0; i < _filteredAnnualSummaries.Count; i++)
                        {
                            var item = _filteredAnnualSummaries[i];

                            worksheet.Cell(currentRow, 1).Value = i + 1;
                            worksheet.Cell(currentRow, 2).Value = item.FullName;
                            worksheet.Cell(currentRow, 3).Value = item.IdentityCardNumber;
                            worksheet.Cell(currentRow, 4).Value = item.TotalTarget;
                            worksheet.Cell(currentRow, 5).Value = item.TakenFromOldYear;
                            worksheet.Cell(currentRow, 6).Value = item.TakenFromCurrentYear;
                            worksheet.Cell(currentRow, 7).Value = item.ActualTaken;
                            worksheet.Cell(currentRow, 8).Value = item.Remaining;

                            string detailExcel = string.Join("\n", item.DetailLines.Select(d =>
                                string.IsNullOrEmpty(d.Suffix) ? d.MainContent : $"{d.MainContent}{d.Suffix}"));
                            worksheet.Cell(currentRow, 9).Value = detailExcel;

                            worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                            worksheet.Cell(currentRow, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(currentRow, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(currentRow, 5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(currentRow, 6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(currentRow, 7).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(currentRow, 8).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(currentRow, 9).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                            worksheet.Cell(currentRow, 9).Style.Alignment.WrapText = true;

                            for (int col = 1; col <= 9; col++)
                            {
                                worksheet.Cell(currentRow, col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                worksheet.Cell(currentRow, col).Style.Border.OutsideBorderColor = XLColor.FromHtml("#E0E0E0");
                            }

                            if (currentRow % 2 == 1)
                            {
                                worksheet.Range(currentRow, 1, currentRow, 9).Style.Fill.BackgroundColor = XLColor.FromHtml("#F9FBFD");
                            }

                            int linesCount = Math.Max(1, item.DetailLines.Count);
                            worksheet.Row(currentRow).Height = Math.Max(26, linesCount * 18);
                            currentRow++;
                        }

                        worksheet.Column(1).Width = 8;
                        worksheet.Column(2).Width = 26;
                        worksheet.Column(3).Width = 18;
                        worksheet.Column(4).Width = 16;
                        worksheet.Column(5).Width = 25;
                        worksheet.Column(6).Width = 25;
                        worksheet.Column(7).Width = 12;
                        worksheet.Column(8).Width = 15;
                        worksheet.Column(9).Width = 85;

                        worksheet.Row(2).Height = 35;
                        worksheet.Row(2).Style.Alignment.WrapText = true;
                        worksheet.Row(3).Height = 55;
                        worksheet.Row(3).Style.Alignment.WrapText = true;

                        workbook.SaveAs(saveFileDialog.FileName);
                        var success = new SuccessWindow("Xuất báo cáo Excel thành công!", null, saveFileDialog.FileName, true);
                        if (Window.GetWindow(this) is Window parent) success.Owner = parent;
                        success.ShowDialog();
                    }
                }
                catch (Exception ex)
                {
                    ShowWarning($"Lỗi khi xuất Excel: {ex.Message}", "Lỗi xuất file");
                }
            }
        }

        private void ExportSickLeaveExcel()
        {
            if (_filteredSickSummaries == null || !_filteredSickSummaries.Any())
            {
                ShowWarning("Không có dữ liệu để xuất!", "Thông báo");
                return;
            }

            int selectedYear = DateTime.Now.Year;
            if (cbYearFilter?.SelectedItem is int yr) selectedYear = yr;

            int? selectedMonth = null;
            if (cbMonthFilter != null && cbMonthFilter.SelectedIndex > 0) selectedMonth = cbMonthFilter.SelectedIndex;

            string monthFileSuffix = selectedMonth.HasValue ? $"Thang{selectedMonth.Value}_" : "";
            var saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = $"BaoCao_NghiOm_{monthFileSuffix}{selectedYear}_{DateTime.Now:yyyyMMdd}.xlsx"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("Chi tiết nghỉ ốm");

                        string title = selectedMonth.HasValue
                            ? $"BẢNG CHI TIẾT NGHỈ ỐM THÁNG {selectedMonth.Value} NĂM {selectedYear}"
                            : $"BẢNG CHI TIẾT NGHỈ ỐM NĂM {selectedYear}";

                        worksheet.Cell(1, 1).Value = title;
                        worksheet.Range(1, 1, 1, 6).Merge();
                        worksheet.Cell(1, 1).Style.Font.Bold = true;
                        worksheet.Cell(1, 1).Style.Font.FontSize = 16;
                        worksheet.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml("#1976D2");
                        worksheet.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(1, 1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Row(1).Height = 40;

                        string[] headers = { "STT", "HỌ VÀ TÊN", "BỘ PHẬN", "SỐ CCCD", "SỐ NGÀY NGHỈ ỐM", "NỘI DUNG CHI TIẾT" };
                        for (int col = 0; col < headers.Length; col++)
                        {
                            var cell = worksheet.Cell(3, col + 1);
                            cell.Value = headers[col];
                            cell.Style.Font.Bold = true;
                            cell.Style.Font.FontSize = 11;
                            cell.Style.Font.FontColor = XLColor.White;
                            cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#1976D2");
                            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            cell.Style.Border.OutsideBorderColor = XLColor.FromHtml("#CCCCCC");
                        }
                        worksheet.Row(3).Height = 30;

                        int currentRow = 4;
                        for (int i = 0; i < _filteredSickSummaries.Count; i++)
                        {
                            var item = _filteredSickSummaries[i];
                            worksheet.Cell(currentRow, 1).Value = i + 1;
                            worksheet.Cell(currentRow, 2).Value = item.FullName;
                            worksheet.Cell(currentRow, 3).Value = item.Department;
                            worksheet.Cell(currentRow, 4).Value = item.IdentityCardNumber;
                            worksheet.Cell(currentRow, 5).Value = item.SickDays;

                            string detailExcel = string.Join("\n", item.DetailLines.Select(d => d.MainContent));
                            worksheet.Cell(currentRow, 6).Value = detailExcel;

                            worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                            worksheet.Cell(currentRow, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                            worksheet.Cell(currentRow, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(currentRow, 5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(currentRow, 5).Style.Font.Bold = true;
                            worksheet.Cell(currentRow, 6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                            worksheet.Cell(currentRow, 6).Style.Alignment.WrapText = true;

                            for (int col = 1; col <= 6; col++)
                            {
                                worksheet.Cell(currentRow, col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                worksheet.Cell(currentRow, col).Style.Border.OutsideBorderColor = XLColor.FromHtml("#E0E0E0");
                            }

                            if (currentRow % 2 == 1)
                            {
                                worksheet.Range(currentRow, 1, currentRow, 6).Style.Fill.BackgroundColor = XLColor.FromHtml("#F8FAFC");
                            }

                            int linesCount = Math.Max(1, item.DetailLines.Count);
                            worksheet.Row(currentRow).Height = Math.Max(26, linesCount * 18);
                            currentRow++;
                        }

                        worksheet.Column(1).Width = 8;
                        worksheet.Column(2).Width = 26;
                        worksheet.Column(3).Width = 25;
                        worksheet.Column(4).Width = 18;
                        worksheet.Column(5).Width = 20;
                        worksheet.Column(6).Width = 80;

                        workbook.SaveAs(saveFileDialog.FileName);
                        var success = new SuccessWindow("Xuất báo cáo Excel thành công!", null, saveFileDialog.FileName, true);
                        if (Window.GetWindow(this) is Window parent) success.Owner = parent;
                        success.ShowDialog();
                    }
                }
                catch (Exception ex)
                {
                    ShowWarning($"Lỗi khi xuất Excel: {ex.Message}", "Lỗi xuất file");
                }
            }
        }

        private void ExportConvalescenceLeaveExcel()
        {
            if (_filteredConvalescenceSummaries == null || !_filteredConvalescenceSummaries.Any())
            {
                ShowWarning("Không có dữ liệu để xuất!", "Thông báo");
                return;
            }

            int selectedYear = DateTime.Now.Year;
            if (cbYearFilter?.SelectedItem is int yr) selectedYear = yr;

            int? selectedMonth = null;
            if (cbMonthFilter != null && cbMonthFilter.SelectedIndex > 0) selectedMonth = cbMonthFilter.SelectedIndex;

            string monthFileSuffix = selectedMonth.HasValue ? $"Thang{selectedMonth.Value}_" : "";
            var saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = $"BaoCao_DuongSuc_{monthFileSuffix}{selectedYear}_{DateTime.Now:yyyyMMdd}.xlsx"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("Chi tiết dưỡng sức");

                        string title = selectedMonth.HasValue
                            ? $"BẢNG CHI TIẾT NGHỈ DƯỠNG SỨC THÁNG {selectedMonth.Value} NĂM {selectedYear}"
                            : $"BẢNG CHI TIẾT NGHỈ DƯỠNG SỨC NĂM {selectedYear}";

                        worksheet.Cell(1, 1).Value = title;
                        worksheet.Range(1, 1, 1, 6).Merge();
                        worksheet.Cell(1, 1).Style.Font.Bold = true;
                        worksheet.Cell(1, 1).Style.Font.FontSize = 16;
                        worksheet.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml("#1976D2");
                        worksheet.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(1, 1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Row(1).Height = 40;

                        string[] headers = { "STT", "HỌ VÀ TÊN", "BỘ PHẬN", "SỐ CCCD", "SỐ NGÀY DƯỠNG SỨC", "NỘI DUNG CHI TIẾT" };
                        for (int col = 0; col < headers.Length; col++)
                        {
                            var cell = worksheet.Cell(3, col + 1);
                            cell.Value = headers[col];
                            cell.Style.Font.Bold = true;
                            cell.Style.Font.FontSize = 11;
                            cell.Style.Font.FontColor = XLColor.White;
                            cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#1976D2");
                            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            cell.Style.Border.OutsideBorderColor = XLColor.FromHtml("#CCCCCC");
                        }
                        worksheet.Row(3).Height = 30;

                        int currentRow = 4;
                        for (int i = 0; i < _filteredConvalescenceSummaries.Count; i++)
                        {
                            var item = _filteredConvalescenceSummaries[i];
                            worksheet.Cell(currentRow, 1).Value = i + 1;
                            worksheet.Cell(currentRow, 2).Value = item.FullName;
                            worksheet.Cell(currentRow, 3).Value = item.Department;
                            worksheet.Cell(currentRow, 4).Value = item.IdentityCardNumber;
                            worksheet.Cell(currentRow, 5).Value = item.ConvalescenceDays;

                            string detailExcel = string.Join("\n", item.DetailLines.Select(d => d.MainContent));
                            worksheet.Cell(currentRow, 6).Value = detailExcel;

                            worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                            worksheet.Cell(currentRow, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                            worksheet.Cell(currentRow, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(currentRow, 5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(currentRow, 5).Style.Font.Bold = true;
                            worksheet.Cell(currentRow, 6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                            worksheet.Cell(currentRow, 6).Style.Alignment.WrapText = true;

                            for (int col = 1; col <= 6; col++)
                            {
                                worksheet.Cell(currentRow, col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                worksheet.Cell(currentRow, col).Style.Border.OutsideBorderColor = XLColor.FromHtml("#E0E0E0");
                            }

                            if (currentRow % 2 == 1)
                            {
                                worksheet.Range(currentRow, 1, currentRow, 6).Style.Fill.BackgroundColor = XLColor.FromHtml("#F8FAFC");
                            }

                            int linesCount = Math.Max(1, item.DetailLines.Count);
                            worksheet.Row(currentRow).Height = Math.Max(26, linesCount * 18);
                            currentRow++;
                        }

                        worksheet.Column(1).Width = 8;
                        worksheet.Column(2).Width = 26;
                        worksheet.Column(3).Width = 25;
                        worksheet.Column(4).Width = 18;
                        worksheet.Column(5).Width = 22;
                        worksheet.Column(6).Width = 80;

                        workbook.SaveAs(saveFileDialog.FileName);
                        var success = new SuccessWindow("Xuất báo cáo Excel thành công!", null, saveFileDialog.FileName, true);
                        if (Window.GetWindow(this) is Window parent) success.Owner = parent;
                        success.ShowDialog();
                    }
                }
                catch (Exception ex)
                {
                    ShowWarning($"Lỗi khi xuất Excel: {ex.Message}", "Lỗi xuất file");
                }
            }
        }

        private void ExportUnpaidLeaveExcel()
        {
            if (_filteredUnpaidSummaries == null || !_filteredUnpaidSummaries.Any())
            {
                ShowWarning("Không có dữ liệu để xuất!", "Thông báo");
                return;
            }

            int selectedYear = DateTime.Now.Year;
            if (cbYearFilter?.SelectedItem is int yr) selectedYear = yr;

            int? selectedMonth = null;
            if (cbMonthFilter != null && cbMonthFilter.SelectedIndex > 0) selectedMonth = cbMonthFilter.SelectedIndex;

            string monthFileSuffix = selectedMonth.HasValue ? $"Thang{selectedMonth.Value}_" : "";
            var saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = $"BaoCao_KhongLuong_{monthFileSuffix}{selectedYear}_{DateTime.Now:yyyyMMdd}.xlsx"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("Chi tiết không lương");

                        string title = selectedMonth.HasValue
                            ? $"BẢNG CHI TIẾT NGHỈ KHÔNG LƯƠNG THÁNG {selectedMonth.Value} NĂM {selectedYear}"
                            : $"BẢNG CHI TIẾT NGHỈ KHÔNG LƯƠNG NĂM {selectedYear}";

                        worksheet.Cell(1, 1).Value = title;
                        worksheet.Range(1, 1, 1, 6).Merge();
                        worksheet.Cell(1, 1).Style.Font.Bold = true;
                        worksheet.Cell(1, 1).Style.Font.FontSize = 16;
                        worksheet.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml("#1976D2");
                        worksheet.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(1, 1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Row(1).Height = 40;

                        string[] headers = { "STT", "HỌ VÀ TÊN", "BỘ PHẬN", "SỐ CCCD", "SỐ NGÀY KHÔNG LƯƠNG", "NỘI DUNG CHI TIẾT" };
                        for (int col = 0; col < headers.Length; col++)
                        {
                            var cell = worksheet.Cell(3, col + 1);
                            cell.Value = headers[col];
                            cell.Style.Font.Bold = true;
                            cell.Style.Font.FontSize = 11;
                            cell.Style.Font.FontColor = XLColor.White;
                            cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#1976D2");
                            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            cell.Style.Border.OutsideBorderColor = XLColor.FromHtml("#CCCCCC");
                        }
                        worksheet.Row(3).Height = 30;

                        int currentRow = 4;
                        for (int i = 0; i < _filteredUnpaidSummaries.Count; i++)
                        {
                            var item = _filteredUnpaidSummaries[i];
                            worksheet.Cell(currentRow, 1).Value = i + 1;
                            worksheet.Cell(currentRow, 2).Value = item.FullName;
                            worksheet.Cell(currentRow, 3).Value = item.Department;
                            worksheet.Cell(currentRow, 4).Value = item.IdentityCardNumber;
                            worksheet.Cell(currentRow, 5).Value = item.UnpaidDays;

                            string detailExcel = string.Join("\n", item.DetailLines.Select(d => d.MainContent));
                            worksheet.Cell(currentRow, 6).Value = detailExcel;

                            worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                            worksheet.Cell(currentRow, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                            worksheet.Cell(currentRow, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(currentRow, 5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(currentRow, 5).Style.Font.Bold = true;
                            worksheet.Cell(currentRow, 6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                            worksheet.Cell(currentRow, 6).Style.Alignment.WrapText = true;

                            for (int col = 1; col <= 6; col++)
                            {
                                worksheet.Cell(currentRow, col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                worksheet.Cell(currentRow, col).Style.Border.OutsideBorderColor = XLColor.FromHtml("#E0E0E0");
                            }

                            if (currentRow % 2 == 1)
                            {
                                worksheet.Range(currentRow, 1, currentRow, 6).Style.Fill.BackgroundColor = XLColor.FromHtml("#F8FAFC");
                            }

                            int linesCount = Math.Max(1, item.DetailLines.Count);
                            worksheet.Row(currentRow).Height = Math.Max(26, linesCount * 18);
                            currentRow++;
                        }

                        worksheet.Column(1).Width = 8;
                        worksheet.Column(2).Width = 26;
                        worksheet.Column(3).Width = 25;
                        worksheet.Column(4).Width = 18;
                        worksheet.Column(5).Width = 24;
                        worksheet.Column(6).Width = 80;

                        workbook.SaveAs(saveFileDialog.FileName);
                        var success = new SuccessWindow("Xuất báo cáo Excel thành công!", null, saveFileDialog.FileName, true);
                        if (Window.GetWindow(this) is Window parent) success.Owner = parent;
                        success.ShowDialog();
                    }
                }
                catch (Exception ex)
                {
                    ShowWarning($"Lỗi khi xuất Excel: {ex.Message}", "Lỗi xuất file");
                }
            }
        }

        #endregion

        #region Helpers & Sizing

        private string RemoveSign(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;

            string normalizedString = text.Normalize(System.Text.NormalizationForm.FormD);
            System.Text.StringBuilder stringBuilder = new System.Text.StringBuilder();

            foreach (char c in normalizedString)
            {
                System.Globalization.UnicodeCategory unicodeCategory = System.Globalization.CharUnicodeInfo.GetUnicodeCategory(c);
                if (unicodeCategory != System.Globalization.UnicodeCategory.NonSpacingMark)
                {
                    stringBuilder.Append(c);
                }
            }

            string result = stringBuilder.ToString().Normalize(System.Text.NormalizationForm.FormC);
            result = result.Replace('đ', 'd').Replace('Đ', 'D');

            return result;
        }

        private void dgLeaveSummary_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (dgLeaveSummary == null || dgLeaveSummary.Columns.Count < 2) return;

            const double fixedColumnsTotal = 70 + 180 + 150 + 150 + 150 + 130 + 400 + 170; // = 1400
            const double nameMinWidth = 300;
            const double scrollBarBuffer = 20;

            double available = dgLeaveSummary.ActualWidth - scrollBarBuffer;
            double nameWidth = Math.Max(nameMinWidth, available - fixedColumnsTotal);

            dgLeaveSummary.Columns[1].Width = new DataGridLength(nameWidth);
        }

        private void dgGeneralSummary_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (dgGeneralSummary == null || dgGeneralSummary.Columns.Count < 2) return;

            const double fixedColumnsTotal = 70 + 220 + 130 + 130 + 130 + 140 + 120; // = 940 (with Dưỡng sức)
            const double nameMinWidth = 300;
            const double scrollBarBuffer = 20;

            double available = dgGeneralSummary.ActualWidth - scrollBarBuffer;
            double nameWidth = Math.Max(nameMinWidth, available - fixedColumnsTotal);

            dgGeneralSummary.Columns[1].Width = new DataGridLength(nameWidth);
        }

        #endregion
    }

    public class LeaveDetailLine
    {
        public string MainContent { get; set; } = string.Empty;
        public string Suffix { get; set; } = string.Empty; // "[NGHỈ PHÉP NĂM CŨ]"
        public bool IsOldYear => !string.IsNullOrEmpty(Suffix);
    }

    public class LeaveGeneralSummaryItem
    {
        public int STT { get; set; }
        public int PersonnelId { get; set; }
        public string StaffId { get; set; } = string.Empty;
        public string IdentityCardNumber { get; set; } = string.Empty;
        public string FullName { get; set; } = string.Empty;
        public string Department { get; set; } = string.Empty;
        public double AnnualLeaveDays { get; set; }
        public double SickLeaveDays { get; set; }
        public double ConvalescenceLeaveDays { get; set; }
        public double UnpaidLeaveDays { get; set; }
        public double TotalLeaveDays { get; set; }
        public string? AvatarBase64 { get; set; }
    }

    public class LeaveSummaryItem
    {
        public int STT { get; set; }
        public int PersonnelId { get; set; }
        public string StaffId { get; set; } = string.Empty;
        public string IdentityCardNumber { get; set; } = string.Empty;
        public string FullName { get; set; } = string.Empty;
        public string Department { get; set; } = string.Empty;
        public int TotalTarget { get; set; }
        public double TakenFromOldYear { get; set; }
        public double TakenFromCurrentYear { get; set; }
        public double ActualTaken { get; set; }
        public string DetailedContent { get; set; } = string.Empty;
        public List<LeaveDetailLine> DetailLines { get; set; } = new List<LeaveDetailLine>();
        public double Remaining { get; set; }
        public string? AvatarBase64 { get; set; }
    }

    public class SickLeaveItem
    {
        public int STT { get; set; }
        public int PersonnelId { get; set; }
        public string StaffId { get; set; } = string.Empty;
        public string IdentityCardNumber { get; set; } = string.Empty;
        public string FullName { get; set; } = string.Empty;
        public string Department { get; set; } = string.Empty;
        public double SickDays { get; set; }
        public string DetailedContent { get; set; } = string.Empty;
        public List<LeaveDetailLine> DetailLines { get; set; } = new List<LeaveDetailLine>();
        public string? AvatarBase64 { get; set; }
    }

    public class ConvalescenceLeaveItem
    {
        public int STT { get; set; }
        public int PersonnelId { get; set; }
        public string StaffId { get; set; } = string.Empty;
        public string IdentityCardNumber { get; set; } = string.Empty;
        public string FullName { get; set; } = string.Empty;
        public string Department { get; set; } = string.Empty;
        public double ConvalescenceDays { get; set; }
        public string DetailedContent { get; set; } = string.Empty;
        public List<LeaveDetailLine> DetailLines { get; set; } = new List<LeaveDetailLine>();
        public string? AvatarBase64 { get; set; }
    }

    public class UnpaidLeaveItem
    {
        public int STT { get; set; }
        public int PersonnelId { get; set; }
        public string StaffId { get; set; } = string.Empty;
        public string IdentityCardNumber { get; set; } = string.Empty;
        public string FullName { get; set; } = string.Empty;
        public string Department { get; set; } = string.Empty;
        public double UnpaidDays { get; set; }
        public string DetailedContent { get; set; } = string.Empty;
        public List<LeaveDetailLine> DetailLines { get; set; } = new List<LeaveDetailLine>();
        public string? AvatarBase64 { get; set; }
    }
}
