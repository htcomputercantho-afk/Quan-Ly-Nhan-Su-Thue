using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using TaxPersonnelManagement.Data;
using TaxPersonnelManagement.Models;

namespace TaxPersonnelManagement.Views
{
    public partial class WorkHistoryDialog : Window
    {
        public WorkHistoryRecord WorkHistory { get; private set; }

        public WorkHistoryDialog(WorkHistoryRecord? recordToEdit = null)
        {
            InitializeComponent();

            InitializeYearPickers();
            LoadSuggestions();

            if (recordToEdit != null)
            {
                txtTitle.Text = "CHỈNH SỬA GIAI ĐOẠN CÔNG TÁC";
                btnSave.Content = "Cập nhật";

                WorkHistory = recordToEdit;
                _isFormattingDate = true;
                txtFromDate.Text = recordToEdit.FromDate ?? "";
                _isFormattingDate = false;

                // Phân tách Bộ phận và Cơ quan/Đơn vị công tác
                string dept = (recordToEdit.Department ?? "").Trim();
                string org = (recordToEdit.OrganizationUnit ?? "").Trim();

                if (string.IsNullOrEmpty(dept) && string.IsNullOrEmpty(org) && !string.IsNullOrWhiteSpace(recordToEdit.DepartmentOrUnit))
                {
                    string full = recordToEdit.DepartmentOrUnit.Trim();
                    if (full.Contains(" - "))
                    {
                        int dash = full.IndexOf(" - ");
                        dept = full.Substring(0, dash).Trim();
                        org = full.Substring(dash + 3).Trim();
                    }
                    else if (full.Contains(","))
                    {
                        int comma = full.IndexOf(',');
                        dept = full.Substring(0, comma).Trim();
                        org = full.Substring(comma + 1).Trim();
                    }
                    else
                    {
                        dept = full;
                    }
                }

                cboDepartment.Text = dept;
                cboOrganizationUnit.Text = org;

                // Chuẩn hóa Chức vụ chuyên môn và Chức vụ Đảng (nếu dữ liệu cũ dính Chức vụ Đảng ở trước hoặc trong ngoặc)
                string rawPos = (recordToEdit.Position ?? "").Trim();
                string partyPos = (recordToEdit.PartyPosition ?? "").Trim();
                bool isParty = recordToEdit.IsPartyConcurrent;

                // Tách nếu có trong ngoặc đơn (Kiêm: ...) hoặc (Chức vụ Đảng)
                var kiemMatch = System.Text.RegularExpressions.Regex.Match(rawPos,
                    @"\s*\((?:Kiêm|Kiêm nhiệm|Kiêm chức vụ Đảng|Kiêm nhiệm chức vụ Đảng):\s*(?<party>[^)]+)\)\s*",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                if (kiemMatch.Success)
                {
                    isParty = true;
                    if (string.IsNullOrWhiteSpace(partyPos))
                        partyPos = kiemMatch.Groups["party"].Value.Trim();
                    rawPos = (rawPos.Substring(0, kiemMatch.Index) + " " + rawPos.Substring(kiemMatch.Index + kiemMatch.Length)).Trim();
                }
                else
                {
                    var parenMatch = System.Text.RegularExpressions.Regex.Match(rawPos, @"\s*\((?<party>[^)]+)\)\s*");
                    if (parenMatch.Success && WorkHistoryRecord.IsPartyRoleText(parenMatch.Groups["party"].Value.Trim()))
                    {
                        isParty = true;
                        if (string.IsNullOrWhiteSpace(partyPos))
                            partyPos = parenMatch.Groups["party"].Value.Trim();
                        rawPos = (rawPos.Substring(0, parenMatch.Index) + " " + rawPos.Substring(parenMatch.Index + parenMatch.Length)).Trim();
                    }
                }

                if (rawPos.Contains(";"))
                {
                    var parts = rawPos.Split(';').Select(p => p.Trim()).Where(p => !string.IsNullOrEmpty(p)).ToList();
                    if (parts.Count >= 2 && WorkHistoryRecord.IsPartyRoleText(parts[0]) && !WorkHistoryRecord.IsPartyRoleText(parts[1]))
                    {
                        partyPos = parts[0];
                        rawPos = string.Join("; ", parts.Skip(1));
                        isParty = true;
                    }
                    else if (parts.Count >= 2 && WorkHistoryRecord.IsPartyRoleText(parts[1]) && !WorkHistoryRecord.IsPartyRoleText(parts[0]))
                    {
                        partyPos = parts[1];
                        rawPos = parts[0];
                        isParty = true;
                    }
                }
                else if (rawPos.Contains(","))
                {
                    int cIdx = rawPos.IndexOf(',');
                    string prefix = rawPos.Substring(0, cIdx).Trim();
                    string suffix = rawPos.Substring(cIdx + 1).Trim();
                    if (WorkHistoryRecord.IsPartyRoleText(prefix))
                    {
                        partyPos = prefix;
                        rawPos = suffix;
                        isParty = true;
                    }
                }

                // Loại bỏ chữ "Kiêm: " nếu có trong partyPos
                if (!string.IsNullOrWhiteSpace(partyPos))
                {
                    var m = System.Text.RegularExpressions.Regex.Match(partyPos, @"^(?:Kiêm|Kiêm nhiệm|Kiêm chức vụ Đảng|Kiêm nhiệm chức vụ Đảng):\s*(.*)$", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                    if (m.Success) partyPos = m.Groups[1].Value.Trim();
                }

                cboPosition.Text = rawPos;

                // Kiêm nhiệm chức vụ Đảng
                chkPartyConcurrent.IsChecked = isParty;
                pnlPartyPosition.Visibility = isParty ? Visibility.Visible : Visibility.Collapsed;
                cboPartyPosition.Text = partyPos;

                txtNote.Text = recordToEdit.Note ?? "";

                string to = (recordToEdit.ToDate ?? "").Trim();
                if (to.Equals("Hiện nay", StringComparison.OrdinalIgnoreCase) ||
                    to.Equals("Đến nay", StringComparison.OrdinalIgnoreCase) ||
                    to.Equals("nay", StringComparison.OrdinalIgnoreCase))
                {
                    chkPresent.IsChecked = true;
                    UpdatePresentState(true);
                }
                else
                {
                    chkPresent.IsChecked = false;
                    _isFormattingDate = true;
                    txtToDate.Text = to;
                    _isFormattingDate = false;
                    UpdatePresentState(false);
                }
            }
            else
            {
                txtTitle.Text = "THÊM GIAI ĐOẠN CÔNG TÁC";
                btnSave.Content = "Lưu giai đoạn";
                WorkHistory = new WorkHistoryRecord();
                chkPartyConcurrent.IsChecked = false;
                pnlPartyPosition.Visibility = Visibility.Collapsed;
                UpdatePresentState(false);
            }
        }

        private ObservableCollection<string> _departmentSuggestions = new ObservableCollection<string>();
        private ObservableCollection<string> _orgSuggestions = new ObservableCollection<string>();
        private ObservableCollection<string> _positionSuggestions = new ObservableCollection<string>();
        private ObservableCollection<string> _partyPositionSuggestions = new ObservableCollection<string>();

        private void LoadSuggestions()
        {
            try
            {
                using (var context = new AppDbContext())
                {
                    // 1. NẠP GỢI Ý BỘ PHẬN (ĐỘC LẬP)
                    var dbDeptSuggestions = context.WorkHistorySuggestions
                        .Where(s => s.Category == "Department" && !string.IsNullOrEmpty(s.Name))
                        .Select(s => s.Name.Trim())
                        .Distinct()
                        .ToList();

                    if (!dbDeptSuggestions.Any())
                    {
                        var seedDepts = new List<string>
                        {
                            "Ban Lãnh đạo",
                            "Tổ Hành chính, tổng hợp",
                            "Tổ Kê khai và Kế toán thuế",
                            "Tổ Quản lý nợ và Cưỡng chế nợ thuế",
                            "Tổ Kiểm tra thuế",
                            "Tổ Nghiệp vụ dự toán pháp chế",
                            "Đội Hành chính - Nhân sự - Tài vụ - Quản trị - Ấn chỉ",
                            "Đội Kiểm tra thuế",
                            "Đội Tuyên truyền - Hỗ trợ người nộp thuế",
                            "Đội Kê khai - Kế toán thuế và Tin học",
                            "Đội Quản lý nợ và Cưỡng chế nợ thuế",
                            "Đội Quản lý thuế liên xã, phường",
                            "Phòng Quản lý nợ và Cưỡng chế nợ thuế",
                            "Phòng Tuyên truyền - Hỗ trợ NNT",
                            "Phòng Kê khai và Kế toán thuế",
                            "Phòng Kiểm tra nội bộ",
                            "Phòng Thanh tra - Kiểm tra thuế"
                        };

                        foreach (var d in seedDepts)
                        {
                            context.WorkHistorySuggestions.Add(new WorkHistorySuggestion { Category = "Department", Name = d });
                        }
                        context.SaveChanges();
                        dbDeptSuggestions = seedDepts;
                    }

                    _departmentSuggestions = new ObservableCollection<string>(
                        dbDeptSuggestions
                            .Where(x => !string.IsNullOrWhiteSpace(x))
                            .Distinct(StringComparer.OrdinalIgnoreCase)
                            .OrderBy(x => x)
                    );
                    cboDepartment.ItemsSource = _departmentSuggestions;

                    // 2. NẠP GỢI Ý CƠ QUAN / ĐƠN VỊ CÔNG TÁC (ĐỘC LẬP)
                    var dbOrgSuggestions = context.WorkHistorySuggestions
                        .Where(s => s.Category == "OrganizationUnit" && !string.IsNullOrEmpty(s.Name))
                        .Select(s => s.Name.Trim())
                        .Distinct()
                        .ToList();

                    if (!dbOrgSuggestions.Any())
                    {
                        var seedOrgs = new List<string>
                        {
                            "Thuế cơ sở 1 thành phố Quảng Ninh",
                            "Thuế cơ sở 2 thành phố Cẩm Phả",
                            "Chi cục Thuế TP Hạ Long",
                            "Chi cục Thuế huyện Hoành Bồ",
                            "Chi cục Thuế TP Cẩm Phả",
                            "Chi cục Thuế thị xã Quảng Yên",
                            "Chi cục Thuế khu vực Uông Bí - Quảng Yên",
                            "Chi cục Thuế khu vực Tiên Yên - Ba Chẽ - Bình Liêu",
                            "Cục Thuế tỉnh Quảng Ninh",
                            "Tổng cục Thuế",
                            "UBND thành phố Hạ Long",
                            "UBND tỉnh Quảng Ninh"
                        };

                        foreach (var o in seedOrgs)
                        {
                            context.WorkHistorySuggestions.Add(new WorkHistorySuggestion { Category = "OrganizationUnit", Name = o });
                        }
                        context.SaveChanges();
                        dbOrgSuggestions = seedOrgs;
                    }

                    _orgSuggestions = new ObservableCollection<string>(
                        dbOrgSuggestions
                            .Where(x => !string.IsNullOrWhiteSpace(x))
                            .Distinct(StringComparer.OrdinalIgnoreCase)
                            .OrderBy(x => x)
                    );
                    cboOrganizationUnit.ItemsSource = _orgSuggestions;

                    // 3. NẠP GỢI Ý CHỨC VỤ (ĐỘC LẬP, KHÔNG BỊ ÉP THEO BỘ PHẬN)
                    var dbPosSuggestions = context.WorkHistorySuggestions
                        .Where(s => s.Category == "Position" && !string.IsNullOrEmpty(s.Name))
                        .Select(s => s.Name.Trim())
                        .Distinct()
                        .ToList();

                    if (!dbPosSuggestions.Any())
                    {
                        var seedPositions = new List<string>
                        {
                            "Kiểm thu viên thuế",
                            "Đội trưởng",
                            "Phó Đội trưởng",
                            "Tổ trưởng",
                            "Phó Tổ trưởng",
                            "Chi cục trưởng",
                            "Phó Chi cục trưởng",
                            "Trưởng phòng",
                            "Phó Trưởng phòng",
                            "Trưởng Thuế cơ sở",
                            "Phó Trưởng Thuế cơ sở",
                            "Cục trưởng",
                            "Phó Cục trưởng",
                            "Công chức",
                            "Kiểm tra viên chính",
                            "Kiểm tra viên",
                            "Chuyên viên chính",
                            "Chuyên viên",
                            "Cán bộ",
                            "Nhân viên"
                        };

                        var dbPos = context.Positions
                            .Where(p => !string.IsNullOrEmpty(p.Name) && p.DepartmentName != "__ĐẢNG__" && p.DepartmentName != "__QUY_HOẠCH__")
                            .Select(p => p.Name.Trim())
                            .ToList();

                        var allSeedPos = seedPositions.Concat(dbPos)
                            .Where(x => !string.IsNullOrWhiteSpace(x))
                            .Distinct(StringComparer.OrdinalIgnoreCase)
                            .ToList();

                        foreach (var p in allSeedPos)
                        {
                            context.WorkHistorySuggestions.Add(new WorkHistorySuggestion { Category = "Position", Name = p });
                        }
                        context.SaveChanges();
                        dbPosSuggestions = allSeedPos;
                    }

                    _positionSuggestions = new ObservableCollection<string>(
                        dbPosSuggestions
                            .Where(x => !string.IsNullOrWhiteSpace(x))
                            .Distinct(StringComparer.OrdinalIgnoreCase)
                            .OrderBy(x => x)
                    );
                    cboPosition.ItemsSource = _positionSuggestions;

                    // 4. NẠP GỢI Ý CHỨC VỤ ĐẢNG (ĐỘC LẬP)
                    var dbPartySuggestions = context.WorkHistorySuggestions
                        .Where(s => s.Category == "PartyPosition" && !string.IsNullOrEmpty(s.Name))
                        .Select(s => s.Name.Trim())
                        .Distinct()
                        .ToList();

                    if (!dbPartySuggestions.Any())
                    {
                        var seedPartyPos = new List<string>
                        {
                            "Bí thư Chi bộ",
                            "Phó Bí thư Chi bộ",
                            "Chi ủy viên",
                            "Bí thư Đảng bộ",
                            "Phó Bí thư Đảng bộ",
                            "Đảng ủy viên",
                            "Bí thư Đảng ủy",
                            "Phó Bí thư Đảng ủy",
                            "Ủy viên Ban thường vụ Đảng ủy",
                            "Ủy viên Ban chấp hành Đảng ủy",
                            "Tổ trưởng Tổ Đảng",
                            "Đảng viên"
                        };

                        foreach (var pp in seedPartyPos)
                        {
                            context.WorkHistorySuggestions.Add(new WorkHistorySuggestion { Category = "PartyPosition", Name = pp });
                        }
                        context.SaveChanges();
                        dbPartySuggestions = seedPartyPos;
                    }

                    _partyPositionSuggestions = new ObservableCollection<string>(
                        dbPartySuggestions
                            .Where(x => !string.IsNullOrWhiteSpace(x))
                            .Select(x => {
                                var m = System.Text.RegularExpressions.Regex.Match(x.Trim(), @"^(?:Kiêm|Kiêm nhiệm|Kiêm chức vụ Đảng|Kiêm nhiệm chức vụ Đảng):\s*(.*)$", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                                return m.Success ? m.Groups[1].Value.Trim() : x.Trim();
                            })
                            .Distinct(StringComparer.OrdinalIgnoreCase)
                            .OrderBy(x => x)
                    );
                    cboPartyPosition.ItemsSource = _partyPositionSuggestions;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WorkHistoryDialog] LoadSuggestions error: {ex.Message}");
            }
        }

        private void btnDeleteDepartmentSuggestion_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            e.Handled = true;
            if (sender is FrameworkElement fe && fe.Tag is string name && !string.IsNullOrWhiteSpace(name))
            {
                try
                {
                    using (var db = new AppDbContext())
                    {
                        var items = db.WorkHistorySuggestions.Where(s => s.Category == "Department" && s.Name == name).ToList();
                        if (items.Any())
                        {
                            db.WorkHistorySuggestions.RemoveRange(items);
                            db.SaveChanges();
                        }
                    }
                    _departmentSuggestions.Remove(name);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[WorkHistoryDialog] Delete department suggestion error: {ex.Message}");
                }
            }
        }

        private void btnDeleteOrgSuggestion_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            e.Handled = true;
            if (sender is FrameworkElement fe && fe.Tag is string name && !string.IsNullOrWhiteSpace(name))
            {
                try
                {
                    using (var db = new AppDbContext())
                    {
                        var items = db.WorkHistorySuggestions.Where(s => s.Category == "OrganizationUnit" && s.Name == name).ToList();
                        if (items.Any())
                        {
                            db.WorkHistorySuggestions.RemoveRange(items);
                            db.SaveChanges();
                        }
                    }
                    _orgSuggestions.Remove(name);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[WorkHistoryDialog] Delete org suggestion error: {ex.Message}");
                }
            }
        }

        private void btnDeletePositionSuggestion_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            e.Handled = true;
            if (sender is FrameworkElement fe && fe.Tag is string name && !string.IsNullOrWhiteSpace(name))
            {
                try
                {
                    using (var db = new AppDbContext())
                    {
                        var items = db.WorkHistorySuggestions.Where(s => s.Category == "Position" && s.Name == name).ToList();
                        if (items.Any())
                        {
                            db.WorkHistorySuggestions.RemoveRange(items);
                            db.SaveChanges();
                        }
                    }
                    _positionSuggestions.Remove(name);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[WorkHistoryDialog] Delete position suggestion error: {ex.Message}");
                }
            }
        }

        private void btnDeletePartyPositionSuggestion_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            e.Handled = true;
            if (sender is FrameworkElement fe && fe.Tag is string name && !string.IsNullOrWhiteSpace(name))
            {
                try
                {
                    using (var db = new AppDbContext())
                    {
                        var items = db.WorkHistorySuggestions.Where(s => s.Category == "PartyPosition" && s.Name == name).ToList();
                        if (items.Any())
                        {
                            db.WorkHistorySuggestions.RemoveRange(items);
                            db.SaveChanges();
                        }
                    }
                    _partyPositionSuggestions.Remove(name);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[WorkHistoryDialog] Delete party position suggestion error: {ex.Message}");
                }
            }
        }

        private void btnDeleteSuggestion_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            e.Handled = true;
        }

        private void chkPartyConcurrent_CheckedChanged(object sender, RoutedEventArgs e)
        {
            bool isParty = chkPartyConcurrent.IsChecked == true;
            if (pnlPartyPosition != null)
            {
                pnlPartyPosition.Visibility = isParty ? Visibility.Visible : Visibility.Collapsed;
            }
            if (!isParty && cboPartyPosition != null)
            {
                cboPartyPosition.Text = "";
            }
        }

        private void chkPresent_CheckedChanged(object sender, RoutedEventArgs e)
        {
            UpdatePresentState(chkPresent.IsChecked == true);
        }

        private bool _isFormattingDate = false;

        private void InitializeYearPickers()
        {
            var years = Enumerable.Range(1960, 80).ToList();
            cboFromYear.ItemsSource = years;
            cboToYear.ItemsSource = years;
            cboFromYear.SelectedItem = DateTime.Now.Year;
            cboToYear.SelectedItem = DateTime.Now.Year;
        }

        #region Auto-slash & Month/Year Typing Masking

        private void DateTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (sender is not TextBox tb) return;

            // Chỉ cho phép nhập chữ số, '/' hoặc '-'
            foreach (char c in e.Text)
            {
                if (!char.IsDigit(c) && c != '/' && c != '-')
                {
                    e.Handled = true;
                    return;
                }
            }
        }

        private void DateTextBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (sender is not TextBox tb) return;

            if (e.Key == Key.Space)
            {
                e.Handled = true;
                return;
            }
        }

        private void DateTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_isFormattingDate) return;
            if (sender is TextBox tb)
            {
                ResetBorderValidation(tb);
            }
        }

        private void DateTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (sender is not TextBox tb) return;

            string text = tb.Text?.Trim() ?? "";
            if (string.IsNullOrEmpty(text) || text.Equals("Hiện nay", StringComparison.OrdinalIgnoreCase))
            {
                ResetBorderValidation(tb);
                return;
            }

            if (TryParseDateOrMonthYear(text, out string valid, out _, out _, out _))
            {
                _isFormattingDate = true;
                tb.Text = valid;
                _isFormattingDate = false;
                ResetBorderValidation(tb);
            }
            else
            {
                SetBorderError(tb, "Định dạng thời gian không hợp lệ (ví dụ: 16/01/2024, 08/2020 hoặc 2020)");
            }
        }

        private void ResetBorderValidation(TextBox tb)
        {
            Border? bdr = tb == txtFromDate ? bdrFromDate : bdrToDate;
            if (bdr != null)
            {
                bdr.BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#CBD5E1"));
                bdr.ToolTip = null;
            }
        }

        private void SetBorderError(TextBox tb, string errorMsg)
        {
            Border? bdr = tb == txtFromDate ? bdrFromDate : bdrToDate;
            if (bdr != null)
            {
                bdr.BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#EF4444"));
                bdr.ToolTip = errorMsg;
            }
        }

        public static bool TryParseDateOrMonthYear(string? input, out string formatted, out DateTime sortDate, out int month, out int year)
        {
            formatted = "";
            sortDate = DateTime.MinValue;
            month = 0;
            year = 0;

            if (string.IsNullOrWhiteSpace(input)) return false;

            string s = input.Trim().Replace('-', '/');

            // 1. Trường hợp có dấu '/'
            if (s.Contains("/"))
            {
                var parts = s.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length == 3)
                {
                    // Định dạng dd/MM/yyyy (hoặc d/M/yyyy)
                    if (int.TryParse(parts[0].Trim(), out int d) &&
                        int.TryParse(parts[1].Trim(), out int m) &&
                        int.TryParse(parts[2].Trim(), out int y))
                    {
                        if (y >= 1900 && y <= 2100 && m >= 1 && m <= 12)
                        {
                            int daysInMonth = DateTime.DaysInMonth(y, m);
                            if (d >= 1 && d <= daysInMonth)
                            {
                                month = m;
                                year = y;
                                formatted = $"{d:D2}/{m:D2}/{y:D4}";
                                sortDate = new DateTime(y, m, d);
                                return true;
                            }
                        }
                    }
                }
                else if (parts.Length == 2)
                {
                    // Định dạng MM/yyyy (hoặc M/yyyy)
                    if (int.TryParse(parts[0].Trim(), out int m) &&
                        int.TryParse(parts[1].Trim(), out int y))
                    {
                        if (y >= 1900 && y <= 2100 && m >= 1 && m <= 12)
                        {
                            month = m;
                            year = y;
                            formatted = $"{m:D2}/{y:D4}";
                            sortDate = new DateTime(y, m, 1);
                            return true;
                        }
                    }
                }
                return false;
            }

            // 2. Trường hợp chuỗi chỉ gồm các chữ số
            string digits = new string(s.Where(char.IsDigit).ToArray());
            if (digits.Length == 8)
            {
                // ddMMyyyy (ví dụ: 16012024)
                if (int.TryParse(digits.Substring(0, 2), out int d) &&
                    int.TryParse(digits.Substring(2, 2), out int m) &&
                    int.TryParse(digits.Substring(4, 4), out int y))
                {
                    if (y >= 1900 && y <= 2100 && m >= 1 && m <= 12)
                    {
                        int daysInMonth = DateTime.DaysInMonth(y, m);
                        if (d >= 1 && d <= daysInMonth)
                        {
                            month = m;
                            year = y;
                            formatted = $"{d:D2}/{m:D2}/{y:D4}";
                            sortDate = new DateTime(y, m, d);
                            return true;
                        }
                    }
                }
            }
            else if (digits.Length == 6)
            {
                // MMyyyy (ví dụ: 082020)
                if (int.TryParse(digits.Substring(0, 2), out int m) &&
                    int.TryParse(digits.Substring(2, 4), out int y))
                {
                    if (y >= 1900 && y <= 2100 && m >= 1 && m <= 12)
                    {
                        month = m;
                        year = y;
                        formatted = $"{m:D2}/{y:D4}";
                        sortDate = new DateTime(y, m, 1);
                        return true;
                    }
                }
            }
            else if (digits.Length == 4)
            {
                // yyyy (ví dụ: 2024)
                if (int.TryParse(digits, out int y) && y >= 1900 && y <= 2100)
                {
                    month = 1;
                    year = y;
                    formatted = $"{y:D4}";
                    sortDate = new DateTime(y, 1, 1);
                    return true;
                }
            }

            return false;
        }

        private static bool TryParseAndFormatMonthYear(string? input, out string formatted, out int month, out int year)
        {
            return TryParseDateOrMonthYear(input, out formatted, out _, out month, out year);
        }

        #endregion

        #region Month & Year Popup Picker Handlers

        private void btnPickFromDate_Click(object sender, RoutedEventArgs e)
        {
            popToDate.IsOpen = false;
            if (TryParseAndFormatMonthYear(txtFromDate.Text, out _, out int m, out int y))
            {
                cboFromYear.SelectedItem = y;
                HighlightSelectedMonth(gridFromMonths, m);
            }
            else
            {
                cboFromYear.SelectedItem = DateTime.Now.Year;
                HighlightSelectedMonth(gridFromMonths, 0);
            }
            popFromDate.IsOpen = !popFromDate.IsOpen;
        }

        private void btnPickToDate_Click(object sender, RoutedEventArgs e)
        {
            if (chkPresent.IsChecked == true) return;
            popFromDate.IsOpen = false;

            if (TryParseAndFormatMonthYear(txtToDate.Text, out _, out int m, out int y))
            {
                cboToYear.SelectedItem = y;
                HighlightSelectedMonth(gridToMonths, m);
            }
            else if (TryParseAndFormatMonthYear(txtFromDate.Text, out _, out _, out int fromYear))
            {
                cboToYear.SelectedItem = fromYear;
                HighlightSelectedMonth(gridToMonths, 0);
            }
            else
            {
                cboToYear.SelectedItem = DateTime.Now.Year;
                HighlightSelectedMonth(gridToMonths, 0);
            }
            popToDate.IsOpen = !popToDate.IsOpen;
        }

        private void FromMonthButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is string tagStr && int.TryParse(tagStr, out int m))
            {
                int y = (cboFromYear.SelectedItem is int yearVal) ? yearVal : DateTime.Now.Year;
                _isFormattingDate = true;
                txtFromDate.Text = $"{m:D2}/{y}";
                _isFormattingDate = false;
                ResetBorderValidation(txtFromDate);
                popFromDate.IsOpen = false;
            }
        }

        private void ToMonthButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is string tagStr && int.TryParse(tagStr, out int m))
            {
                int y = (cboToYear.SelectedItem is int yearVal) ? yearVal : DateTime.Now.Year;
                chkPresent.IsChecked = false;
                UpdatePresentState(false);
                _isFormattingDate = true;
                txtToDate.Text = $"{m:D2}/{y}";
                _isFormattingDate = false;
                ResetBorderValidation(txtToDate);
                popToDate.IsOpen = false;
            }
        }

        private void btnFromPrev1_Click(object sender, RoutedEventArgs e)
        {
            if (cboFromYear.SelectedItem is int y)
            {
                cboFromYear.SelectedItem = Math.Max(1960, y - 1);
            }
        }

        private void btnFromNext1_Click(object sender, RoutedEventArgs e)
        {
            if (cboFromYear.SelectedItem is int y)
            {
                cboFromYear.SelectedItem = Math.Min(2039, y + 1);
            }
        }

        private void btnFromPrev5_Click(object sender, RoutedEventArgs e)
        {
            if (cboFromYear.SelectedItem is int y)
            {
                cboFromYear.SelectedItem = Math.Max(1960, y - 5);
            }
        }

        private void btnFromNext5_Click(object sender, RoutedEventArgs e)
        {
            if (cboFromYear.SelectedItem is int y)
            {
                cboFromYear.SelectedItem = Math.Min(2039, y + 5);
            }
        }

        private void btnToPrev1_Click(object sender, RoutedEventArgs e)
        {
            if (cboToYear.SelectedItem is int y)
            {
                cboToYear.SelectedItem = Math.Max(1960, y - 1);
            }
        }

        private void btnToNext1_Click(object sender, RoutedEventArgs e)
        {
            if (cboToYear.SelectedItem is int y)
            {
                cboToYear.SelectedItem = Math.Min(2039, y + 1);
            }
        }

        private void btnToPrev5_Click(object sender, RoutedEventArgs e)
        {
            if (cboToYear.SelectedItem is int y)
            {
                cboToYear.SelectedItem = Math.Max(1960, y - 5);
            }
        }

        private void btnToNext5_Click(object sender, RoutedEventArgs e)
        {
            if (cboToYear.SelectedItem is int y)
            {
                cboToYear.SelectedItem = Math.Min(2039, y + 5);
            }
        }

        private void cboFromYear_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (TryParseAndFormatMonthYear(txtFromDate.Text, out _, out int m, out int y))
            {
                if (cboFromYear.SelectedItem is int selYear && selYear == y)
                {
                    HighlightSelectedMonth(gridFromMonths, m);
                    return;
                }
            }
            HighlightSelectedMonth(gridFromMonths, 0);
        }

        private void cboToYear_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (TryParseAndFormatMonthYear(txtToDate.Text, out _, out int m, out int y))
            {
                if (cboToYear.SelectedItem is int selYear && selYear == y)
                {
                    HighlightSelectedMonth(gridToMonths, m);
                    return;
                }
            }
            HighlightSelectedMonth(gridToMonths, 0);
        }

        private void btnFromCurrentMonth_Click(object sender, RoutedEventArgs e)
        {
            _isFormattingDate = true;
            txtFromDate.Text = $"{DateTime.Now.Month:D2}/{DateTime.Now.Year}";
            _isFormattingDate = false;
            ResetBorderValidation(txtFromDate);
            popFromDate.IsOpen = false;
        }

        private void btnToCurrentMonth_Click(object sender, RoutedEventArgs e)
        {
            chkPresent.IsChecked = false;
            UpdatePresentState(false);
            _isFormattingDate = true;
            txtToDate.Text = $"{DateTime.Now.Month:D2}/{DateTime.Now.Year}";
            _isFormattingDate = false;
            ResetBorderValidation(txtToDate);
            popToDate.IsOpen = false;
        }

        private void btnFromClear_Click(object sender, RoutedEventArgs e)
        {
            _isFormattingDate = true;
            txtFromDate.Text = "";
            _isFormattingDate = false;
            ResetBorderValidation(txtFromDate);
            popFromDate.IsOpen = false;
        }

        private void btnToClear_Click(object sender, RoutedEventArgs e)
        {
            _isFormattingDate = true;
            txtToDate.Text = "";
            _isFormattingDate = false;
            ResetBorderValidation(txtToDate);
            popToDate.IsOpen = false;
        }

        private void HighlightSelectedMonth(UniformGrid grid, int selectedMonth)
        {
            if (grid == null) return;
            foreach (var child in grid.Children)
            {
                if (child is Button b && b.Tag is string tagStr && int.TryParse(tagStr, out int m))
                {
                    if (m == selectedMonth)
                    {
                        b.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0284C7"));
                        b.Foreground = Brushes.White;
                        b.FontWeight = FontWeights.Bold;
                        b.BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0284C7"));
                    }
                    else
                    {
                        b.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F8FAFC"));
                        b.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#334155"));
                        b.FontWeight = FontWeights.Medium;
                        b.BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E2E8F0"));
                    }
                }
            }
        }

        #endregion

        private void UpdatePresentState(bool isPresent)
        {
            if (bdrToDate == null || txtToDate == null) return;

            if (btnPickToDate != null)
            {
                btnPickToDate.IsEnabled = !isPresent;
            }

            if (isPresent)
            {
                popToDate.IsOpen = false;
                txtToDate.Text = "Hiện nay";
                txtToDate.IsEnabled = false;
                txtToDate.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0284C7"));
                txtToDate.FontWeight = FontWeights.Bold;
                bdrToDate.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F0F9FF"));
                bdrToDate.BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#BAE6FD"));
                ResetBorderValidation(txtToDate);
            }
            else
            {
                if (txtToDate.Text == "Hiện nay")
                {
                    txtToDate.Text = "";
                }
                txtToDate.IsEnabled = true;
                txtToDate.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0F172A"));
                txtToDate.FontWeight = FontWeights.Normal;
                bdrToDate.Background = Brushes.White;
                bdrToDate.BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#CBD5E1"));
            }
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            string fromDateRaw = txtFromDate.Text?.Trim() ?? "";
            string position = cboPosition.Text?.Trim() ?? "";
            string toDateRaw = chkPresent.IsChecked == true ? "Hiện nay" : (txtToDate.Text?.Trim() ?? "");

            if (string.IsNullOrWhiteSpace(fromDateRaw))
            {
                var warn = new WarningWindow("Vui lòng nhập Thời gian bắt đầu (Từ tháng/năm)!", "Thiếu thông tin") { Owner = this };
                warn.ShowDialog();
                txtFromDate.Focus();
                return;
            }

            if (!TryParseDateOrMonthYear(fromDateRaw, out string fromDate, out DateTime fromSortDate, out _, out _))
            {
                var warn = new WarningWindow("Thời gian bắt đầu công tác không hợp lệ!\nVui lòng nhập đúng định dạng ngày/tháng/năm (ví dụ: 16/01/2024), tháng/năm (ví dụ: 08/2015) hoặc năm (ví dụ: 2020).", "Lỗi nhập liệu") { Owner = this };
                warn.ShowDialog();
                SetBorderError(txtFromDate, "Thời gian không hợp lệ");
                txtFromDate.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(position))
            {
                var warn = new WarningWindow("Vui lòng nhập hoặc chọn Chức vụ, chức danh công tác!", "Thiếu thông tin") { Owner = this };
                warn.ShowDialog();
                cboPosition.Focus();
                return;
            }

            string toDate = "Hiện nay";
            if (chkPresent.IsChecked != true && !string.IsNullOrWhiteSpace(toDateRaw))
            {
                if (!TryParseDateOrMonthYear(toDateRaw, out string validTo, out DateTime toSortDate, out _, out _))
                {
                    var warn = new WarningWindow("Thời gian kết thúc công tác không hợp lệ!\nVui lòng nhập đúng định dạng ngày/tháng/năm (ví dụ: 28/02/2025), tháng/năm (ví dụ: 09/2021) hoặc tích chọn 'Đang giữ chức (Đến nay)'.", "Lỗi nhập liệu") { Owner = this };
                    warn.ShowDialog();
                    SetBorderError(txtToDate, "Thời gian không hợp lệ");
                    txtToDate.Focus();
                    return;
                }
                toDate = validTo;

                // Kiểm tra logic: Đến ngày không được trước Từ ngày
                if (toSortDate < fromSortDate)
                {
                    var warn = new WarningWindow($"Thời gian kết thúc ({toDate}) không thể trước thời gian bắt đầu ({fromDate})!", "Mốc thời gian không hợp lệ") { Owner = this };
                    warn.ShowDialog();
                    SetBorderError(txtToDate, "Thời gian kết thúc phải sau hoặc bằng thời gian bắt đầu");
                    txtToDate.Focus();
                    return;
                }
            }

            string dept = cboDepartment.Text?.Trim() ?? "";
            string org = cboOrganizationUnit.Text?.Trim() ?? "";

            WorkHistory.FromDate = fromDate;
            WorkHistory.ToDate = toDate;
            WorkHistory.Position = position;
            WorkHistory.Department = dept;
            WorkHistory.OrganizationUnit = org;
            WorkHistory.IsPartyConcurrent = chkPartyConcurrent.IsChecked == true;
            string cleanSavedParty = chkPartyConcurrent.IsChecked == true ? (cboPartyPosition.Text?.Trim() ?? "") : "";
            if (!string.IsNullOrWhiteSpace(cleanSavedParty))
            {
                var m = System.Text.RegularExpressions.Regex.Match(cleanSavedParty, @"^(?:Kiêm|Kiêm nhiệm|Kiêm chức vụ Đảng|Kiêm nhiệm chức vụ Đảng):\s*(.*)$", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                if (m.Success) cleanSavedParty = m.Groups[1].Value.Trim();
            }
            WorkHistory.PartyPosition = chkPartyConcurrent.IsChecked == true ? cleanSavedParty : null;

            if (!string.IsNullOrWhiteSpace(dept) && !string.IsNullOrWhiteSpace(org))
            {
                WorkHistory.DepartmentOrUnit = $"{dept}, {org}";
            }
            else if (!string.IsNullOrWhiteSpace(dept))
            {
                WorkHistory.DepartmentOrUnit = dept;
            }
            else if (!string.IsNullOrWhiteSpace(org))
            {
                WorkHistory.DepartmentOrUnit = org;
            }
            else
            {
                WorkHistory.DepartmentOrUnit = "";
            }

            WorkHistory.Note = txtNote.Text?.Trim();
            WorkHistory.NotifyChanged();

            SaveSuggestionIfNew("Department", dept);
            SaveSuggestionIfNew("OrganizationUnit", org);
            SaveSuggestionIfNew("Position", position);
            if (WorkHistory.IsPartyConcurrent && !string.IsNullOrWhiteSpace(WorkHistory.PartyPosition))
            {
                SaveSuggestionIfNew("PartyPosition", WorkHistory.PartyPosition);
            }

            DialogResult = true;
            Close();
        }

        private void SaveSuggestionIfNew(string category, string? value)
        {
            if (string.IsNullOrWhiteSpace(value)) return;
            string trimmed = value.Trim();

            try
            {
                using (var db = new AppDbContext())
                {
                    bool exists = db.WorkHistorySuggestions.Any(s => s.Category == category && s.Name == trimmed);
                    if (!exists)
                    {
                        db.WorkHistorySuggestions.Add(new WorkHistorySuggestion
                        {
                            Category = category,
                            Name = trimmed
                        });
                        db.SaveChanges();
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WorkHistoryDialog] SaveSuggestionIfNew error: {ex.Message}");
            }
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                DragMove();
            }
        }
    }
}
