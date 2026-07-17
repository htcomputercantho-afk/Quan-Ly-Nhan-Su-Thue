using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using MaterialDesignThemes.Wpf;
using Microsoft.EntityFrameworkCore;
using TaxPersonnelManagement.Data;
using TaxPersonnelManagement.Models;
namespace TaxPersonnelManagement.Views
{
    public partial class PersonnelDetailView : UserControl
    {
        private Personnel? _personnel;
        private string? _currentAvatarBase64;
        private bool _isAvatarChanged = false;
        private bool _isRefreshing = false;
        private bool _isFormatting = false;
        private List<Position> _allPositions = new List<Position>();
        private SalaryRecord? _editingSalaryRecord;
        private EvaluationRecord? _editingEvaluationRecord;

        /// <summary>
        /// Khởi tạo màn hình chi tiết hồ sơ cán bộ.
        /// </summary>
        /// <param name="personnel">Đối tượng cán bộ cần hiển thị. Nếu thêm mới thì truyền null.</param>
        public PersonnelDetailView(Personnel? personnel, int activeTab = 0)
        {
            InitializeComponent();
            dpPositionCalculationDate.SelectedDate = DateTime.Now;
            if (activeTab > 0) tcPersonnelDetails.SelectedIndex = activeTab;
            try { LoadSalaryDelayReasons(); } catch { /* Ignore DB init errors */ }
            try { LoadDisciplineTypes(); } catch { /* Ignore DB init errors */ }
            _personnel = personnel;

            _isRefreshing = true; // NGĂN CHẶN TÍNH TOÁN TỰ ĐỘNG TRONG QUÁ TRÌNH NẠP DỮ LIỆU BAN ĐẦU

            // Load Metadata First
            LoadDepartments();
            LoadPositions();
            LoadRanks();

            if (_personnel != null)
            {
                // Load Data
                txtStaffId.Text = _personnel.StaffId;
                txtName.Text = _personnel.FullName;
                SetComboBoxByContent(cboGender, _personnel.Gender);
                dpDOB.SelectedDate = _personnel.DateOfBirth;
                txtPhone.Text = FormatPhoneNumber(_personnel.PhoneNumber);
                txtIdentityNumber.Text = _personnel.IdentityCardNumber;
                txtIdentityPlace.Text = _personnel.IdentityCardPlace;
                txtEmail.Text = _personnel.Email;
                txtSocialInsurance.Text = _personnel.SocialSecurityNumber;
                txtBirthPlace.Text = _personnel.BirthPlace;
                txtHometown.Text = _personnel.Hometown;
                txtCurrentResidence.Text = _personnel.CurrentResidence;
                txtEthnicity.Text = _personnel.Ethnicity;
                txtReligion.Text = _personnel.Religion;

                // Work
                // cboDepartment.Text handled in LoadDepartments which sets SelectedItem if _personnel is valid, 
                // BUT we just moved LoadDepartments up. 
                // LoadDepartments has logic: if (_personnel != null) cboDepartment.SelectedItem = ...
                // Since _personnel is set, it might work inside the methods.
                // However, the original code had manual setting inside LoadDepartments.
                // Let's verify LoadDepartments logic below.

                // If the Load methods ALREADY set the values using _personnel, we don't need to set them again here?
                // LoadDepartments: uses _personnel.Department.
                // LoadPositions: uses _personnel.Position.
                // LoadRanks: uses _personnel.RankCode.

                // So lines 41-44 might be redundant OR need to be careful not to overwrite if Load didn't work.
                // The explicit assignments here (e.g. cboPosition.Text) are good backups or primary if Load methods didn't set Text property.
                // Actually, let's keep explicit assignments for safety, but use SelectedValue were appropriate.

                // cboDepartment is set inside LoadDepartments.
                // cboPosition is set inside LoadPositions.

                txtRankName.Text = _personnel.RankName;
                dpStartDate.SelectedDate = _personnel.TaxAuthorityStartDate;


                // Education
                SetComboBoxByContent(cboEducationLevel, _personnel.EducationLevel);
                txtMajor.Text = _personnel.Major;
                txtUniversity.Text = _personnel.University;
                SetComboBoxByContent(cboStateManagement, _personnel.StateManagementLevel);
                SetComboBoxByContent(cboPoliticalTheory, _personnel.PoliticalTheoryLevel);
                txtITSkill.Text = _personnel.ITSkillLevel;

                txtLanguageSkill.Text = _personnel.LanguageSkillLevel;

                // Tab 2
                dpPositionDecisionDate.SelectedDate = _personnel.PositionDecisionDate;
                dpPositionCalculationDate.SelectedDate = DateTime.Now;
                txtPositionYear.Text = _personnel.PositionYear;
                txtDetailedWorkHistory.Text = _personnel.DetailedWorkHistory;

                CalculateWorkDuration();

                // Tab 3
                dpRetirementDate.SelectedDate = _personnel.RetirementDate;
                CalculateRetirementInfo();

                // Tab 4
                dpPartyEntryDate.SelectedDate = _personnel.PartyEntryDate;
                dpPartyOfficialDate.SelectedDate = _personnel.PartyOfficialDate;
                CalculatePartyAge();

                // Tab 6: Salary Info
                cboSalaryStep.SelectedValue = _personnel.CurrentSalaryStep; // Use SelectedValue
                txtSalaryCoefficient.Text = _personnel.CurrentSalaryCoefficient.ToString();
                txtExceedFrame.Text = _personnel.ExceedFramePercent.ToString();
                txtPositionAllowance.Text = _personnel.PositionAllowance;
                dpSalaryReservation.SelectedDate = _personnel.SalaryReservationDeadline;
                dpNextSalaryStepDate.SelectedDate = _personnel.NextSalaryStepDate;

                SetSalaryDelaySelection(_personnel.SalaryIncreaseDelayType);

                dpExpectedSalaryIncrease.SelectedDate = _personnel.ExpectedSalaryIncreaseDate;
                // Load Salary Records
                if (_personnel.SalaryRecords == null) _personnel.SalaryRecords = new System.Collections.Generic.List<SalaryRecord>();
                RefreshSalaryHistoryGrid();

                // Tab 5: Leave Info
                dpLeaveCalculationDate.SelectedDate = _personnel.LeaveCalculationDate;
                // Ensure list is not null
                if (_personnel.LeaveHistories == null) _personnel.LeaveHistories = new System.Collections.Generic.List<LeaveHistory>();
                RefreshLeaveHistoryGrid();
                // UpdateLeaveStatistics(); // Handled by CalculateAnnualLeave via CalculateWorkDuration

                // Tab 7: Reward Info
                txtEmulationTitles.Text = _personnel.EmulationTitles;
                txtRewardForms.Text = _personnel.RewardForms;

                // Tab 8: Discipline Info
                if (!string.IsNullOrEmpty(_personnel.DisciplineType))
                    SetComboBoxByContent(cboDisciplineType, _personnel.DisciplineType);

                txtDisciplineNumber.Text = _personnel.DisciplineDecisionNumber;
                dpDisciplineDate.SelectedDate = _personnel.DisciplineDecisionDate;
                txtDisciplineReason.Text = _personnel.DisciplineReason;

                // Load Evaluation Records
                if (_personnel.EvaluationRecords == null) _personnel.EvaluationRecords = new System.Collections.Generic.List<EvaluationRecord>();
                RefreshEvaluationHistoryGrid();

                // Load Avatar
                if (!string.IsNullOrEmpty(_personnel.AvatarBase64))
                {
                    try
                    {
                        var bitmap = Base64ToImage(_personnel.AvatarBase64);
                        imgAvatar.Fill = new System.Windows.Media.ImageBrush(bitmap) { Stretch = System.Windows.Media.Stretch.UniformToFill };
                        iconAvatarPlaceholder.Visibility = Visibility.Collapsed;
                        btnRemoveAvatar.Visibility = Visibility.Visible;
                    }
                    catch { /* Ignore invalid image data */ }
                }
            }
            else
            {
                txtIdentityPlace.Text = "Cục Cảnh sát quản lý hành chính về trật tự xã hội";
            }

            if (_personnel != null)
            {
                lblSaveText.Text = "Cập nhật";
            }
            else
            {
                lblSaveText.Text = "Lưu Hồ Sơ";
            }

            // Initial Calculation to ensure consistency
            if (_personnel == null || !dpExpectedSalaryIncrease.SelectedDate.HasValue)
            {
                // Chỉ tự động tính ngày dự kiến lên lương nếu thêm mới
                // hoặc ngày đã lưu bị rỗng
                bool wasRefreshing = _isRefreshing;
                _isRefreshing = false; // Tạm mở khóa để hàm cập nhật
                CalculateExpectedSalaryDate();
                _isRefreshing = wasRefreshing;
            }

            CalculateAnnualLeave(DateTime.Now); // Force recalculation with valid _personnel
            RefreshLeaveHistoryGrid(); // Populate and clean leave history

            // Initialize Year ComboBox
            var currentYear = DateTime.Now.Year;
            var startYear = 2025;
            var years = new List<int>();
            for (int i = startYear; i <= currentYear; i++) years.Add(i);
            cboLeaveYear.ItemsSource = years;
            cboLeaveYear.SelectedItem = currentYear;

            // Event for Year Change
            cboLeaveYear.SelectionChanged += (s, e) => { UpdateLeaveStatistics(); };

            LoadLeaveYears();
            LoadEvaluationYears();
            ApplyAuthorization();

            _isRefreshing = false; // HOÀN TẤT NẠP DỮ LIỆU BAN ĐẦU
            CalculateExpectedSalaryDate();
        }

        /// <summary>
        /// Phân quyền hiển thị các chức năng thêm/sửa/xóa tùy theo Role của người dùng đăng nhập.
        /// </summary>
        private void ApplyAuthorization()
        {
            if (App.CurrentUser?.Role == UserRole.Staff)
            {
                // Hide Save button
                btnSave.Visibility = Visibility.Collapsed;

                // Hide config/add buttons
                btnAddDepartment.Visibility = Visibility.Collapsed;
                btnAddPosition.Visibility = Visibility.Collapsed;
                btnAddRank.Visibility = Visibility.Collapsed;

                // Hide avatar buttons (if they exist as explicit named buttons)
                btnRemoveAvatar.Visibility = Visibility.Collapsed;

                // Hide Add Leave button
                btnAddLeave.Visibility = Visibility.Collapsed;
            }
        }

        private void AdminOnly_Loaded(object? sender, RoutedEventArgs e)
        {
            if (App.CurrentUser?.Role == UserRole.Staff && sender is FrameworkElement element)
            {
                element.Visibility = Visibility.Collapsed;
            }
        }

        private void LoadLeaveYears(int? includeYear = null)
        {
            if (cboLeaveYear == null) return;

            var currentYear = DateTime.Now.Year;
            var startYear = 2025;
            var years = new List<int>();

            // Logic: Iterate 2025 to CurrentYear
            for (int i = startYear; i <= currentYear; i++)
            {
                // Always add current year
                if (i == currentYear)
                {
                    years.Add(i);
                    continue;
                }

                // Add if explicitly included (e.g. editing old record)
                if (includeYear.HasValue && i == includeYear.Value)
                {
                    years.Add(i);
                    continue;
                }

                // Add if has remaining days
                if (GetRemainingAnnualLeaveDays(i) > 0)
                {
                    years.Add(i);
                }
            }

            // Preserve selection if possible
            var previouslySelected = cboLeaveYear.SelectedItem;

            cboLeaveYear.ItemsSource = years;

            if (previouslySelected is int val && years.Contains(val))
            {
                cboLeaveYear.SelectedItem = val;
            }
            else
            {
                cboLeaveYear.SelectedItem = currentYear;
            }
        }

        private void LoadDepartments()
        {
            var deptOrder = new System.Collections.Generic.List<string> {
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

            using (var context = new AppDbContext())
            {
                // Load from Departments table
                var dbDepts = context.Departments.Select(d => d.Name).ToList();

                // Thêm bộ phận hiện tại của nhân sự đang xem (nếu chưa có trong danh mục)
                if (_personnel != null && !string.IsNullOrEmpty(_personnel.Department) && !dbDepts.Contains(_personnel.Department))
                {
                    dbDepts.Add(_personnel.Department);
                }

                var allDepts = dbDepts.Where(x => !string.IsNullOrEmpty(x))
                                      .Distinct()
                                      .OrderBy(x =>
                                      {
                                          int idx = deptOrder.FindIndex(d => d.Equals(x, StringComparison.OrdinalIgnoreCase));
                                          return idx == -1 ? 999 : idx;
                                      })
                                      .ThenBy(x => x)
                                      .ToList();
                allDepts.Insert(0, "");

                cboDepartment.ItemsSource = allDepts;

                if (_personnel != null)
                {
                    SetComboBoxByContent(cboDepartment, _personnel.Department);
                }
            }
        }

        private void cboDepartment_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isRefreshing) return;
            UpdatePositionFilter();
        }

        private void UpdatePositionFilter()
        {
            if (cboDepartment == null || cboPosition == null) return;

            string selectedDept = "";
            if (cboDepartment.SelectedItem != null)
            {
                selectedDept = cboDepartment.SelectedItem.ToString() ?? "";
            }
            else if (!string.IsNullOrEmpty(cboDepartment.Text))
            {
                selectedDept = cboDepartment.Text;
            }

            System.Collections.Generic.List<string> filtered;

            if (string.IsNullOrEmpty(selectedDept))
            {
                filtered = new System.Collections.Generic.List<string>();
            }
            else
            {
                var leadershipPositions = new[] { "Trưởng Thuế cơ sở", "Quyền Trưởng Thuế cơ sở", "Phó Trưởng Thuế cơ sở" };
                var unitPositions = new[] { "Tổ trưởng", "Phó Tổ trưởng", "Công chức" };

                filtered = _allPositions.Where(pos =>
                {
                    if (!string.IsNullOrEmpty(pos.DepartmentName))
                    {
                        if (pos.DepartmentName.Equals(selectedDept, StringComparison.OrdinalIgnoreCase))
                            return true;

                        if (pos.DepartmentName.Equals("Các tổ", StringComparison.OrdinalIgnoreCase) &&
                            selectedDept.Contains("Tổ", StringComparison.OrdinalIgnoreCase))
                            return true;

                        return false;
                    }

                    // Fallback for default positions
                    if (leadershipPositions.Any(lp => lp.Equals(pos.Name, StringComparison.OrdinalIgnoreCase)))
                    {
                        return selectedDept.Equals("Ban lãnh đạo", StringComparison.OrdinalIgnoreCase);
                    }

                    if (unitPositions.Any(up => up.Equals(pos.Name, StringComparison.OrdinalIgnoreCase)))
                    {
                        return selectedDept.Contains("Tổ", StringComparison.OrdinalIgnoreCase);
                    }

                    // Custom positions with empty DepartmentName apply to all
                    return true;
                })
                .Select(pos => pos.Name)
                .ToList();
            }

            var currentPos = cboPosition.SelectedItem?.ToString() ?? (_personnel?.Position);
            filtered.Insert(0, "");
            cboPosition.ItemsSource = filtered;

            if (!string.IsNullOrEmpty(currentPos) && filtered.Any(f => f.Equals(currentPos, StringComparison.OrdinalIgnoreCase)))
            {
                SetComboBoxByContent(cboPosition, currentPos);
            }
            else
            {
                cboPosition.SelectedIndex = 0;
            }
        }


        private void btnAddDepartment_Click(object? sender, RoutedEventArgs e)
        {
            var dialog = new AddDepartmentDialog(); // No args
            if (dialog.ShowDialog() == true)
            {
                LoadDepartments(); // Reload list from DB

                if (!string.IsNullOrEmpty(dialog.SelectedDepartment))
                {
                    cboDepartment.SelectedItem = dialog.SelectedDepartment;
                    cboDepartment.Text = dialog.SelectedDepartment;
                }
            }
        }

        private void btnConfigSalary_Click(object? sender, RoutedEventArgs e)
        {
            var dialog = new SalaryConfigDialog();
            dialog.Owner = Window.GetWindow(this);
            dialog.ShowDialog();

            // Refresh steps for current rank if selected
            if (cboRankCode.SelectedItem is Rank selectedRank && !string.IsNullOrEmpty(selectedRank.Code))
            {
                LoadSalarySteps(selectedRank.Code);
            }
        }

        private void LoadPositions()
        {
            var posOrder = new System.Collections.Generic.List<string> {
                "Chi cục trưởng",
                "Quyền Chi cục trưởng",
                "Phó Chi cục trưởng",
                "Trưởng Thuế cơ sở",
                "Quyền Trưởng Thuế cơ sở",
                "Phó Trưởng Thuế cơ sở",
                "Đội trưởng",
                "Trưởng phòng",
                "Phó Đội trưởng",
                "Phó Trưởng phòng",
                "Tổ trưởng",
                "Phó Tổ trưởng",
                "Công chức",
                "Nhân viên"
            };

            using (var context = new AppDbContext())
            {
                var dbPos = context.Positions.ToList();

                // Thêm chức vụ hiện tại của nhân sự đang xem nếu chưa có trong danh mục chính
                if (_personnel != null && !string.IsNullOrEmpty(_personnel.Position) && !dbPos.Any(p => p.Name.Equals(_personnel.Position, StringComparison.OrdinalIgnoreCase)))
                {
                    dbPos.Add(new Position { Name = _personnel.Position });
                }

                _allPositions = dbPos.Where(x => !string.IsNullOrEmpty(x.Name) && x.DepartmentName != "__ĐẢNG__")
                                  .GroupBy(x => x.Name.ToLower())
                                  .Select(g => g.First())
                                  .OrderBy(x =>
                                  {
                                      int idx = posOrder.FindIndex(p => p.Equals(x.Name, StringComparison.OrdinalIgnoreCase));
                                      return idx == -1 ? 999 : idx;
                                  })
                                  .ThenBy(x => x.Name)
                                  .ToList();

                UpdatePositionFilter();
            }
        }

        private void btnAddPosition_Click(object? sender, RoutedEventArgs e)
        {
            var dialog = new AddPositionDialog();
            if (dialog.ShowDialog() == true)
            {
                LoadPositions();

                if (!string.IsNullOrEmpty(dialog.SelectedPosition))
                {
                    cboPosition.SelectedItem = dialog.SelectedPosition;
                    cboPosition.Text = dialog.SelectedPosition;
                }
            }
        }

        private void LoadRanks()
        {
            using (var context = new AppDbContext())
            {
                var dbRanks = context.Ranks.ToList();

                // Nếu nhân sự đang xem có mã ngạch không tồn tại trong danh mục chính, thêm mã ngạch đó vào danh sách hiển thị
                if (_personnel != null && !string.IsNullOrEmpty(_personnel.RankCode) && !dbRanks.Any(r => r.Code == _personnel.RankCode))
                {
                    dbRanks.Add(new Rank { Code = _personnel.RankCode, Name = _personnel.RankName ?? "" });
                }

                var sortedRanks = dbRanks.OrderBy(r => r.Code).ToList();
                sortedRanks.Insert(0, new Rank { Code = "", Name = "" });

                cboRankCode.ItemsSource = sortedRanks;
                cboRankCode.DisplayMemberPath = "Code";
                cboRankCode.SelectedValuePath = "Code";

                // Also populate the history input section
                cboSalaryHistRankCode.ItemsSource = sortedRanks;
                cboSalaryHistRankCode.DisplayMemberPath = "Code";
                cboSalaryHistRankCode.SelectedValuePath = "Code";

                if (_personnel != null)
                {
                    cboRankCode.SelectedValue = _personnel.RankCode;
                    if (!string.IsNullOrEmpty(_personnel.RankCode))
                    {
                        LoadSalarySteps(_personnel.RankCode);
                    }
                }
            }
        }


        private void LoadSalaryDelayReasons()
        {
            try
            {
                using (var context = new AppDbContext())
                {
                    try
                    {
                        var reasons = context.SalaryDelayReasons.OrderBy(r => r.Id).Select(r => r.Name).ToList();

                        cboSalaryDelay.Items.Clear();
                        cboSalaryDelay.Items.Add("-- Không lùi --");
                        foreach (var r in reasons)
                        {
                            if (!r.Contains("Nghỉ không lương", StringComparison.OrdinalIgnoreCase))
                            {
                                cboSalaryDelay.Items.Add(r);
                            }
                        }
                    }
                    catch (Exception)
                    {
                        // Fallback: Table might not exist, create it manually
                        // Use FRESH context
                        using (var repairContext = new AppDbContext())
                        {
                            // 1. Create SQL
                            repairContext.Database.ExecuteSqlRaw(@"
                                CREATE TABLE IF NOT EXISTS SalaryDelayReasons (
                                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                                    Name TEXT NOT NULL
                                );");

                            // 2. Check and Insert using SQL
                            var count = -1;
                            using (var command = repairContext.Database.GetDbConnection().CreateCommand())
                            {
                                command.CommandText = "SELECT COUNT(*) FROM SalaryDelayReasons";
                                repairContext.Database.OpenConnection();
                                try
                                {
                                    var result = command.ExecuteScalar();
                                    if (result != null) count = Convert.ToInt32(result);
                                }
                                catch { }
                                repairContext.Database.CloseConnection();
                            }

                            if (count == 0)
                            {
                                repairContext.Database.ExecuteSqlRaw("INSERT INTO SalaryDelayReasons (Id, Name) VALUES (1, 'Lùi 3 tháng (Khiển trách)')");
                                repairContext.Database.ExecuteSqlRaw("INSERT INTO SalaryDelayReasons (Id, Name) VALUES (2, 'Lùi 6 tháng (Cảnh cáo)')");
                                repairContext.Database.ExecuteSqlRaw("INSERT INTO SalaryDelayReasons (Id, Name) VALUES (3, 'Lùi 12 tháng (Giáng chức/Cách chức)')");
                                repairContext.Database.ExecuteSqlRaw("INSERT INTO SalaryDelayReasons (Id, Name) VALUES (4, 'Nghỉ không lương')");
                            }
                        }

                        // Retry load with ANOTHER fresh context
                        using (var retryContext = new AppDbContext())
                        {
                            var retryReasons = retryContext.SalaryDelayReasons.OrderBy(r => r.Id).Select(r => r.Name).ToList();
                            cboSalaryDelay.Items.Clear();
                            cboSalaryDelay.Items.Add("-- Không lùi --");
                            foreach (var r in retryReasons)
                            {
                                if (!r.Contains("Nghỉ không lương", StringComparison.OrdinalIgnoreCase))
                                {
                                    cboSalaryDelay.Items.Add(r);
                                }
                            }
                        }
                    }

                    if (_personnel != null)
                    {
                        SetSalaryDelaySelection(_personnel.SalaryIncreaseDelayType);
                    }
                    else
                    {
                        SetSalaryDelaySelection(null);
                    }
                }
            }
            catch (Exception ex)
            {
                // Last resort: Just show default item if totally borked
                cboSalaryDelay.Items.Clear();
                cboSalaryDelay.Items.Add("-- Không lùi --");
                cboSalaryDelay.Items.Add("Lùi 3 tháng (Khiển trách)");
                cboSalaryDelay.Items.Add("Lùi 6 tháng (Cảnh cáo)");
                cboSalaryDelay.Items.Add("Lùi 12 tháng (Giáng chức/Cách chức)");
                System.Diagnostics.Debug.WriteLine($"Error loading salary delay reasons: {ex.Message}");
            }
        }

        private void btnSalaryDelayConfig_Click(object? sender, RoutedEventArgs e)
        {
            var dialog = new SalaryDelayConfigDialog();
            dialog.Owner = Window.GetWindow(this);
            dialog.ShowDialog();
            LoadSalaryDelayReasons(); // Refresh after close
        }

        private void cboRankCode_SelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (cboRankCode.SelectedItem is Rank selectedRank && !string.IsNullOrEmpty(selectedRank.Code))
            {
                if (!string.IsNullOrEmpty(selectedRank.Name))
                    txtRankName.Text = selectedRank.Name;
                else
                    txtRankName.Text = "";

                LoadSalarySteps(selectedRank.Code);
            }
            else
            {
                // Clear steps if rank is deselected or blank Rank is chosen
                cboSalaryStep.ItemsSource = null;
                txtRankName.Text = "";
            }
            CalculateExpectedSalaryDate();
        }

        private void cboSalaryHistRankCode_SelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (cboSalaryHistRankCode.SelectedItem is Rank selectedRank)
            {
                LoadSalaryHistSteps(selectedRank.Code);
            }
            else
            {
                cboSalaryHistStep.ItemsSource = null;
            }
        }

        private void LoadSalaryHistSteps(string rankCode)
        {
            try
            {
                using (var context = new AppDbContext())
                {
                    var steps = context.RankSalarySpecs
                                       .Where(s => s.RankCode == rankCode)
                                       .ToList()
                                       .OrderBy(s => {
                                           if (string.IsNullOrEmpty(s.SalaryStep)) return 0;
                                           var parts = s.SalaryStep.Split('/');
                                           if (parts.Length > 0 && int.TryParse(parts[0], out int val))
                                               return val;
                                           return 999;
                                       })
                                       .ThenBy(s => {
                                           if (string.IsNullOrEmpty(s.SalaryStep)) return 0;
                                           var parts = s.SalaryStep.Split('/');
                                           if (parts.Length > 1 && int.TryParse(parts[1], out int val))
                                               return val;
                                           return 999;
                                       })
                                       .ToList();

                    cboSalaryHistStep.ItemsSource = steps;
                    cboSalaryHistStep.DisplayMemberPath = "SalaryStep";
                    cboSalaryHistStep.SelectedValuePath = "SalaryStep";
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading salary hist steps: {ex.Message}");
            }
        }

        private void LoadSalarySteps(string rankCode)
        {
            try
            {
                if (string.IsNullOrEmpty(rankCode))
                {
                    cboSalaryStep.ItemsSource = null;
                    return;
                }
                using (var context = new AppDbContext())
                {
                    var steps = context.RankSalarySpecs
                                       .Where(s => s.RankCode == rankCode)
                                       .ToList()
                                       .OrderBy(s => {
                                           if (string.IsNullOrEmpty(s.SalaryStep)) return 0;
                                           var parts = s.SalaryStep.Split('/');
                                           if (parts.Length > 0 && int.TryParse(parts[0], out int val))
                                               return val;
                                           return 999;
                                       })
                                       .ThenBy(s => {
                                           if (string.IsNullOrEmpty(s.SalaryStep)) return 0;
                                           var parts = s.SalaryStep.Split('/');
                                           if (parts.Length > 1 && int.TryParse(parts[1], out int val))
                                               return val;
                                           return 999;
                                       })
                                       .ToList();
                    steps.Insert(0, new RankSalarySpec { RankCode = rankCode, SalaryStep = "", Coefficient = 0 });

                    cboSalaryStep.ItemsSource = steps;
                    cboSalaryStep.DisplayMemberPath = "SalaryStep";
                    cboSalaryStep.SelectedValuePath = "SalaryStep";
                }
            }
            catch (Exception ex)
            {
                // Silently log or ignore in production, or show status bar message
                System.Diagnostics.Debug.WriteLine($"Error loading salary steps: {ex.Message}");
            }
        }

        private void cboSalaryStep_SelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (cboSalaryStep.SelectedItem is RankSalarySpec spec)
            {
                if (spec.Coefficient > 0)
                {
                    string coeffStr = spec.Coefficient.ToString("N2", System.Globalization.CultureInfo.GetCultureInfo("vi-VN"));
                    txtSalaryCoefficient.Text = coeffStr;
                }
                else
                {
                    txtSalaryCoefficient.Text = "";
                }
            }
            else
            {
                txtSalaryCoefficient.Text = "";
            }
        }

        private void cboSalaryHistStep_SelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (cboSalaryHistStep.SelectedItem is RankSalarySpec spec)
            {
                txtSalaryHistCoeff.Text = spec.Coefficient.ToString("N2", System.Globalization.CultureInfo.GetCultureInfo("vi-VN"));
            }
        }

        private void btnAddRank_Click(object? sender, RoutedEventArgs e)
        {
            var dialog = new AddRankDialog();
            if (dialog.ShowDialog() == true)
            {
                LoadRanks();
                if (!string.IsNullOrEmpty(dialog.SelectedRankCode))
                {
                    cboRankCode.SelectedValue = dialog.SelectedRankCode;
                    // Trigger name update manually if needed, but SelectionChanged should handle it if item exists
                }
            }
        }

        private void LoadDisciplineTypes()
        {
            using (var context = new AppDbContext())
            {
                var types = context.DisciplineTypes.Select(t => t.Name).ToList();
                types.Insert(0, "-- Không có --");
                cboDisciplineType.ItemsSource = types;

                // Default to "-- Không có --" if list is not empty
                if (types.Count > 0)
                {
                    cboDisciplineType.SelectedIndex = 0;
                }
            }
        }

        private void btnAddDisciplineType_Click(object? sender, RoutedEventArgs e)
        {
            var dialog = new DisciplineConfigDialog();
            dialog.ShowDialog(); // Just show dialog, refresh data afterwards regardless of return value
            LoadDisciplineTypes();
        }

        /// <summary>
        /// Xử lý sự kiện khi người dùng nhấn nút Lưu/Cập nhật.
        /// Thực hiện Validate dữ liệu và tiến hành Lưu mới hoặc Cập nhật bản ghi vào Cơ sở dữ liệu.
        /// </summary>
        private async void btnSave_Click(object? sender, RoutedEventArgs e)
        {
            // Kiểm tra tính hợp lệ của dữ liệu (Validation)
            if (string.IsNullOrWhiteSpace(txtName.Text))
            {
                MessageBox.Show("Vui lòng nhập tên!", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // Phone Validation
            string phoneDigits = new string(txtPhone.Text.Where(char.IsDigit).ToArray());
            if (!string.IsNullOrEmpty(txtPhone.Text) && (phoneDigits.Length != 10 || !phoneDigits.StartsWith("0")))
            {
                MessageBox.Show("Số điện thoại không hợp lệ! (Phải bắt đầu bằng số 0 và có 10 chữ số)", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // Email Validation
            if (!string.IsNullOrWhiteSpace(txtEmail.Text) && !IsValidEmail(txtEmail.Text))
            {
                MessageBox.Show("Địa chỉ Email không đúng định dạng!", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            using (var context = new AppDbContext())
            {
                if (_personnel == null || _personnel.Id == 0)
                {
                    // Add New
                    var newP = new Personnel
                    {
                        StaffId = txtStaffId.Text,
                        FullName = txtName.Text,
                        Gender = cboGender.Text,
                        DateOfBirth = dpDOB.SelectedDate,
                        PhoneNumber = new string(txtPhone.Text.Where(char.IsDigit).ToArray()),
                        IdentityCardNumber = txtIdentityNumber.Text,
                        IdentityCardPlace = txtIdentityPlace.Text,
                        Email = txtEmail.Text,
                        SocialSecurityNumber = txtSocialInsurance.Text,
                        BirthPlace = txtBirthPlace.Text,
                        Hometown = txtHometown.Text,
                        CurrentResidence = txtCurrentResidence.Text,
                        Ethnicity = txtEthnicity.Text,
                        Religion = txtReligion.Text,

                        Department = cboDepartment.Text,
                        Position = cboPosition.Text,
                        RankCode = (cboRankCode.SelectedItem as Rank)?.Code ?? "", // Allow free text if it's editable? Style is ComboBox, usually allows text if IsEditable=True. Default style might not.
                                                                                   // If DropDown only, use SelectedValue or Text (which returns selected item string)
                        RankName = txtRankName.Text,
                        TaxAuthorityStartDate = dpStartDate.SelectedDate,
                        StartDate = dpStartDate.SelectedDate,
                        Status = "Đang công tác",

                        EducationLevel = cboEducationLevel.Text,
                        Major = txtMajor.Text,
                        University = txtUniversity.Text,
                        StateManagementLevel = cboStateManagement.Text,
                        PoliticalTheoryLevel = cboPoliticalTheory.Text,
                        ITSkillLevel = txtITSkill.Text,
                        LanguageSkillLevel = txtLanguageSkill.Text,
                        AvatarBase64 = _currentAvatarBase64, // Save new avatar

                        // Tab 2
                        PositionDecisionDate = dpPositionDecisionDate.SelectedDate,
                        PositionCalculationDate = (dpPositionCalculationDate.SelectedDate.HasValue && dpPositionCalculationDate.SelectedDate.Value.Date == DateTime.Now.Date) ? null : dpPositionCalculationDate.SelectedDate,
                        PositionYear = txtPositionYear.Text,
                        DetailedWorkHistory = txtDetailedWorkHistory.Text,

                        // Tab 3
                        RetirementDate = dpRetirementDate.SelectedDate,

                        // Tab 4
                        PartyEntryDate = dpPartyEntryDate.SelectedDate,
                        PartyOfficialDate = dpPartyOfficialDate.SelectedDate,

                        // Tab 6
                        CurrentSalaryStep = cboSalaryStep.SelectedValue?.ToString(),
                        CurrentSalaryCoefficient = double.TryParse(txtSalaryCoefficient.Text, out double coeff) ? coeff : 0,
                        ExceedFramePercent = double.TryParse(txtExceedFrame.Text, out double exceed) ? exceed : 0,
                        PositionAllowance = txtPositionAllowance.Text,
                        SalaryReservationDeadline = dpSalaryReservation.SelectedDate,
                        NextSalaryStepDate = dpNextSalaryStepDate.SelectedDate,
                        SalaryIncreaseDelayType = cboSalaryDelay.Text,
                        ExpectedSalaryIncreaseDate = dpExpectedSalaryIncrease.SelectedDate,
                        SalaryHistoryLog = string.Join("\n", _personnel?.SalaryRecords?.Select(s => $"{s.StartDate?.ToString("dd/MM/yyyy")} - {s.Coefficient}") ?? Enumerable.Empty<string>()), // Legacy fallback
                        SalaryRecords = _personnel?.SalaryRecords,

                        // Tab 7
                        EmulationTitles = txtEmulationTitles.Text,
                        RewardForms = txtRewardForms.Text,

                        // Tab 8
                        DisciplineType = cboDisciplineType.Text,
                        DisciplineDecisionNumber = txtDisciplineNumber.Text,
                        DisciplineDecisionDate = dpDisciplineDate.SelectedDate,
                        DisciplineReason = txtDisciplineReason.Text,

                        // Tab 5 Leave
                        TotalAnnualLeaveDays = int.TryParse(txtTotalAnnualLeave.Text, out int leaves) ? leaves : 12,
                        LeaveCalculationDate = dpLeaveCalculationDate.SelectedDate,
                    };

                    // Copy Leave Histories if any (Handling the case where _personnel was initialized by btnAddLeave)
                    if (_personnel != null && _personnel.LeaveHistories != null)
                    {
                        foreach (var lh in _personnel.LeaveHistories)
                        {
                            newP.LeaveHistories.Add(new LeaveHistory
                            {
                                LeaveType = lh.LeaveType,
                                StartDate = lh.StartDate,
                                EndDate = lh.EndDate,
                                DurationDays = lh.DurationDays,
                                Reason = lh.Reason,
                                LeaveYear = lh.LeaveYear
                            });
                        }
                    }

                    // Copy Evaluation Records if any
                    if (_personnel != null && _personnel.EvaluationRecords != null)
                    {
                        newP.EvaluationRecords ??= new List<EvaluationRecord>();
                        foreach (var ev in _personnel.EvaluationRecords)
                        {
                            newP.EvaluationRecords.Add(new EvaluationRecord
                            {
                                Year = ev.Year,
                                Rating = ev.Rating,
                                DecisionNumber = ev.DecisionNumber,
                                DecisionDate = ev.DecisionDate,
                                DecisionAgency = ev.DecisionAgency
                            });
                        }
                    }

                    context.Personnel.Add(newP);
                    _personnel = newP;
                }
                else
                {
                    // Update
                    // Update
                    var existingP = context.Personnel.Include("LeaveHistories").Include("SalaryRecords").Include("EvaluationRecords").FirstOrDefault(p => p.Id == _personnel.Id);
                    if (existingP != null)
                    {
                        existingP.StaffId = txtStaffId.Text;
                        existingP.FullName = txtName.Text;
                        // ... (Other fields remain same, not restating all for brevity in thought, but implementation needs them)
                        // Actually replace_file_content needs EXACT TargetContent.
                        // I will target the existingP retrieval and the end of the block.

                        existingP.Gender = cboGender.Text;
                        existingP.DateOfBirth = dpDOB.SelectedDate;
                        existingP.PhoneNumber = new string(txtPhone.Text.Where(char.IsDigit).ToArray());
                        existingP.IdentityCardNumber = txtIdentityNumber.Text;
                        existingP.IdentityCardPlace = txtIdentityPlace.Text;
                        existingP.Email = txtEmail.Text;
                        existingP.SocialSecurityNumber = txtSocialInsurance.Text;
                        existingP.BirthPlace = txtBirthPlace.Text;
                        existingP.Hometown = txtHometown.Text;
                        existingP.CurrentResidence = txtCurrentResidence.Text;
                        existingP.Ethnicity = txtEthnicity.Text;
                        existingP.Religion = txtReligion.Text;

                        existingP.Department = cboDepartment.Text;
                        existingP.Position = cboPosition.Text;
                        existingP.RankCode = cboRankCode.Text;
                        existingP.RankName = txtRankName.Text;
                        existingP.TaxAuthorityStartDate = dpStartDate.SelectedDate;

                        existingP.EducationLevel = cboEducationLevel.Text;
                        existingP.Major = txtMajor.Text;
                        existingP.University = txtUniversity.Text;
                        existingP.StateManagementLevel = cboStateManagement.Text;
                        existingP.PoliticalTheoryLevel = cboPoliticalTheory.Text;
                        existingP.ITSkillLevel = txtITSkill.Text;
                        existingP.LanguageSkillLevel = txtLanguageSkill.Text;

                        if (_isAvatarChanged)
                        {
                            existingP.AvatarBase64 = _currentAvatarBase64;
                        }

                        // Tab 2
                        existingP.PositionDecisionDate = dpPositionDecisionDate.SelectedDate;
                        existingP.PositionCalculationDate = (dpPositionCalculationDate.SelectedDate.HasValue && dpPositionCalculationDate.SelectedDate.Value.Date == DateTime.Now.Date) ? null : dpPositionCalculationDate.SelectedDate;
                        existingP.PositionYear = txtPositionYear.Text;
                        existingP.DetailedWorkHistory = txtDetailedWorkHistory.Text;

                        // Tab 3
                        existingP.RetirementDate = dpRetirementDate.SelectedDate;

                        // Tab 4
                        existingP.PartyEntryDate = dpPartyEntryDate.SelectedDate;
                        existingP.PartyOfficialDate = dpPartyOfficialDate.SelectedDate;

                        // Tab 6
                        existingP.CurrentSalaryStep = cboSalaryStep.Text;
                        existingP.CurrentSalaryCoefficient = double.TryParse(txtSalaryCoefficient.Text, out double sc2) ? sc2 : 0;
                        existingP.ExceedFramePercent = double.TryParse(txtExceedFrame.Text, out double ef2) ? ef2 : 0;
                        existingP.PositionAllowance = txtPositionAllowance.Text;
                        existingP.SalaryReservationDeadline = dpSalaryReservation.SelectedDate;
                        existingP.NextSalaryStepDate = dpNextSalaryStepDate.SelectedDate;
                        existingP.SalaryIncreaseDelayType = cboSalaryDelay.Text;
                        existingP.ExpectedSalaryIncreaseDate = dpExpectedSalaryIncrease.SelectedDate;

                        // Handle Salary Records Sync
                        if (existingP.SalaryRecords == null) existingP.SalaryRecords = new List<SalaryRecord>();
                        if (_personnel.SalaryRecords != null)
                        {
                            // 1. Identify IDs to keep
                            var currentSalaryIds = _personnel.SalaryRecords.Select(s => s.Id).Where(id => id != 0).ToList();

                            // 2. Remove deleted
                            var toDeleteSalary = existingP.SalaryRecords.Where(s => !currentSalaryIds.Contains(s.Id)).ToList();
                            foreach (var item in toDeleteSalary)
                            {
                                context.SalaryRecords.Remove(item);
                            }

                            // 3. Add new
                            var toAddSalary = _personnel.SalaryRecords.Where(s => s.Id == 0).ToList();
                            foreach (var item in toAddSalary)
                            {
                                item.PersonnelId = existingP.Id;
                                context.SalaryRecords.Add(item);
                            }

                            // 4. Update existing
                            var toUpdateSalary = _personnel.SalaryRecords.Where(s => s.Id != 0).ToList();
                            foreach (var item in toUpdateSalary)
                            {
                                var dbItem = existingP.SalaryRecords.FirstOrDefault(x => x.Id == item.Id);
                                if (dbItem != null)
                                {
                                    dbItem.StartDate = item.StartDate;
                                    dbItem.EndDate = item.EndDate;
                                    dbItem.SalaryCalculationDate = item.SalaryCalculationDate;
                                    dbItem.Coefficient = item.Coefficient;
                                    dbItem.Percentage = item.Percentage;
                                    dbItem.DecisionNumber = item.DecisionNumber;
                                    dbItem.DecisionDate = item.DecisionDate;
                                }
                            }
                        }

                        existingP.SalaryHistoryLog = string.Join("\n", existingP.SalaryRecords?.Select(s => $"{s.StartDate?.ToString("dd/MM/yyyy")} - {s.Coefficient}") ?? new List<string>());

                        // Tab 7
                        existingP.EmulationTitles = txtEmulationTitles.Text;
                        existingP.RewardForms = txtRewardForms.Text;

                        // Tab 8
                        existingP.DisciplineType = cboDisciplineType.Text;
                        existingP.DisciplineDecisionNumber = txtDisciplineNumber.Text;
                        existingP.DisciplineDecisionDate = dpDisciplineDate.SelectedDate;
                        existingP.DisciplineReason = txtDisciplineReason.Text;

                        // Salary Config Update Logic
                        // Check if salary step changed to update history automatically?
                        // Tab 5
                        if (int.TryParse(txtTotalAnnualLeave.Text, out int totalLeave))
                        {
                            existingP.TotalAnnualLeaveDays = totalLeave;
                        }
                        existingP.LeaveCalculationDate = dpLeaveCalculationDate.SelectedDate;

                        // Sync Leave History
                        if (_personnel.LeaveHistories != null)
                        {
                            // 1. Identify IDs to keep
                            var currentIds = _personnel.LeaveHistories.Select(L => L.Id).Where(id => id != 0).ToList();

                            // 2. Remove deleted
                            var toDelete = existingP.LeaveHistories.Where(lh => !currentIds.Contains(lh.Id)).ToList();
                            foreach (var item in toDelete)
                            {
                                context.LeaveHistories.Remove(item);
                            }

                            // 3. Add new
                            var toAdd = _personnel.LeaveHistories.Where(lh => lh.Id == 0).ToList();
                            foreach (var item in toAdd)
                            {
                                item.PersonnelId = existingP.Id; // Link ForeignKey
                                context.LeaveHistories.Add(item);
                            }

                            // 4. Update existing
                            var toUpdate = _personnel.LeaveHistories.Where(lh => lh.Id != 0).ToList();
                            foreach (var item in toUpdate)
                            {
                                var dbItem = existingP.LeaveHistories.FirstOrDefault(x => x.Id == item.Id);
                                if (dbItem != null)
                                {
                                    dbItem.LeaveType = item.LeaveType;
                                    dbItem.StartDate = item.StartDate;
                                    dbItem.EndDate = item.EndDate;
                                    dbItem.DurationDays = item.DurationDays;
                                    dbItem.Reason = item.Reason;
                                    dbItem.LeaveYear = item.LeaveYear;
                                }
                            }
                        }

                        // Sync Evaluation Records
                        if (existingP.EvaluationRecords == null) existingP.EvaluationRecords = new List<EvaluationRecord>();
                        if (_personnel.EvaluationRecords != null)
                        {
                            // 1. Identify IDs to keep
                            var currentEvalIds = _personnel.EvaluationRecords.Select(e => e.Id).Where(id => id != 0).ToList();

                            // 2. Remove deleted
                            var toDeleteEval = existingP.EvaluationRecords.Where(e => !currentEvalIds.Contains(e.Id)).ToList();
                            foreach (var item in toDeleteEval)
                            {
                                context.EvaluationRecords.Remove(item);
                            }

                            // 3. Add new
                            var toAddEval = _personnel.EvaluationRecords.Where(e => e.Id == 0).ToList();
                            foreach (var item in toAddEval)
                            {
                                item.PersonnelId = existingP.Id;
                                context.EvaluationRecords.Add(item);
                            }

                            // 4. Update existing
                            var toUpdateEval = _personnel.EvaluationRecords.Where(e => e.Id != 0).ToList();
                            foreach (var item in toUpdateEval)
                            {
                                var dbItem = existingP.EvaluationRecords.FirstOrDefault(x => x.Id == item.Id);
                                if (dbItem != null)
                                {
                                    dbItem.Year = item.Year;
                                    dbItem.Rating = item.Rating;
                                    dbItem.DecisionNumber = item.DecisionNumber;
                                    dbItem.DecisionDate = item.DecisionDate;
                                    dbItem.DecisionAgency = item.DecisionAgency;
                                }
                            }
                        }
                    }
                }
                try
                {
                    context.SaveChanges();
                }
                catch (Exception ex)
                {
                    string errorMsg = ex.Message;
                    if (ex.InnerException != null)
                    {
                        errorMsg += "\nInner Error: " + ex.InnerException.Message;
                    }
                    MessageBox.Show("Lỗi khi lưu dữ liệu:\n" + errorMsg, "Lỗi Database", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                // Show Custom Success Window
                var successWindow = new SuccessWindow();
                successWindow.Owner = Application.Current.MainWindow; // Set owner to center over main window
                successWindow.ShowDialog();
            }

            if (Application.Current.MainWindow is MainWindow mw)
            {
                int? targetId = _personnel?.Id;
                mw.ClearPersonnelCache();
                mw.NavigateToDashboard(targetId);
            }
        }

        private void btnAvatar_Click(object? sender, RoutedEventArgs e)
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Image files (*.png;*.jpg)|*.png;*.jpg"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    // 1. Display (Orient image using EXIF orientation metadata)
                    var bitmap = TaxPersonnelManagement.Helpers.ImageHelper.LoadAndOrientImage(openFileDialog.FileName);
                    var brush = new System.Windows.Media.ImageBrush(bitmap)
                    {
                        Stretch = System.Windows.Media.Stretch.UniformToFill
                    };
                    imgAvatar.Fill = brush;
                    iconAvatarPlaceholder.Visibility = Visibility.Collapsed;
                    btnRemoveAvatar.Visibility = Visibility.Visible;

                    // 2. Convert to Base64 (Resize if needed)
                    _currentAvatarBase64 = ImageToBase64(openFileDialog.FileName);
                    _isAvatarChanged = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Lỗi khi tải ảnh: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void btnRemoveAvatar_Click(object? sender, RoutedEventArgs e)
        {
            imgAvatar.Fill = new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#F5F5F5"));
            iconAvatarPlaceholder.Visibility = Visibility.Visible;
            btnRemoveAvatar.Visibility = Visibility.Collapsed;

            _currentAvatarBase64 = null;
            _isAvatarChanged = true;
        }

        // Helper: File -> Base64 (with resizing to max 500px width/height to save space, and correcting EXIF orientation)
        private string ImageToBase64(string path)
        {
            try
            {
                // Load and orient the image first
                var image = TaxPersonnelManagement.Helpers.ImageHelper.LoadAndOrientImage(path);

                double scale = 1.0;
                double maxDimension = 600;
                if (image.PixelWidth > maxDimension || image.PixelHeight > maxDimension)
                {
                    scale = Math.Min(maxDimension / image.PixelWidth, maxDimension / image.PixelHeight);
                }

                var transformed = new System.Windows.Media.Imaging.TransformedBitmap(image, new System.Windows.Media.ScaleTransform(scale, scale));

                var encoder = new System.Windows.Media.Imaging.JpegBitmapEncoder(); // JPG is smaller than PNG usually
                encoder.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(transformed));
                encoder.QualityLevel = 80;

                using (var ms = new System.IO.MemoryStream())
                {
                    encoder.Save(ms);
                    return Convert.ToBase64String(ms.ToArray());
                }
            }
            catch
            {
                return "";
            }
        }

        // Helper: Base64 -> BitmapSource (Orient image using EXIF orientation metadata)
        private System.Windows.Media.Imaging.BitmapSource? Base64ToImage(string base64String)
        {
            try
            {
                byte[] binaryData = Convert.FromBase64String(base64String);
                return TaxPersonnelManagement.Helpers.ImageHelper.LoadAndOrientImage(binaryData);
            }
            catch
            {
                return null;
            }
        }

        private void btnRefresh_Click(object? sender, RoutedEventArgs e)
        {
            _isRefreshing = true;
            try
            {
                // Detach events to prevent loops
                cboLeaveType.SelectionChanged -= OnLeaveDateChanged;
                dpLeaveStartDate.SelectedDateChanged -= OnLeaveDateChanged;
                dpLeaveEndDate.SelectedDateChanged -= OnLeaveDateChanged;
                dpStartDate.SelectedDateChanged -= OnRetirementDateChanged;

                // Reset fields
                _personnel = null;
                lblSaveText.Text = "Lưu Hồ Sơ";

                txtStaffId.Clear();
                txtName.Clear();
                cboGender.SelectedItem = null;
                cboGender.Text = "";
                dpDOB.SelectedDate = null;
                txtPhone.Clear();
                txtIdentityNumber.Clear();
                txtIdentityPlace.Text = "Cục Cảnh sát quản lý hành chính về trật tự xã hội";
                txtEmail.Clear();
                txtSocialInsurance.Clear();
                txtBirthPlace.Clear();
                txtHometown.Clear();
                txtCurrentResidence.Clear();
                txtEthnicity.Clear();
                txtReligion.Clear();

                cboDepartment.SelectedItem = null;
                cboDepartment.Text = "";
                cboPosition.SelectedItem = null;
                cboPosition.Text = "";

                cboRankCode.SelectedItem = null;
                cboRankCode.Text = "";
                txtRankName.Clear();
                dpStartDate.SelectedDate = null;

                cboEducationLevel.SelectedItem = null;
                cboEducationLevel.Text = "";
                txtMajor.Clear();
                txtUniversity.Clear();
                cboStateManagement.SelectedItem = null;
                cboStateManagement.Text = "";
                cboPoliticalTheory.SelectedItem = null;
                cboPoliticalTheory.Text = "";
                txtITSkill.Clear();
                txtLanguageSkill.Clear();

                // Tab 2
                dpPositionDecisionDate.SelectedDate = null;
                dpPositionCalculationDate.SelectedDate = DateTime.Now;
                txtPositionYear.Clear();
                txtDetailedWorkHistory.Clear();
                txtYearsWorked.Clear();
                txtYearsWorked.Clear();
                txtMonthsWorked.Clear();

                // Tab 3
                dpRetirementDate.SelectedDate = null;
                txtRetirementYearsWorked.Clear();
                txtRemainingYears.Clear();

                // Tab 4
                dpPartyEntryDate.SelectedDate = null;
                dpPartyOfficialDate.SelectedDate = null;
                txtPartyAge.Clear();

                // Tab 6: Salary Info
                cboSalaryStep.ItemsSource = null;
                cboSalaryStep.SelectedIndex = -1;
                cboSalaryStep.Text = "";
                txtSalaryCoefficient.Clear();
                txtExceedFrame.Clear();
                txtPositionAllowance.Clear();
                dpSalaryReservation.SelectedDate = null;
                dpNextSalaryStepDate.SelectedDate = null;
                cboSalaryDelay.SelectedIndex = -1;
                cboSalaryDelay.Text = "";
                dpExpectedSalaryIncrease.SelectedDate = null;
                // Tab 6
                if (_personnel == null) _personnel = new Personnel();
                if (_personnel.SalaryRecords == null) _personnel.SalaryRecords = new System.Collections.Generic.List<SalaryRecord>();
                _personnel.SalaryRecords.Clear();
                RefreshSalaryHistoryGrid();

                dpSalaryHistStart.SelectedDate = null;
                dpSalaryHistEnd.SelectedDate = null;
                dpSalaryHistCalc.SelectedDate = null;
                txtSalaryHistCoeff.Clear();
                txtSalaryHistPercent.Text = "100";
                txtSalaryHistDecisionNo.Clear();
                dpSalaryHistDecisionDate.SelectedDate = null;

                // Tab 5
                // Tab 5
                txtTotalAnnualLeave.Text = "12";
                txtAnnualLeaveTaken.Text = "0";
                txtAnnualLeaveRemaining.Text = "12";
                txtSickLeaveTaken.Text = "0";
                txtUnpaidLeaveTaken.Text = "0";
                txtMaternityLeaveTaken.Text = "0";
                dpLeaveCalculationDate.SelectedDate = null;

                dpLeaveStartDate.SelectedDate = null;
                dpLeaveEndDate.SelectedDate = null;
                cboLeaveType.SelectedIndex = -1;
                txtLeaveReason.Clear();
                txtLeaveDuration.Clear();
                dgLeaveHistory.ItemsSource = null;

                // Reset Edit Mode
                _editingLeaveHistory = null;
                if (btnAddLeave.Content is StackPanel sp)
                {
                    foreach (var child in sp.Children)
                    {
                        if (child is TextBlock tb)
                        {
                            tb.Text = "Thêm vào bảng";
                            break;
                        }
                    }
                }

                // Tab 9 Resets
                cboEvaluationYear.SelectedIndex = 0;
                cboEvaluationRating.SelectedIndex = 0;
                txtEvaluationDecisionNo.Clear();
                dpEvaluationDecisionDate.SelectedDate = null;
                txtEvaluationDecisionAgency.Clear();
                _editingEvaluationRecord = null;
                dgEvaluationHistory.ItemsSource = null;

                // Clear Avatar
                btnRemoveAvatar_Click(sender, e);
                _isAvatarChanged = false; // Reset change flag as we are starting fresh

                // Clear Errors
                lblPhoneError.Visibility = Visibility.Collapsed;
                lblEmailError.Visibility = Visibility.Collapsed;

                // Focus first field
                txtStaffId.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi làm mới: " + ex.Message);
            }
            finally
            {
                // Reattach events
                cboLeaveType.SelectionChanged += OnLeaveDateChanged;
                dpLeaveStartDate.SelectedDateChanged += OnLeaveDateChanged;
                dpLeaveEndDate.SelectedDateChanged += OnLeaveDateChanged;
                dpStartDate.SelectedDateChanged += OnRetirementDateChanged;
                _isRefreshing = false;
            }
        }


        private void txtEmail_TextChanged(object? sender, TextChangedEventArgs e)
        {
            if (string.IsNullOrEmpty(txtEmail.Text))
            {
                lblEmailError.Visibility = Visibility.Collapsed;
                return;
            }

            if (IsValidEmail(txtEmail.Text))
            {
                lblEmailError.Visibility = Visibility.Collapsed;
            }
            else
            {
                // To avoid annoying on first char, maybe only show if it LOOKS like they tried to finish?
                // But user asked for "nhập không đúng định dạng báo luôn". 
                // Let's show it always when invalid.
                lblEmailError.Visibility = Visibility.Visible;
            }
        }

        private void dpExpectedSalaryIncrease_DateValidationError(object? sender, DatePickerDateValidationErrorEventArgs e)
        {
            var warning = new WarningWindow("Ngày vừa nhập không hợp lệ. Vui lòng nhập ngày theo định dạng dd/MM/yyyy (VD: 15/05/2026).", "Lỗi định dạng");
            warning.Owner = Window.GetWindow(this);
            warning.ShowDialog();
            e.ThrowException = false;
        }
        private void OnPositionDateChanged(object? sender, SelectionChangedEventArgs e)
        {
            CalculateWorkDuration();
        }

        private void CalculateWorkDuration()
        {
            // Requirement Update:
            // Start Date = PositionDecisionDate ("Thời gian công tác tính theo QĐ gần nhất")
            // End Date = PositionCalculationDate or DateTime.Now
            if (dpPositionDecisionDate.SelectedDate.HasValue)
            {
                DateTime startDate = dpPositionDecisionDate.SelectedDate.Value;
                DateTime endDate = dpPositionCalculationDate.SelectedDate ?? DateTime.Now;

                if (endDate < startDate)
                {
                    txtYearsWorked.Text = "0";
                    txtMonthsWorked.Text = "0";
                    txtPositionYear.Text = "0 năm 0 tháng";
                    return;
                }

                // Years: Full years
                int years = endDate.Year - startDate.Year;
                if (startDate.Date > endDate.AddYears(-years)) years--;

                // Months: Remainder months
                DateTime tmpDate = startDate.AddYears(years);
                int months = 0;
                while (tmpDate.AddMonths(1) <= endDate)
                {
                    months++;
                    tmpDate = tmpDate.AddMonths(1);
                }

                txtYearsWorked.Text = years.ToString();
                txtMonthsWorked.Text = months.ToString();
                txtPositionYear.Text = $"{years} năm {months} tháng";
            }
            else
            {
                txtYearsWorked.Clear();
                txtMonthsWorked.Clear();
                txtPositionYear.Clear();
            }
        }

        private void OnRetirementDateChanged(object? sender, SelectionChangedEventArgs e)
        {
            CalculateRetirementInfo();
        }

        private void dpPartyEntryDate_SelectedDateChanged(object? sender, SelectionChangedEventArgs e)
        {
            CalculatePartyAge();
        }

        private void CalculatePartyAge()
        {
            if (txtPartyAge == null) return;
            if (dpPartyEntryDate.SelectedDate.HasValue)
            {
                DateTime startDate = dpPartyEntryDate.SelectedDate.Value;
                DateTime endDate = DateTime.Now;

                if (endDate < startDate)
                {
                    txtPartyAge.Text = "0 năm 0 tháng";
                    return;
                }

                int years = endDate.Year - startDate.Year;
                if (startDate.Date > endDate.AddYears(-years)) years--;

                DateTime tmpDate = startDate.AddYears(years);
                int months = 0;
                while (tmpDate.AddMonths(1) <= endDate)
                {
                    months++;
                    tmpDate = tmpDate.AddMonths(1);
                }

                txtPartyAge.Text = $"{years} năm {months} tháng";
            }
            else
            {
                txtPartyAge.Clear();
            }
        }

        private void CalculateRetirementInfo()
        {
            // Requirement Update:
            // 1. Số năm công tác = Now - StartDate (TaxAuthorityStartDate) -> Output: X năm Y tháng Z ngày
            // 2. Số năm còn công tác = RetirementDate - Now -> Output: X năm Y tháng Z ngày

            DateTime now = DateTime.Now.Date;

            // Calculate Annual Leave (Tab 5)
            CalculateAnnualLeave(now);


            // 1. Calculate Years Worked (Tax Authority Start Date -> Now)
            if (dpStartDate.SelectedDate.HasValue)
            {
                DateTime start = dpStartDate.SelectedDate.Value;

                if (now > start)
                {
                    txtRetirementYearsWorked.Text = CalculateDetailedDateDifference(start, now);
                }
                else
                {
                    txtRetirementYearsWorked.Text = "0 năm 0 tháng 0 ngày";
                }
            }
            else
            {
                txtRetirementYearsWorked.Clear();
            }

            // 2. Calculate Remaining Years (Now -> Retirement Date)
            // Requirement: Add 1 day to the difference (Inclusive)
            if (dpRetirementDate.SelectedDate.HasValue)
            {
                DateTime retirement = dpRetirementDate.SelectedDate.Value;

                if (retirement > now)
                {
                    txtRemainingYears.Text = CalculateDetailedDateDifference(now, retirement);
                }
                else
                {
                    txtRemainingYears.Text = "0 năm 0 tháng 0 ngày"; // Already retired
                }
            }
            else
            {
                txtRemainingYears.Clear();
            }
        }

        private void OnLeaveCalculationDateChanged(object? sender, SelectionChangedEventArgs e)
        {
            CalculateAnnualLeave(DateTime.Now);
        }

        private void CalculateAnnualLeave(DateTime now)
        {
            // Rule: Total Annual Leave = (Years Worked / 5) + 12
            int totalLeave = 12;

            DateTime? targetStartDate = dpLeaveCalculationDate.SelectedDate ?? dpStartDate.SelectedDate;

            if (targetStartDate.HasValue)
            {
                DateTime start = targetStartDate.Value;

                // New formula: Số năm tính ngày phép = Năm hiện tại - Năm tính phép + 1
                int years = now.Year - start.Year + 1;

                if (years < 0) years = 0;

                int bonusDays = years / 5;
                totalLeave = 12 + bonusDays;
            }

            // Update Model and UI
            if (_personnel != null) _personnel.TotalAnnualLeaveDays = totalLeave;
            txtTotalAnnualLeave.Text = totalLeave.ToString();
            UpdateLeaveStatistics(); // Recalculate and display statistics for the current selected year
        }

        private string CalculateDetailedDateDifference(DateTime startDate, DateTime endDate)
        {
            if (startDate > endDate) return "0 năm 0 tháng 0 ngày";
            DateTime tempDate = startDate;
            int years = 0, months = 0, days = 0;
            while (tempDate.AddYears(1) <= endDate) { years++; tempDate = tempDate.AddYears(1); }
            while (tempDate.AddMonths(1) <= endDate) { months++; tempDate = tempDate.AddMonths(1); }
            days = (endDate - tempDate).Days;
            return $"{years} năm {months} tháng {days} ngày";
        }

        private void RefreshSalaryHistoryGrid()
        {
            if (_personnel == null) return;
            if (_personnel.SalaryRecords == null) _personnel.SalaryRecords = new System.Collections.Generic.List<SalaryRecord>();
            dgSalaryHistory.ItemsSource = null;
            dgSalaryHistory.ItemsSource = _personnel.SalaryRecords.OrderByDescending(s => s.StartDate).ToList();
        }

        private void ReconcileSalaryRecordsAfterAdd(System.Collections.Generic.ICollection<SalaryRecord> records, SalaryRecord newItem)
        {
            if (newItem.StartDate == null) return;
            DateTime newStart = newItem.StartDate.Value;

            var otherRecords = records
                .Where(r => r != newItem && r.StartDate.HasValue)
                .OrderBy(r => r.StartDate!.Value)
                .ToList();

            // Find predecessor and successor
            SalaryRecord? predecessor = otherRecords.LastOrDefault(r => r.StartDate!.Value < newStart);
            SalaryRecord? successor = otherRecords.FirstOrDefault(r => r.StartDate!.Value > newStart);

            // If newItem's EndDate is null and successor exists, set it
            if (newItem.EndDate == null && successor != null && successor.StartDate.HasValue)
            {
                newItem.EndDate = successor.StartDate!.Value.AddDays(-1);
            }

            // If predecessor exists
            if (predecessor != null)
            {
                if (predecessor.EndDate == null)
                {
                    predecessor.EndDate = newStart.AddDays(-1);
                }
                else if (successor != null && successor.StartDate.HasValue && predecessor.EndDate == successor.StartDate!.Value.AddDays(-1))
                {
                    predecessor.EndDate = newStart.AddDays(-1);
                }
            }
        }

        private void ReconcileSalaryRecordsAfterDelete(System.Collections.Generic.ICollection<SalaryRecord> records, DateTime? deletedStartDate)
        {
            if (deletedStartDate == null) return;
            DateTime delStart = deletedStartDate.Value;

            var remainingRecords = records
                .Where(r => r.StartDate.HasValue)
                .OrderBy(r => r.StartDate!.Value)
                .ToList();

            SalaryRecord? predecessor = remainingRecords.LastOrDefault(r => r.StartDate!.Value < delStart);
            SalaryRecord? successor = remainingRecords.FirstOrDefault(r => r.StartDate!.Value > delStart);

            if (predecessor != null && predecessor.EndDate == delStart.AddDays(-1))
            {
                if (successor != null && successor.StartDate.HasValue)
                {
                    predecessor.EndDate = successor.StartDate!.Value.AddDays(-1);
                }
                else
                {
                    predecessor.EndDate = null;
                }
            }
        }

        private void ReconcileSalaryRecordsAfterUpdate(System.Collections.Generic.ICollection<SalaryRecord> records, SalaryRecord editedItem, DateTime? oldStartDate)
        {
            if (editedItem.StartDate == oldStartDate)
            {
                // Just check if we need to auto-calculate editedItem's own EndDate because it was set to null
                if (editedItem.EndDate == null && editedItem.StartDate.HasValue)
                {
                    var successor = records
                        .Where(r => r != editedItem && r.StartDate.HasValue && r.StartDate!.Value > editedItem.StartDate!.Value)
                        .OrderBy(r => r.StartDate!.Value)
                        .FirstOrDefault();
                    if (successor != null && successor.StartDate.HasValue)
                    {
                        editedItem.EndDate = successor.StartDate!.Value.AddDays(-1);
                    }
                }
                return;
            }

            // If StartDate changed:
            // 1. Temporarily remove editedItem from list to simulate delete of oldStartDate
            records.Remove(editedItem);

            // 2. Run delete reconciliation for oldStartDate
            ReconcileSalaryRecordsAfterDelete(records, oldStartDate);

            // 3. Put editedItem back
            records.Add(editedItem);

            // 4. Run add reconciliation for newStartDate
            ReconcileSalaryRecordsAfterAdd(records, editedItem);
        }

        private bool ValidateSalaryRecordsOverlap(System.Collections.Generic.ICollection<SalaryRecord> records)
        {
            var sorted = records
                .Where(r => r.StartDate.HasValue)
                .OrderBy(r => r.StartDate!.Value)
                .ToList();

            for (int i = 0; i < sorted.Count - 1; i++)
            {
                var current = sorted[i];
                var next = sorted[i + 1];

                if (current.StartDate!.Value == next.StartDate!.Value)
                {
                    return false; // Trùng ngày bắt đầu
                }

                if (current.EndDate.HasValue && current.EndDate!.Value >= next.StartDate!.Value)
                {
                    return false; // Trùng lặp/chồng chéo khoảng thời gian
                }
            }

            return true;
        }

        private void btnAddSalaryHistory_Click(object sender, RoutedEventArgs e)
        {
            if (dpSalaryHistStart.SelectedDate == null)
            {
                new WarningWindow("Vui lòng nhập ngày bắt đầu!", "Thiếu thông tin").ShowDialog();
                return;
            }

            if (_personnel == null) _personnel = new Personnel();
            if (_personnel.SalaryRecords == null) _personnel.SalaryRecords = new System.Collections.Generic.List<SalaryRecord>();

            // 1. Create a clone list for dry-run
            var tempRecords = new System.Collections.Generic.List<SalaryRecord>();
            foreach (var r in _personnel.SalaryRecords)
            {
                tempRecords.Add(new SalaryRecord
                {
                    Id = r.Id,
                    PersonnelId = r.PersonnelId,
                    StartDate = r.StartDate,
                    EndDate = r.EndDate,
                    SalaryCalculationDate = r.SalaryCalculationDate,
                    RankCode = r.RankCode,
                    SalaryStep = r.SalaryStep,
                    Coefficient = r.Coefficient,
                    Percentage = r.Percentage,
                    ExceedFramePercent = r.ExceedFramePercent,
                    DecisionNumber = r.DecisionNumber,
                    DecisionDate = r.DecisionDate,
                    Note = r.Note
                });
            }

            if (_editingSalaryRecord != null)
            {
                // Find matching record in tempRecords
                var tempEdit = tempRecords.FirstOrDefault(r => r.Id == _editingSalaryRecord.Id && r.StartDate == _editingSalaryRecord.StartDate && r.RankCode == _editingSalaryRecord.RankCode && r.SalaryStep == _editingSalaryRecord.SalaryStep);
                if (tempEdit == null)
                {
                    // Fallback to match by reference index if Id is 0
                    int idx = _personnel.SalaryRecords.ToList().IndexOf(_editingSalaryRecord);
                    if (idx >= 0 && idx < tempRecords.Count)
                    {
                        tempEdit = tempRecords[idx];
                    }
                }

                if (tempEdit != null)
                {
                    DateTime? oldStartDate = tempEdit.StartDate;
                    tempEdit.StartDate = dpSalaryHistStart.SelectedDate;
                    tempEdit.EndDate = dpSalaryHistEnd.SelectedDate;
                    tempEdit.SalaryCalculationDate = dpSalaryHistCalc.SelectedDate;
                    tempEdit.RankCode = cboSalaryHistRankCode.SelectedValue?.ToString() ?? "";
                    tempEdit.SalaryStep = cboSalaryHistStep.SelectedValue?.ToString() ?? "";
                    tempEdit.Coefficient = txtSalaryHistCoeff.Text;
                    tempEdit.Percentage = double.TryParse(txtSalaryHistPercent.Text, out double p) ? p : 100;
                    tempEdit.ExceedFramePercent = double.TryParse(txtSalaryHistExceedPercent.Text, out double ef) ? ef : 0;
                    tempEdit.DecisionNumber = txtSalaryHistDecisionNo.Text;
                    tempEdit.DecisionDate = dpSalaryHistDecisionDate.SelectedDate;

                    ReconcileSalaryRecordsAfterUpdate(tempRecords, tempEdit, oldStartDate);
                }
            }
            else
            {
                var tempNew = new SalaryRecord
                {
                    StartDate = dpSalaryHistStart.SelectedDate,
                    EndDate = dpSalaryHistEnd.SelectedDate,
                    SalaryCalculationDate = dpSalaryHistCalc.SelectedDate,
                    RankCode = cboSalaryHistRankCode.SelectedValue?.ToString() ?? "",
                    SalaryStep = cboSalaryHistStep.SelectedValue?.ToString() ?? "",
                    Coefficient = txtSalaryHistCoeff.Text,
                    Percentage = double.TryParse(txtSalaryHistPercent.Text, out double p) ? p : 100,
                    ExceedFramePercent = double.TryParse(txtSalaryHistExceedPercent.Text, out double ef) ? ef : 0,
                    DecisionNumber = txtSalaryHistDecisionNo.Text,
                    DecisionDate = dpSalaryHistDecisionDate.SelectedDate,
                    PersonnelId = _personnel.Id
                };
                tempRecords.Add(tempNew);
                ReconcileSalaryRecordsAfterAdd(tempRecords, tempNew);
            }

            // 2. Validate overlaps in tempRecords
            if (!ValidateSalaryRecordsOverlap(tempRecords))
            {
                new WarningWindow("Khoảng thời gian quá trình lương bị trùng hoặc ngày bắt đầu đã tồn tại!", "Trùng lặp thời gian").ShowDialog();
                return;
            }

            // 3. If valid, apply to actual _personnel.SalaryRecords
            if (_editingSalaryRecord != null)
            {
                DateTime? oldStartDate = _editingSalaryRecord.StartDate;
                _editingSalaryRecord.StartDate = dpSalaryHistStart.SelectedDate;
                _editingSalaryRecord.EndDate = dpSalaryHistEnd.SelectedDate;
                _editingSalaryRecord.SalaryCalculationDate = dpSalaryHistCalc.SelectedDate;
                _editingSalaryRecord.RankCode = cboSalaryHistRankCode.SelectedValue?.ToString() ?? "";
                _editingSalaryRecord.SalaryStep = cboSalaryHistStep.SelectedValue?.ToString() ?? "";
                _editingSalaryRecord.Coefficient = txtSalaryHistCoeff.Text;
                _editingSalaryRecord.Percentage = double.TryParse(txtSalaryHistPercent.Text, out double p) ? p : 100;
                _editingSalaryRecord.ExceedFramePercent = double.TryParse(txtSalaryHistExceedPercent.Text, out double ef) ? ef : 0;
                _editingSalaryRecord.DecisionNumber = txtSalaryHistDecisionNo.Text;
                _editingSalaryRecord.DecisionDate = dpSalaryHistDecisionDate.SelectedDate;

                ReconcileSalaryRecordsAfterUpdate(_personnel.SalaryRecords, _editingSalaryRecord, oldStartDate);
                _editingSalaryRecord = null; // Clear editing state
            }
            else
            {
                var newItem = new SalaryRecord
                {
                    StartDate = dpSalaryHistStart.SelectedDate,
                    EndDate = dpSalaryHistEnd.SelectedDate,
                    SalaryCalculationDate = dpSalaryHistCalc.SelectedDate,
                    RankCode = cboSalaryHistRankCode.SelectedValue?.ToString() ?? "",
                    SalaryStep = cboSalaryHistStep.SelectedValue?.ToString() ?? "",
                    Coefficient = txtSalaryHistCoeff.Text,
                    Percentage = double.TryParse(txtSalaryHistPercent.Text, out double p) ? p : 100,
                    ExceedFramePercent = double.TryParse(txtSalaryHistExceedPercent.Text, out double ef) ? ef : 0,
                    DecisionNumber = txtSalaryHistDecisionNo.Text,
                    DecisionDate = dpSalaryHistDecisionDate.SelectedDate,
                    PersonnelId = _personnel.Id
                };
                _personnel.SalaryRecords.Add(newItem);
                ReconcileSalaryRecordsAfterAdd(_personnel.SalaryRecords, newItem);
            }

            RefreshSalaryHistoryGrid();

            // Clear inputs
            dpSalaryHistStart.SelectedDate = null;
            dpSalaryHistEnd.SelectedDate = null;
            dpSalaryHistCalc.SelectedDate = null;
            cboSalaryHistRankCode.SelectedIndex = -1;
            cboSalaryHistStep.ItemsSource = null;
            txtSalaryHistCoeff.Clear();
            txtSalaryHistPercent.Text = "100";
            txtSalaryHistExceedPercent.Text = "0";
            txtSalaryHistDecisionNo.Clear();
            dpSalaryHistDecisionDate.SelectedDate = null;
        }

        private void dgSalaryHistory_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            var dep = (DependencyObject)e.OriginalSource;
            while ((dep != null) && !(dep is DataGridRow))
            {
                dep = VisualTreeHelper.GetParent(dep);
            }

            if (dep is DataGridRow row)
            {
                var item = row.DataContext as SalaryRecord;
                if (item == null) return;

                // Load data back to form
                dpSalaryHistStart.SelectedDate = item.StartDate;
                dpSalaryHistEnd.SelectedDate = item.EndDate;
                dpSalaryHistCalc.SelectedDate = item.SalaryCalculationDate;
                cboSalaryHistRankCode.SelectedValue = item.RankCode;
                if (!string.IsNullOrEmpty(item.RankCode))
                {
                    LoadSalaryHistSteps(item.RankCode);
                    cboSalaryHistStep.SelectedValue = item.SalaryStep;
                }
                txtSalaryHistCoeff.Text = item.Coefficient;
                txtSalaryHistPercent.Text = item.Percentage.ToString();
                txtSalaryHistExceedPercent.Text = item.ExceedFramePercent.ToString();
                txtSalaryHistDecisionNo.Text = item.DecisionNumber;
                dpSalaryHistDecisionDate.SelectedDate = item.DecisionDate;

                // Set as editing so we don't duplicate on save
                _editingSalaryRecord = item;
            }
        }

        private void btnDeleteSalaryHistory_Click(object sender, RoutedEventArgs e)
        {
            var item = (sender as FrameworkElement)?.DataContext as SalaryRecord;
            if (item != null && _personnel?.SalaryRecords != null)
            {
                var confirm = new ConfirmDialog("Bạn có chắc chắn muốn xóa dòng quá trình lương này không?");
                if (confirm.ShowDialog() == true)
                {
                    DateTime? deletedStartDate = item.StartDate;
                    _personnel.SalaryRecords.Remove(item);
                    ReconcileSalaryRecordsAfterDelete(_personnel.SalaryRecords, deletedStartDate);
                    RefreshSalaryHistoryGrid();
                }
            }
        }

        private void btnAddLeave_Click(object? sender, RoutedEventArgs e)
        {
            // Validate input
            if (cboLeaveType.SelectedIndex == -1)
            {
                new WarningWindow("Vui lòng chọn loại nghỉ phép!", "Thiếu thông tin").ShowDialog();
                return;
            }
            var type = (cboLeaveType.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "";

            // Validate dates
            if (dpLeaveStartDate.SelectedDate == null)
            {
                new WarningWindow("Vui lòng nhập ngày bắt đầu!", "Thiếu thông tin").ShowDialog();
                return;
            }

            // End date is mandatory for Annual Leave ("Phép năm")
            if (type == "Phép năm" && dpLeaveEndDate.SelectedDate == null)
            {
                new WarningWindow("Vui lòng nhập ngày kết thúc cho Phép năm!", "Thiếu thông tin").ShowDialog();
                return;
            }

            if (dpLeaveEndDate.SelectedDate != null && dpLeaveEndDate.SelectedDate < dpLeaveStartDate.SelectedDate)
            {
                new WarningWindow("Ngày kết thúc không thể nhỏ hơn ngày bắt đầu!", "Lỗi thời gian").ShowDialog();
                return;
            }


            var start = dpLeaveStartDate.SelectedDate.Value;
            DateTime? end = dpLeaveEndDate.SelectedDate;


            // Check for overlap
            if (_personnel != null && _personnel.LeaveHistories != null)
            {
                foreach (var history in _personnel.LeaveHistories)
                {
                    // Skip self if editing
                    if (_editingLeaveHistory != null && history == _editingLeaveHistory) continue;

                    // Handle overlap with ongoing or closed leaves
                    bool isOverlap = false;

                    if (history.EndDate.HasValue)
                    {
                        if (end.HasValue)
                        {
                            // Case: Both have end dates
                            isOverlap = (start <= history.EndDate.Value && end.Value >= history.StartDate);
                        }
                        else
                        {
                            // Case: New is ongoing
                            isOverlap = (start <= history.EndDate.Value);
                        }
                    }
                    else
                    {
                        // Existing is ongoing
                        if (end.HasValue)
                        {
                            isOverlap = (end.Value >= history.StartDate);
                        }
                        else
                        {
                            // Both ongoing
                            isOverlap = true;
                        }
                    }

                    if (isOverlap)
                    {
                        string historyEndStr = history.EndDate.HasValue ? history.EndDate.Value.ToString("dd/MM/yyyy") : "Đang nghỉ";
                        var msg = $"Khoảng thời gian này đã trùng với đợt nghỉ phép: {history.LeaveType} \n(Từ {history.StartDate:dd/MM/yyyy} đến {historyEndStr}).\nVui lòng kiểm tra lại!";
                        new WarningWindow(msg, "Trùng lịch nghỉ").ShowDialog();
                        return;
                    }

                }
            }

            double duration = 0;
            if (double.TryParse(txtLeaveDuration.Text, out double manualDuration))
            {
                duration = manualDuration;
            }
            else
            {
                // Fallback if empty or invalid, though UpdateLeaveDuration should handle this
                if (type == "Phép năm")
                {
                    // end is guaranteed non-null here due to validation above
                    duration = CalculateWorkingDays(start, end!.Value);
                }
                else
                {
                    if (end.HasValue)
                        duration = (end.Value - start).TotalDays + 1;
                    else
                        duration = 0; // Unknown duration for ongoing leave
                }

            }

            // --- VALIDATION START: Check Balance ---
            if (type == "Phép năm")
            {
                int currentYear = DateTime.Now.Year;
                int selectedYear = (cboLeaveYear.SelectedItem as int?) ?? currentYear;
                double totalAvailable = 0;

                // Scenario 1: Adding New Record for Current Year (Eligible for Split/Carry-over)
                if (_editingLeaveHistory == null && selectedYear == currentYear)
                {
                    double remainingCurrent = GetRemainingAnnualLeaveDays(currentYear);
                    double remainingOld = GetRemainingAnnualLeaveDays(currentYear - 1);

                    if (chkPrioritizeOldYear.IsChecked == true)
                        totalAvailable = remainingCurrent + (remainingOld > 0 ? remainingOld : 0);
                    else
                        totalAvailable = remainingCurrent;
                }
                // Scenario 2: Editing OR Specific Past Year (Strict Limit)
                else
                {
                    totalAvailable = GetRemainingAnnualLeaveDays(selectedYear);

                    // If editing, add back the *original* duration to the available pool
                    if (_editingLeaveHistory != null &&
                        _editingLeaveHistory.LeaveType == "Phép năm" &&
                        (_editingLeaveHistory.LeaveYear ?? _editingLeaveHistory.StartDate.Year) == selectedYear)
                    {
                        totalAvailable += _editingLeaveHistory.DurationDays;
                    }
                }

                if (duration > totalAvailable)
                {
                    new WarningWindow($"Số ngày nghỉ ({duration}) vượt quá số ngày phép còn lại ({totalAvailable})!", "Lỗi quá hạn mức").ShowDialog();
                    return;
                }
            }
            // --- VALIDATION END ---

            var reason = txtLeaveReason.Text;

            if (_editingLeaveHistory != null)
            {
                // Update Existing
                _editingLeaveHistory.LeaveType = type;
                _editingLeaveHistory.StartDate = start;
                _editingLeaveHistory.EndDate = end;
                _editingLeaveHistory.DurationDays = duration;
                _editingLeaveHistory.DurationDays = duration;
                _editingLeaveHistory.Reason = reason;

                if (type == "Phép năm")
                {
                    _editingLeaveHistory.LeaveYear = (int?)cboLeaveYear.SelectedItem;
                }
                else
                {
                    _editingLeaveHistory.LeaveYear = null;
                }

                // Save system note from txtSystemNote
                _editingLeaveHistory.SystemNote = string.IsNullOrWhiteSpace(txtSystemNote.Text) ? null : txtSystemNote.Text.Trim();



                // Reset Edit Mode
                _editingLeaveHistory = null;
                dgLeaveHistory.SelectedItem = null;

                // Reset Button Text
                // Find TextBlock inside button
                if (btnAddLeave.Content is StackPanel sp)
                {
                    foreach (var child in sp.Children)
                    {
                        if (child is TextBlock tb)
                        {
                            tb.Text = "Thêm vào bảng";
                            break;
                        }
                    }
                }
            }
            else

            {
                // NEW LOGIC: Check for Auto-Deduct (Splitting Records)
                // NEW LOGIC: Smart Consumption (Prioritize Previous Year)
                if (type == "Phép năm")
                {
                    int currentYear = DateTime.Now.Year;
                    int previousYear = currentYear - 1;

                    // Check if there is remaining leave in the PREVIOUS year AND user wants to prioritize it
                    double remainingOld = GetRemainingAnnualLeaveDays(previousYear);
                    int selectedYear = (cboLeaveYear.SelectedItem as int?) ?? currentYear;

                    if (chkPrioritizeOldYear.IsChecked == true && selectedYear == currentYear && remainingOld > 0)
                    {
                        // 1. Take from Previous Year first
                        double takenFromOld = Math.Min(duration, remainingOld);
                        double takenFromCurrent = duration - takenFromOld;

                        // Generate Link ID
                        string linkId = Guid.NewGuid().ToString();

                        // Create Old Year Record
                        if (takenFromOld > 0)
                        {
                            var itemOld = new LeaveHistory
                            {
                                LeaveType = type,
                                StartDate = start,
                                EndDate = start.AddDays(takenFromOld - 1),
                                DurationDays = takenFromOld,
                                Reason = reason + $"|SYS:Ưu tiên trừ phép tồn năm {previousYear}|LINK:{linkId}",
                                PersonnelId = _personnel != null ? _personnel.Id : 0,
                                LeaveYear = previousYear,
                                SystemNote = string.IsNullOrWhiteSpace(txtSystemNote.Text) ? null : txtSystemNote.Text.Trim()
                            };


                            if (_personnel == null) _personnel = new Personnel();
                            if (_personnel.LeaveHistories == null) _personnel.LeaveHistories = new List<LeaveHistory>();
                            _personnel.LeaveHistories.Add(itemOld);
                        }

                        // Create Current Year Record (if overflow)
                        if (takenFromCurrent > 0)
                        {
                            var itemCurrent = new LeaveHistory
                            {
                                LeaveType = type,
                                StartDate = start, // Simplified date logic as per user request for auto-deduct
                                EndDate = end,
                                DurationDays = takenFromCurrent,
                                Reason = reason + $"|SYS:Trừ tiếp vào phép tồn năm {currentYear}|LINK:{linkId}",
                                PersonnelId = _personnel != null ? _personnel.Id : 0,
                                LeaveYear = currentYear,
                                SystemNote = string.IsNullOrWhiteSpace(txtSystemNote.Text) ? null : txtSystemNote.Text.Trim()
                            };


                            if (_personnel == null) _personnel = new Personnel();
                            if (_personnel.LeaveHistories == null) _personnel.LeaveHistories = new List<LeaveHistory>();
                            _personnel.LeaveHistories.Add(itemCurrent);
                        }

                        CalculateAnnualLeave(DateTime.Now); // Refresh totals
                        RefreshLeaveHistoryGrid();
                        UpdateLeaveStatistics(); // Show stats

                        string msg = $"Đã ưu tiên trừ {takenFromOld} ngày từ phép tồn năm {previousYear}.";
                        if (takenFromCurrent > 0)
                        {
                            msg += $"\nHệ thống tự động chuyển {takenFromCurrent} ngày còn lại vào phép năm {currentYear}.";
                        }

                        new NotificationWindow(msg, "Thông báo tự động").ShowDialog();

                        // Cleanup UI
                        cboLeaveType.SelectedIndex = -1;
                        cboLeaveYear.IsEnabled = false;
                        cboLeaveYear.SelectedItem = DateTime.Now.Year;
                        dpLeaveStartDate.SelectedDate = null;
                        dpLeaveEndDate.SelectedDate = null;
                        txtLeaveDuration.Clear();
                        txtLeaveReason.Clear();
                        txtSystemNote.Clear();
                        return; // EXIT FUNCTION
                    }
                }

                // Standard Logic (No split)
                var newItem = new LeaveHistory
                {
                    LeaveType = type,
                    StartDate = start,
                    EndDate = end,
                    DurationDays = duration,
                    Reason = reason,
                    PersonnelId = _personnel != null ? _personnel.Id : 0,
                    LeaveYear = (cboLeaveType.SelectedItem as ComboBoxItem)?.Content?.ToString() == "Phép năm" ? (int?)cboLeaveYear.SelectedItem : null,
                    SystemNote = string.IsNullOrWhiteSpace(txtSystemNote.Text) ? null : txtSystemNote.Text.Trim()
                };



                if (_personnel == null) _personnel = new Personnel();
                if (_personnel.LeaveHistories == null) _personnel.LeaveHistories = new System.Collections.Generic.List<LeaveHistory>();

                _personnel.LeaveHistories.Add(newItem);
            }

            _isRefreshing = false;
            // Refresh Grid
            RefreshLeaveHistoryGrid();

            UpdateLeaveStatistics();

            // Clear inputs
            cboLeaveType.SelectedIndex = -1;
            cboLeaveYear.IsEnabled = false;
            cboLeaveYear.SelectedItem = DateTime.Now.Year;
            dpLeaveStartDate.SelectedDate = null;
            dpLeaveEndDate.SelectedDate = null;
            txtLeaveDuration.Clear();
            txtLeaveReason.Clear();
            txtSystemNote.Clear();
        }

        private void OnLeaveDateChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (sender != dpLeaveEndDate)
            {
                CalculateMaternityEndDate();
            }
            UpdateLeaveDuration();

            // Enable Year Selection if "Phép năm"
            var type = (cboLeaveType.SelectedItem as ComboBoxItem)?.Content?.ToString();
            if (type == "Phép năm")
            {
                cboLeaveYear.IsEnabled = true;
                if (!_isRefreshing) LoadLeaveYears(); // Don't reload if we are already in selection changed from LoadLeaveYears

                int selectedYear = (cboLeaveYear.SelectedItem as int?) ?? DateTime.Now.Year;
                if (selectedYear == DateTime.Now.Year)
                {
                    chkPrioritizeOldYear.Visibility = Visibility.Visible;
                }
                else
                {
                    chkPrioritizeOldYear.Visibility = Visibility.Collapsed;
                    chkPrioritizeOldYear.IsChecked = false;
                }
            }
            else
            {
                cboLeaveYear.IsEnabled = false;
                chkPrioritizeOldYear.Visibility = Visibility.Collapsed;
                chkPrioritizeOldYear.IsChecked = false;
            }

            // Trigger Stat Update for selected year
            UpdateLeaveStatistics();
        }

        private void OnLeaveReasonTextChanged(object sender, TextChangedEventArgs e)
        {
            CalculateMaternityEndDate();
            UpdateLeaveDuration();
        }

        private void SetComboBoxByContent(ComboBox cbo, string? content)
        {
            if (string.IsNullOrEmpty(content))
            {
                cbo.SelectedIndex = -1;
                return;
            }

            foreach (var item in cbo.Items)
            {
                if (item is ComboBoxItem cbi && cbi.Content?.ToString() == content)
                {
                    cbo.SelectedItem = cbi;
                    return;
                }
                else if (item?.ToString() == content)
                {
                    cbo.SelectedItem = item;
                    return;
                }
            }

            // Fallback for non-item list (like strings)
            cbo.Text = content;
        }

        private void SetSalaryDelaySelection(string? delayType)
        {
            if (string.IsNullOrEmpty(delayType) || delayType == "-- Không lùi --" || delayType == "-- Không lương --" || delayType.StartsWith("Nghỉ không lương (", StringComparison.OrdinalIgnoreCase))
            {
                if (cboSalaryDelay.Items.Count > 0)
                {
                    cboSalaryDelay.SelectedIndex = 0;
                }
                else
                {
                    cboSalaryDelay.Text = "-- Không lùi --";
                }
            }
            else
            {
                SetComboBoxByContent(cboSalaryDelay, delayType);
            }
        }

        private void CalculateMaternityEndDate()
        {
            if (_isRefreshing) return;

            var type = (cboLeaveType.SelectedItem as ComboBoxItem)?.Content?.ToString();
            if (type != "Thai sản") return;

            if (dpLeaveStartDate.SelectedDate == null) return;

            DateTime start = dpLeaveStartDate.SelectedDate.Value;
            DateTime cutoff = new DateTime(2026, 7, 1);
            
            // Base maternity leave duration: 7 months from 1/7/2026 onwards, otherwise 6 months
            int months = (start < cutoff) ? 6 : 7;

            // Plus 1 month if "sinh đôi" is in the reason
            string reason = txtLeaveReason?.Text ?? "";
            if (reason.ToLower().Contains("sinh đôi"))
            {
                months += 1;
            }

            DateTime end = start.AddMonths(months).AddDays(-1);

            if (dpLeaveEndDate.SelectedDate != end)
            {
                dpLeaveEndDate.SelectedDate = end;
            }
        }

        private void UpdateLeaveDuration()
        {
            // Debug logging
            // MessageBox.Show($"Update triggered. Refreshing: {_isRefreshing}. Start: {dpLeaveStartDate.SelectedDate}, End: {dpLeaveEndDate.SelectedDate}");

            if (_isRefreshing) return;

            if (dpLeaveStartDate.SelectedDate == null || dpLeaveEndDate.SelectedDate == null)
            {
                txtLeaveDuration.Text = "";
                txtSystemNote.Text = "";
                return;
            }

            var start = dpLeaveStartDate.SelectedDate.Value;
            var end = dpLeaveEndDate.SelectedDate.Value;

            if (end < start)
            {
                txtLeaveDuration.Text = "";
                txtSystemNote.Text = "";
                return;
            }

            var type = (cboLeaveType.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "";
            double duration = 0;

            if (type == "Phép năm")
            {
                duration = CalculateWorkingDays(start, end);
            }
            else
            {
                duration = (end - start).TotalDays + 1;
            }

            txtLeaveDuration.Text = duration.ToString();

            // Auto-calculate System Note
            txtSystemNote.Text = GetLeaveBreakdownNote(start, end, type);
        }

        private double CalculateWorkingDays(DateTime start, DateTime end)
        {
            double days = 0;
            using (var context = new AppDbContext())
            {
                var holidays = context.PublicHolidays
                                      .Where(h => h.Date >= start.Date && h.Date <= end.Date)
                                      .Select(h => h.Date.Date)
                                      .ToList();

                for (var date = start.Date; date <= end.Date; date = date.AddDays(1))
                {
                    // Exclude weekends
                    if (date.DayOfWeek == DayOfWeek.Saturday || date.DayOfWeek == DayOfWeek.Sunday)
                        continue;

                    // Exclude public holidays
                    if (holidays.Contains(date))
                        continue;

                    days++;
                }
            }
            return days;
        }

        private string GetLeaveBreakdownNote(DateTime start, DateTime end, string leaveType)
        {
            int totalDays = (int)(end.Date - start.Date).TotalDays + 1;
            int weekends = 0;
            int holidaysCount = 0;
            int workDays = 0;

            using (var context = new AppDbContext())
            {
                var holidaysList = context.PublicHolidays
                                          .Where(h => h.Date >= start.Date && h.Date <= end.Date)
                                          .Select(h => h.Date.Date)
                                          .ToList();

                for (var date = start.Date; date <= end.Date; date = date.AddDays(1))
                {
                    bool isWeekend = (date.DayOfWeek == DayOfWeek.Saturday || date.DayOfWeek == DayOfWeek.Sunday);
                    bool isHoliday = holidaysList.Contains(date);

                    if (isWeekend)
                    {
                        weekends++;
                    }
                    else if (isHoliday)
                    {
                        holidaysCount++;
                    }
                    else
                    {
                        workDays++;
                    }
                }
            }

            string typeName = leaveType == "Phép năm" ? "nghỉ phép năm" : "ngày nghỉ";
            return $"Nghỉ {totalDays:00} ngày, trong đó: {weekends:00} ngày thứ 7, chủ nhật; {holidaysCount:00} ngày nghỉ lễ; {workDays:00} {typeName}";
        }



        private LeaveHistory? _editingLeaveHistory = null;

        private void dgLeaveHistory_SelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (dgLeaveHistory.SelectedItem is LeaveHistory item)
            {
                _isRefreshing = true;
                _editingLeaveHistory = item;

                // Populate fields
                foreach (ComboBoxItem cbi in cboLeaveType.Items)
                {
                    if (cbi.Content?.ToString() == item.LeaveType)
                    {
                        cboLeaveType.SelectedItem = cbi;
                        break;
                    }
                }

                dpLeaveStartDate.SelectedDate = item.StartDate;
                dpLeaveEndDate.SelectedDate = item.EndDate;
                txtLeaveDuration.Text = item.DurationDays.ToString();
                txtLeaveReason.Text = item.UserReasonDisplay; // Only show user part
                txtSystemNote.Text = item.SystemNote ?? "";

                if (item.LeaveType == "Phép năm")
                {
                    if (item.LeaveYear.HasValue)
                    {
                        LoadLeaveYears(item.LeaveYear.Value); // Ensure the edited year is visible
                        cboLeaveYear.SelectedItem = item.LeaveYear.Value;
                        cboLeaveYear.IsEnabled = true;
                    }
                    else
                    {
                        // Fallback to start date year if LeaveYear is missing
                        int startYear = item.StartDate.Year;
                        LoadLeaveYears(startYear);
                        cboLeaveYear.SelectedItem = startYear;
                        cboLeaveYear.IsEnabled = true;
                    }

                    // NEW: Detect Priority from Reason
                    bool isPriority = !string.IsNullOrEmpty(item.Reason) &&
                                     (item.Reason.Contains("|SYS:Ưu tiên trừ phép tồn năm") ||
                                      item.Reason.Contains("|SYS:Trừ tiếp vào phép tồn năm"));

                    if (isPriority)
                    {
                        chkPrioritizeOldYear.Visibility = Visibility.Visible;
                        chkPrioritizeOldYear.IsChecked = true;
                    }
                }
                else
                {
                    cboLeaveYear.IsEnabled = false;
                }

                _isRefreshing = false;

                // Change Button Text
                // Find TextBlock inside button
                if (btnAddLeave.Content is StackPanel sp)
                {
                    foreach (var child in sp.Children)
                    {
                        if (child is TextBlock tb)
                        {
                            tb.Text = "Cập nhật ngày nghỉ";
                            break;
                        }
                    }
                }
            }
        }

        private void btnDeleteLeave_Click(object? sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is int id)
            {
                // Note: If ID is 0 (newly added), we need to find by object reference logic or assume last one? 
                // Or simply find by object in list matching the bounded item?
                // The Tag binding binds to Id, which might be 0 for new items. 
                // Better approach: Get DataContext of the row.
                var item = btn.DataContext as LeaveHistory;
                if (item != null && _personnel?.LeaveHistories != null)
                {
                    // Check for Linked Records
                    var linkedItems = new List<LeaveHistory>();
                    if (!string.IsNullOrEmpty(item.LinkId))
                    {
                        linkedItems = _personnel.LeaveHistories
                            .Where(x => x.LinkId == item.LinkId && x != item)
                            .ToList();
                    }

                    if (linkedItems.Any())
                    {
                        // Simplified Logic: Always delete ALL linked records
                        var confirm = new ConfirmWindow(
                            $"Bản ghi này thuộc về một đợt nghỉ phép bị tách ({linkedItems.Count + 1} phần).\n" +
                            "Hệ thống sẽ xóa TẤT CẢ các bản ghi liên quan để đảm bảo đồng bộ.\n" +
                            "Bạn có chắc chắn muốn xóa không?",
                            "Xác nhận xóa liên kết");

                        if (confirm.ShowDialog() == true)
                        {
                            _personnel.LeaveHistories.Remove(item);
                            foreach (var link in linkedItems)
                            {
                                _personnel.LeaveHistories.Remove(link);
                            }
                        }
                    }
                    else
                    {
                        // Standard delete
                        var confirm = new ConfirmWindow($"Bạn có chắc muốn xóa lịch sử nghỉ: {item.LeaveType}?", "Xác nhận xóa");
                        if (confirm.ShowDialog() == true)
                        {
                            _personnel.LeaveHistories.Remove(item);
                        }
                    }

                    // Refresh Grid
                    RefreshLeaveHistoryGrid();
                    UpdateLeaveStatistics();
                }
            }
        }



        private void UpdateLeaveStatistics(int? yearOverride = null)
        {
            // if (_personnel == null) return; // Removed to allow UI update with override values
            if (_personnel != null && _personnel.LeaveHistories == null) _personnel.LeaveHistories = new List<LeaveHistory>();

            var histories = _personnel?.LeaveHistories ?? new List<LeaveHistory>();

            double annualTakenCurrentYear = 0;
            double annualTakenOldYear = 0;
            double sickTaken = 0;

            // Maternity Logic
            string maternityDisplay = "0";
            // Find ALL Maternity records, sort by StartDate descending
            var maternityLeaves = histories
                                    .Where(x => x.LeaveType == "Thai sản")
                                    .OrderByDescending(x => x.StartDate)
                                    .ToList();

            if (maternityLeaves.Any())
            {
                // Display the range of the LATEST maternity leave
                var latest = maternityLeaves.First();
                // Since LeaveHistory properties are likely non-nullable DateTime based on the error
                maternityDisplay = $"{latest.StartDate:dd/MM/yyyy} - {latest.EndDate:dd/MM/yyyy}";
            }

            // Determine which year to calculate for
            int yearToCalculate = yearOverride ?? DateTime.Now.Year;

            // If user selected a year in ComboBox, we might want to prioritize that for display
            if (yearOverride == null && cboLeaveYear.SelectedItem is int selectedYear)
            {
                yearToCalculate = selectedYear;
            }

            double totalUnpaidTakenAllYears = 0; // NEW: Track unpaid leave across ALL years for salary delay

            foreach (var item in histories)
            {
                // Always accumulate unpaid leave across ALL years for salary delay calculation
                if (item.LeaveType == "Không lương")
                {
                    totalUnpaidTakenAllYears += item.DurationDays;
                }

                // We want to process all leaves that were TAKEN in this year
                if (item.StartDate.Year == yearToCalculate)
                {
                    if (item.LeaveType == "Phép năm")
                    {
                        // Check if it's from old year quota
                        // It's old year if LeaveYear is explicitly set to a previous year
                        if (item.LeaveYear.HasValue && item.LeaveYear.Value < yearToCalculate)
                        {
                            annualTakenOldYear += item.DurationDays;
                        }
                        else
                        {
                            annualTakenCurrentYear += item.DurationDays;
                        }
                    }
                    else if (item.LeaveType == "Nghỉ ốm") sickTaken += item.DurationDays;
                }
            }

            // Bind to UI
            // Assuming default 12 if not set in UI text
            int totalAnnual = 12;

            // Prefer Model over Text parsing
            if (_personnel != null && _personnel.TotalAnnualLeaveDays > 0)
            {
                totalAnnual = _personnel.TotalAnnualLeaveDays;
            }
            else
            {
                int.TryParse(txtTotalAnnualLeave.Text, out totalAnnual);
            }

            // Bind
            txtAnnualLeaveTaken.Text = (annualTakenCurrentYear + annualTakenOldYear).ToString();
            txtAnnualLeaveRemaining.Text = (totalAnnual - annualTakenCurrentYear).ToString();
            txtSickLeaveTaken.Text = sickTaken.ToString();
            txtUnpaidLeaveTaken.Text = totalUnpaidTakenAllYears.ToString();
            txtMaternityLeaveTaken.Text = maternityDisplay;

            CalculateExpectedSalaryDate();
        }

        private double GetRemainingAnnualLeaveDays(int year)
        {
            // Default to 12 or value from UI if personnel is null/empty
            int total = 12;

            if (_personnel != null && _personnel.TotalAnnualLeaveDays > 0)
            {
                total = _personnel.TotalAnnualLeaveDays;
            }
            else
            {
                // Fallback to UI text if model not set (e.g. New User)
                if (int.TryParse(txtTotalAnnualLeave.Text, out int uiTotal))
                {
                    total = uiTotal;
                }
            }

            double taken = 0;
            if (_personnel != null && _personnel.LeaveHistories != null)
            {
                taken = _personnel.LeaveHistories
                     .Where(x => x.LeaveType == "Phép năm" && (x.LeaveYear ?? x.StartDate.Year) == year)
                     .Sum(x => x.DurationDays);
            }

            return total - taken;
        }

        private void RefreshLeaveHistoryGrid()
        {
            if (_personnel == null || _personnel.LeaveHistories == null) return;

            // Auto-clean: Remove old years (Keep only current year OR overlapping)
            // Rule: Remove if EndDate is in a previous year.
            // MODIFIED: Keep all history as per user request
            // int currentYear = DateTime.Now.Year;
            // _personnel.LeaveHistories.RemoveAll(x => x.EndDate.Year < currentYear);

            // Sort by StartDate Descending, then by Id Descending (newest leaves/recently added ones are placed on top)
            var sortedList = _personnel.LeaveHistories
                                .OrderByDescending(x => x.StartDate)
                                .ThenByDescending(x => x.Id)
                                .ToList();

            // Do not group anymore - show delete button for every individual item
            foreach (var item in sortedList)
            {
                item.ShowDeleteButton = true;
                item.IsLinkedGroupHead = false;
            }

            _personnel.LeaveHistories = sortedList;

            int stt = 1;
            foreach (var item in _personnel.LeaveHistories)
            {
                item.STT = stt++;
            }

            dgLeaveHistory.ItemsSource = null;
            dgLeaveHistory.ItemsSource = _personnel.LeaveHistories;
        }
        private void OnSalaryStructureChanged(object? sender, RoutedEventArgs e)
        {
            CalculateExpectedSalaryDate();
        }

        private void OnSalaryStructureChanged(object? sender, SelectionChangedEventArgs e)
        {
            CalculateExpectedSalaryDate();
        }

        private void CalculateExpectedSalaryDate()
        {
            if (_isRefreshing) return;

            // Handle dynamic insert/remove of "Nghỉ không lương (X ngày)" in cboSalaryDelay
            double unpaidDays = 0;
            bool hasUnpaidLeave = double.TryParse(txtUnpaidLeaveTaken.Text, out unpaidDays) && unpaidDays > 0;
            if (cboSalaryDelay.Items.Count > 0)
            {
                // Ensure the first item is always "-- Không lùi --"
                var firstItemObj = cboSalaryDelay.Items[0];
                string firstItemText = firstItemObj is ComboBoxItem cbi0 ? cbi0.Content?.ToString() ?? "" : firstItemObj.ToString() ?? "";
                if (firstItemText != "-- Không lùi --")
                {
                    bool wasRefreshing = _isRefreshing;
                    _isRefreshing = true;
                    try
                    {
                        if (firstItemObj is ComboBoxItem cbiUpdate)
                            cbiUpdate.Content = "-- Không lùi --";
                        else
                            cboSalaryDelay.Items[0] = "-- Không lùi --";
                    }
                    finally
                    {
                        _isRefreshing = wasRefreshing;
                    }
                }

                // Check if any item starts with "Nghỉ không lương (" in cboSalaryDelay.Items
                int unpaidItemIndex = -1;
                string currentUnpaidTextInList = "";
                for (int i = 0; i < cboSalaryDelay.Items.Count; i++)
                {
                    var itemObj = cboSalaryDelay.Items[i];
                    string itemText = itemObj is ComboBoxItem cbiItem ? cbiItem.Content?.ToString() ?? "" : itemObj.ToString() ?? "";
                    if (itemText.StartsWith("Nghỉ không lương (", StringComparison.OrdinalIgnoreCase))
                    {
                        unpaidItemIndex = i;
                        currentUnpaidTextInList = itemText;
                        break;
                    }
                }

                string expectedUnpaidText = $"Nghỉ không lương ({unpaidDays} ngày)";

                if (hasUnpaidLeave)
                {
                    if (unpaidItemIndex == -1)
                    {
                        bool wasRefreshing = _isRefreshing;
                        _isRefreshing = true;
                        try
                        {
                            // Insert at index 1 (right after "-- Không lùi --")
                            cboSalaryDelay.Items.Insert(1, expectedUnpaidText);
                            
                            // If selectedIndex was 0, -1, or "-- Không lương --", set it to 1 ("Nghỉ không lương (X ngày)")
                            int selectedIndex = cboSalaryDelay.SelectedIndex;
                            string selectedText = cboSalaryDelay.SelectedItem is ComboBoxItem cbiSel ? cbiSel.Content?.ToString() ?? "" : cboSalaryDelay.SelectedItem?.ToString() ?? "";
                            if (selectedIndex == 0 || selectedIndex == -1 || selectedText == "-- Không lương --" || selectedText == "-- Không lùi --")
                            {
                                cboSalaryDelay.SelectedIndex = 1;
                            }
                        }
                        finally
                        {
                            _isRefreshing = wasRefreshing;
                        }
                    }
                    else
                    {
                        // Check if the current selected index is 0 or -1 or "-- Không lương --" / "-- Không lùi --", and force it to the unpaid item index
                        int selectedIndex = cboSalaryDelay.SelectedIndex;
                        string selectedText = cboSalaryDelay.SelectedItem is ComboBoxItem cbiSel ? cbiSel.Content?.ToString() ?? "" : cboSalaryDelay.SelectedItem?.ToString() ?? "";
                        if (selectedIndex == 0 || selectedIndex == -1 || selectedText == "-- Không lương --" || selectedText == "-- Không lùi --")
                        {
                            bool wasRefreshing = _isRefreshing;
                            _isRefreshing = true;
                            try
                            {
                                cboSalaryDelay.SelectedIndex = unpaidItemIndex;
                            }
                            finally
                            {
                                _isRefreshing = wasRefreshing;
                            }
                        }

                        if (currentUnpaidTextInList != expectedUnpaidText)
                        {
                            // Update the text in place
                            bool wasRefreshing = _isRefreshing;
                            _isRefreshing = true;
                            try
                            {
                                var existingItem = cboSalaryDelay.Items[unpaidItemIndex];
                                if (existingItem is ComboBoxItem cbiUpdate)
                                {
                                    cbiUpdate.Content = expectedUnpaidText;
                                }
                                else
                                {
                                    cboSalaryDelay.Items[unpaidItemIndex] = expectedUnpaidText;
                                }
                                
                                // Force selection refresh if this item is selected
                                if (cboSalaryDelay.SelectedIndex == unpaidItemIndex)
                                {
                                    cboSalaryDelay.SelectedIndex = -1;
                                    cboSalaryDelay.SelectedIndex = unpaidItemIndex;
                                }
                            }
                            finally
                            {
                                _isRefreshing = wasRefreshing;
                            }
                        }
                    }
                }
                else
                {
                    if (unpaidItemIndex != -1)
                    {
                        bool wasRefreshing = _isRefreshing;
                        _isRefreshing = true;
                        try
                        {
                            int selectedIndex = cboSalaryDelay.SelectedIndex;
                            cboSalaryDelay.Items.RemoveAt(unpaidItemIndex);
                            if (selectedIndex == unpaidItemIndex)
                            {
                                cboSalaryDelay.SelectedIndex = 0; // Default back to "-- Không lùi --"
                            }
                        }
                        finally
                        {
                            _isRefreshing = wasRefreshing;
                        }
                    }
                }
            }

            // 1. Base Date: Next Salary Step Date
            if (!dpNextSalaryStepDate.SelectedDate.HasValue)
            {
                dpExpectedSalaryIncrease.SelectedDate = null;
                return;
            }

            DateTime baseDate = dpNextSalaryStepDate.SelectedDate.Value;
            DateTime expectedDate = baseDate;

            // 2. Period Calculation
            int periodYears = 3; // Default 3 years

            // Check Exceed Frame
            // "Tuy nhiên Có ai đã lên % vượt khung Thì 1 năm lên 1 lần"
            double exceedFrame = 0;
            if (double.TryParse(txtExceedFrame.Text, out exceedFrame) && exceedFrame > 0)
            {
                periodYears = 1;
            }
            else
            {
                // Check Rank Code
                // "Nếu mã ngach là 06.039-1, 01.011, 01.009 thì 2 năm lên 1 lần"
                string rankCode = "";
                if (cboRankCode.SelectedItem is Rank r) rankCode = r.Code;
                else if (cboRankCode.SelectedValue != null) rankCode = cboRankCode.SelectedValue.ToString() ?? "";
                else rankCode = cboRankCode.Text;

                rankCode = rankCode.Trim();

                if (string.Equals(rankCode, "06.039-1", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(rankCode, "01.011", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(rankCode, "01.009", StringComparison.OrdinalIgnoreCase))
                {
                    periodYears = 2;
                }
            }

            // Apply Period
            expectedDate = baseDate.AddYears(periodYears);

            // 3. Delay Calculation (Disciplinary)
            // "kèm thêm xét điều kiện ở ô lùi thời gian nâng lương"
            string delayReason = "";
            if (cboSalaryDelay.SelectedItem is ComboBoxItem cbi)
                delayReason = cbi.Content?.ToString() ?? "";
            else if (cboSalaryDelay.SelectedItem != null)
                delayReason = cboSalaryDelay.SelectedItem.ToString() ?? "";
            else
                delayReason = cboSalaryDelay.Text;

            if (!string.IsNullOrEmpty(delayReason))
            {
                if (delayReason.Contains("Lùi 3 tháng"))
                {
                    expectedDate = expectedDate.AddMonths(3);
                }
                else if (delayReason.Contains("Lùi 6 tháng"))
                {
                    expectedDate = expectedDate.AddMonths(6);
                }
                else if (delayReason.Contains("Lùi 12 tháng"))
                {
                    expectedDate = expectedDate.AddMonths(12);
                }
            }

            // 4. Automatic Delay from Unpaid Leave
            // "trong đó nếu nghỉ không lương thì Nghỉ bao lâu Thì lùi đúng số tháng nghỉ"
            if (delayReason != "-- Không lùi --" && unpaidDays > 0)
            {
                expectedDate = expectedDate.AddDays(unpaidDays);
            }

            // Set Result
            if (dpExpectedSalaryIncrease.SelectedDate != expectedDate)
            {
                dpExpectedSalaryIncrease.SelectedDate = expectedDate;
            }
        }

        private void NumberValidationTextBox(object sender, TextCompositionEventArgs e)
        {
            System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex("[^0-9]+");
            e.Handled = regex.IsMatch(e.Text);
        }

        private void txtPhone_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_isRefreshing || _isFormatting) return;

            _isFormatting = true;
            try
            {
                int caretIndex = txtPhone.CaretIndex;
                string originalText = txtPhone.Text;
                string digits = new string(originalText.Where(char.IsDigit).ToArray());

                // Limit to 10 digits
                if (digits.Length > 10) digits = digits.Substring(0, 10);

                string formatted = FormatPhoneNumber(digits);

                if (txtPhone.Text != formatted)
                {
                    txtPhone.Text = formatted;

                    // Restore caret position
                    int digitsBefore = originalText.Substring(0, Math.Min(caretIndex, originalText.Length)).Count(char.IsDigit);
                    int newCaretIndex = 0;
                    int digitCount = 0;
                    while (newCaretIndex < formatted.Length && digitCount < digitsBefore)
                    {
                        if (char.IsDigit(formatted[newCaretIndex])) digitCount++;
                        newCaretIndex++;
                    }

                    if (newCaretIndex < formatted.Length && formatted[newCaretIndex] == '.')
                    {
                        if (e.Changes.Any(c => c.AddedLength > 0))
                            newCaretIndex++;
                    }

                    txtPhone.CaretIndex = Math.Min(newCaretIndex, formatted.Length);
                }

                // Error validation logic
                if (digits.Length > 0 && !digits.StartsWith("0"))
                {
                    lblPhoneError.Text = "SĐT phải bắt đầu bằng số 0";
                    lblPhoneError.Visibility = Visibility.Visible;
                }
                else if (digits.Length > 0 && digits.Length < 10)
                {
                    lblPhoneError.Text = "SĐT phải có 10 chữ số";
                    lblPhoneError.Visibility = Visibility.Visible;
                }
                else
                {
                    lblPhoneError.Visibility = Visibility.Collapsed;
                }
            }
            finally
            {
                _isFormatting = false;
            }
        }

        private string FormatPhoneNumber(string? digits)
        {
            if (string.IsNullOrEmpty(digits)) return "";

            // Clean non-digits first
            digits = new string(digits.Where(char.IsDigit).ToArray());

            if (digits.Length == 0) return "";

            string formatted = digits.Substring(0, Math.Min(4, digits.Length));
            if (digits.Length > 4)
            {
                formatted += "." + digits.Substring(4, Math.Min(3, digits.Length - 4));
                if (digits.Length > 7)
                {
                    formatted += "." + digits.Substring(7, Math.Min(3, digits.Length - 7));
                }
            }
            return formatted;
        }

        private bool IsValidEmail(string email)
        {
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address == email;
            }
            catch
            {
                return false;
            }
        }

        private void LoadEvaluationYears()
        {
            var years = new List<string> { "" };
            int currentYear = DateTime.Now.Year;
            for (int y = currentYear; y >= 1980; y--)
            {
                years.Add(y.ToString());
            }
            cboEvaluationYear.ItemsSource = years;
            cboEvaluationYear.SelectedIndex = 0;
        }

        private void RefreshEvaluationHistoryGrid()
        {
            if (_personnel == null) return;
            if (_personnel.EvaluationRecords == null) _personnel.EvaluationRecords = new List<EvaluationRecord>();
            var sortedList = _personnel.EvaluationRecords.OrderByDescending(e => e.Year).ToList();
            dgEvaluationHistory.ItemsSource = null;
            dgEvaluationHistory.ItemsSource = sortedList;
        }

        private void btnAddEvaluation_Click(object sender, RoutedEventArgs e)
        {
            var yearStr = cboEvaluationYear.SelectedItem?.ToString() ?? "";
            if (string.IsNullOrEmpty(yearStr))
            {
                new WarningWindow("Vui lòng chọn năm xếp loại!", "Thiếu thông tin").ShowDialog();
                return;
            }
            if (!int.TryParse(yearStr, out int year))
            {
                new WarningWindow("Năm xếp loại không hợp lệ!", "Lỗi dữ liệu").ShowDialog();
                return;
            }

            var rating = cboEvaluationRating.SelectedItem is ComboBoxItem cbi ? cbi.Content?.ToString() ?? "" : cboEvaluationRating.Text;
            if (string.IsNullOrEmpty(rating))
            {
                new WarningWindow("Vui lòng chọn xếp loại năm!", "Thiếu thông tin").ShowDialog();
                return;
            }

            var decisionNo = txtEvaluationDecisionNo.Text;
            var decisionDate = dpEvaluationDecisionDate.SelectedDate;
            var decisionAgency = txtEvaluationDecisionAgency.Text;

            if (_personnel != null && _personnel.EvaluationRecords != null)
            {
                foreach (var ev in _personnel.EvaluationRecords)
                {
                    if (_editingEvaluationRecord != ev && ev.Year == year)
                    {
                        new WarningWindow($"Đã tồn tại xếp loại cho năm {year}!", "Trùng thông tin").ShowDialog();
                        return;
                    }
                }
            }

            if (_editingEvaluationRecord != null)
            {
                _editingEvaluationRecord.Year = year;
                _editingEvaluationRecord.Rating = rating;
                _editingEvaluationRecord.DecisionNumber = decisionNo;
                _editingEvaluationRecord.DecisionDate = decisionDate;
                _editingEvaluationRecord.DecisionAgency = decisionAgency;

                _editingEvaluationRecord = null;
                dgEvaluationHistory.SelectedItem = null;
            }
            else
            {
                var newItem = new EvaluationRecord
                {
                    Year = year,
                    Rating = rating,
                    DecisionNumber = decisionNo,
                    DecisionDate = decisionDate,
                    DecisionAgency = decisionAgency,
                    PersonnelId = _personnel != null ? _personnel.Id : 0
                };
                if (_personnel == null) _personnel = new Personnel();
                if (_personnel.EvaluationRecords == null) _personnel.EvaluationRecords = new List<EvaluationRecord>();
                _personnel.EvaluationRecords.Add(newItem);
            }

            RefreshEvaluationHistoryGrid();

            cboEvaluationYear.SelectedIndex = 0;
            cboEvaluationRating.SelectedIndex = 0;
            txtEvaluationDecisionNo.Clear();
            dpEvaluationDecisionDate.SelectedDate = null;
            txtEvaluationDecisionAgency.Clear();
        }

        private void dgEvaluationHistory_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            var dep = (DependencyObject)e.OriginalSource;
            while ((dep != null) && !(dep is DataGridRow))
            {
                dep = VisualTreeHelper.GetParent(dep);
            }

            if (dep is DataGridRow row)
            {
                var item = row.DataContext as EvaluationRecord;
                if (item == null) return;

                cboEvaluationYear.SelectedItem = item.Year.ToString();
                
                cboEvaluationRating.SelectedIndex = 0;
                foreach (var obj in cboEvaluationRating.Items)
                {
                    if (obj is ComboBoxItem cbi && cbi.Content?.ToString() == item.Rating)
                    {
                        cboEvaluationRating.SelectedItem = cbi;
                        break;
                    }
                }

                txtEvaluationDecisionNo.Text = item.DecisionNumber;
                dpEvaluationDecisionDate.SelectedDate = item.DecisionDate;
                txtEvaluationDecisionAgency.Text = item.DecisionAgency;

                _editingEvaluationRecord = item;
            }
        }

        private void btnDeleteEvaluation_Click(object sender, RoutedEventArgs e)
        {
            var item = (sender as FrameworkElement)?.DataContext as EvaluationRecord;
            if (item != null && _personnel?.EvaluationRecords != null)
            {
                var confirm = new ConfirmDialog("Bạn có chắc chắn muốn xóa dòng xếp loại này không?");
                if (confirm.ShowDialog() == true)
                {
                    _personnel.EvaluationRecords.Remove(item);
                    RefreshEvaluationHistoryGrid();

                    if (_editingEvaluationRecord == item)
                    {
                        _editingEvaluationRecord = null;
                        cboEvaluationYear.SelectedIndex = 0;
                        cboEvaluationRating.SelectedIndex = 0;
                        txtEvaluationDecisionNo.Clear();
                        dpEvaluationDecisionDate.SelectedDate = null;
                        txtEvaluationDecisionAgency.Clear();
                    }
                }
            }
        }
    }
}
