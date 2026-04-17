using System;
using System.Linq;
using System.Windows;
using MaterialDesignThemes.Wpf;
using System.Windows.Controls;
using TaxPersonnelManagement.Data;
using TaxPersonnelManagement.Models;
using Microsoft.EntityFrameworkCore;
namespace TaxPersonnelManagement.Views
{
    public partial class PersonnelDetailView : UserControl
    {
        private Personnel? _personnel;
        private string? _currentAvatarBase64;
        private bool _isAvatarChanged = false;
        private bool _isRefreshing = false;
        private bool _isFormatting = false;

        /// <summary>
        /// Khởi tạo màn hình chi tiết hồ sơ cán bộ.
        /// </summary>
        /// <param name="personnel">Đối tượng cán bộ cần hiển thị. Nếu thêm mới thì truyền null.</param>
        public PersonnelDetailView(Personnel? personnel, int activeTab = 0)
        {
            InitializeComponent();
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
                txtPhone.Text = _personnel.PhoneNumber;
                txtIdentityNumber.Text = _personnel.IdentityCardNumber;
                txtIdentityPlace.Text = _personnel.IdentityCardPlace;
                txtEmail.Text = _personnel.Email;
                txtSocialInsurance.Text = _personnel.SocialSecurityNumber;
                txtBirthPlace.Text = _personnel.BirthPlace;
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
                // Duplicate txtITSkill removed
                txtLanguageSkill.Text = _personnel.LanguageSkillLevel;

                // Tab 2
                dpPositionDecisionDate.SelectedDate = _personnel.PositionDecisionDate;
                dpPositionCalculationDate.SelectedDate = _personnel.PositionCalculationDate;
                txtPositionYear.Text = _personnel.PositionYear;
                txtDetailedWorkHistory.Text = _personnel.DetailedWorkHistory;
                // Duplicate txtDetailedWorkHistory removed
                CalculateWorkDuration();

                // Tab 3
                dpRetirementDate.SelectedDate = _personnel.RetirementDate;
                CalculateRetirementInfo();

                // Tab 4
                dpPartyEntryDate.SelectedDate = _personnel.PartyEntryDate;
                dpPartyOfficialDate.SelectedDate = _personnel.PartyOfficialDate;

                // Tab 6: Salary Info
                cboSalaryStep.SelectedValue = _personnel.CurrentSalaryStep; // Use SelectedValue
                txtSalaryCoefficient.Text = _personnel.CurrentSalaryCoefficient.ToString();
                txtExceedFrame.Text = _personnel.ExceedFramePercent.ToString();
                txtPositionAllowance.Text = _personnel.PositionAllowance;
                dpSalaryReservation.SelectedDate = _personnel.SalaryReservationDeadline;
                dpNextSalaryStepDate.SelectedDate = _personnel.NextSalaryStepDate;
                
                if (!string.IsNullOrEmpty(_personnel.SalaryIncreaseDelayType))
                    SetComboBoxByContent(cboSalaryDelay, _personnel.SalaryIncreaseDelayType);
                else
                    cboSalaryDelay.SelectedIndex = 0; // Default to first item ("-- Không lùi --")

                dpExpectedSalaryIncrease.SelectedDate = _personnel.ExpectedSalaryIncreaseDate;
                txtSalaryHistoryLog.Text = _personnel.SalaryHistoryLog;

                // Tab 5: Leave Info
                // txtTotalAnnualLeave.Text = _personnel.TotalAnnualLeaveDays.ToString(); // Handled by CalculateAnnualLeave
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
            ApplyAuthorization();

            _isRefreshing = false; // HOÀN TẤT NẠP DỮ LIỆU BAN ĐẦU
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
            using (var context = new AppDbContext())
            {
               // Load from Departments table + any legacy string departments in Personnel table
               var dbDepts = context.Departments.Select(d => d.Name).ToList();
               var distinctExisting = context.Personnel.Select(p => p.Department).Distinct().Where(d => !string.IsNullOrEmpty(d)).ToList();
               
               var allDepts = dbDepts.Union(distinctExisting!).Where(x => !string.IsNullOrEmpty(x)).OrderBy(x => x).ToList();
               
               cboDepartment.ItemsSource = allDepts;
                
                if (_personnel != null)
                {
                    SetComboBoxByContent(cboDepartment, _personnel.Department);
                }
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
            using (var context = new AppDbContext())
            {
               var dbPos = context.Positions.Select(p => p.Name).ToList();
               var distinctExisting = context.Personnel.Select(p => p.Position).Distinct().Where(p => !string.IsNullOrEmpty(p)).ToList();
               
               var allPos = dbPos.Union(distinctExisting!).Where(x => !string.IsNullOrEmpty(x)).OrderBy(x => x).ToList();
               
               cboPosition.ItemsSource = allPos;
                
                if (_personnel != null)
                {
                    SetComboBoxByContent(cboPosition, _personnel.Position);
                }
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
                var existingCodes = context.Personnel.Select(p => p.RankCode).Distinct().Where(c => !string.IsNullOrEmpty(c)).ToList();

                // Add any missing codes as dummy Ranks
                foreach (var code in existingCodes)
                {
                    if (!dbRanks.Any(r => r.Code == code))
                    {
                        dbRanks.Add(new Rank { Code = code!, Name = "" });
                    }
                }

                cboRankCode.ItemsSource = dbRanks.OrderBy(r => r.Code).ToList();
                cboRankCode.DisplayMemberPath = "Code";
                cboRankCode.SelectedValuePath = "Code";

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
                                try {
                                    var result = command.ExecuteScalar();
                                    if (result != null) count = Convert.ToInt32(result);
                                } catch {}
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
                    
                    if (_personnel != null && !string.IsNullOrEmpty(_personnel.SalaryIncreaseDelayType))
                    {
                        SetComboBoxByContent(cboSalaryDelay, _personnel.SalaryIncreaseDelayType);
                    }
                    else
                    {
                        // Default to first item ("-- Không lùi --")
                        if (cboSalaryDelay.Items.Count > 0)
                            cboSalaryDelay.SelectedIndex = 0;
                        else 
                            cboSalaryDelay.Text = "-- Không lùi --";
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
            if (cboRankCode.SelectedItem is Rank selectedRank)
            {
                if (!string.IsNullOrEmpty(selectedRank.Name)) 
                    txtRankName.Text = selectedRank.Name;
                
                // Debug
                // MessageBox.Show($"Rank Selected: {selectedRank.Code}");
                LoadSalarySteps(selectedRank.Code);
            }
            else
            {
                // Clear steps if rank is deslected
                cboSalaryStep.ItemsSource = null;
                txtRankName.Text = "";
            }
            CalculateExpectedSalaryDate();
        }

        private void LoadSalarySteps(string rankCode)
        {
            try
            {
                using (var context = new AppDbContext())
                {
                    var steps = context.RankSalarySpecs
                                       .Where(s => s.RankCode == rankCode)
                                       .OrderBy(s => s.SalaryStep)
                                       .ToList();

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
                // Auto-fill coefficient
                txtSalaryCoefficient.Text = spec.Coefficient.ToString("N2", System.Globalization.CultureInfo.GetCultureInfo("vi-VN"));
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
            if (!string.IsNullOrEmpty(txtPhone.Text) && (phoneDigits.Length != 10 || !txtPhone.Text.StartsWith("0")))
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
                        PhoneNumber = txtPhone.Text,
                        IdentityCardNumber = txtIdentityNumber.Text,
                        IdentityCardPlace = txtIdentityPlace.Text,
                        Email = txtEmail.Text,
                        SocialSecurityNumber = txtSocialInsurance.Text,
                        BirthPlace = txtBirthPlace.Text,
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
                        PositionCalculationDate = dpPositionCalculationDate.SelectedDate,
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
                        SalaryHistoryLog = txtSalaryHistoryLog.Text,

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

                    context.Personnel.Add(newP);
                }
                else
                {
                    // Update
                    // Update
                    var existingP = context.Personnel.Include("LeaveHistories").FirstOrDefault(p => p.Id == _personnel.Id);
                    if (existingP != null)
                    {
                        existingP.StaffId = txtStaffId.Text;
                        existingP.FullName = txtName.Text;
                        // ... (Other fields remain same, not restating all for brevity in thought, but implementation needs them)
                        // Actually replace_file_content needs EXACT TargetContent.
                        // I will target the existingP retrieval and the end of the block.
                        
                        existingP.Gender = cboGender.Text;
                        existingP.DateOfBirth = dpDOB.SelectedDate;
                        existingP.PhoneNumber = txtPhone.Text;
                        existingP.IdentityCardNumber = txtIdentityNumber.Text;
                        existingP.IdentityCardPlace = txtIdentityPlace.Text;
                        existingP.Email = txtEmail.Text;
                        existingP.SocialSecurityNumber = txtSocialInsurance.Text;
                        existingP.BirthPlace = txtBirthPlace.Text;
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
                        existingP.PositionCalculationDate = dpPositionCalculationDate.SelectedDate;
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
                        existingP.SalaryHistoryLog = txtSalaryHistoryLog.Text;
                    
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

                        // Sync Leave History
                        if (_personnel.LeaveHistories != null)
                        {
                            // 1. Identify IDs to keep
                            var currentIds = _personnel.LeaveHistories.Select(L => L.Id).Where(id => id != 0).ToList();
                            
                            // 2. Remove deleted
                            var toDelete = existingP.LeaveHistories.Where(lh => !currentIds.Contains(lh.Id)).ToList();
                            foreach(var item in toDelete)
                            {
                                context.LeaveHistories.Remove(item);
                            }

                            // 3. Add new
                            var toAdd = _personnel.LeaveHistories.Where(lh => lh.Id == 0).ToList();
                            foreach(var item in toAdd)
                            {
                                item.PersonnelId = existingP.Id; // Link ForeignKey
                                context.LeaveHistories.Add(item); 
                            }
                            
                            // 4. Update existing
                            var toUpdate = _personnel.LeaveHistories.Where(lh => lh.Id != 0).ToList();
                            foreach(var item in toUpdate)
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
                    }
                }
                context.SaveChanges();
                // Show Custom Success Window
                var successWindow = new SuccessWindow();
                successWindow.Owner = Application.Current.MainWindow; // Set owner to center over main window
                successWindow.ShowDialog();
            }

            if (Application.Current.MainWindow is MainWindow mw)
            {
                mw.NavigateToDashboard();
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
                    // 1. Display
                    var bitmap = new System.Windows.Media.Imaging.BitmapImage(new Uri(openFileDialog.FileName));
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

        // Helper: File -> Base64 (with resizing to max 500px width/height to save space)
        private string ImageToBase64(string path)
        {
            try 
            {
                byte[] imageBytes = System.IO.File.ReadAllBytes(path);
                
                // Pure Check: If > 1MB, maybe warn? OR just resize.
                // For simplicity in WPF without extra libraries, we use built-in classes.
                
                var image = new System.Windows.Media.Imaging.BitmapImage();
                image.BeginInit();
                image.UriSource = new Uri(path);
                image.CacheOption = System.Windows.Media.Imaging.BitmapCacheOption.OnLoad;
                image.EndInit();

                // Resize logic could be complex without System.Drawing (GDI+), checking if we can just store raw bytes.
                // The user asked for "nhẹ", so let's try to limit simple storage.
                // Storing raw bytes of a 5MB photo is bad.
                // Let's implement a simple Resize using TransformedBitmap.

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

        // Helper: Base64 -> BitmapImage
        private System.Windows.Media.Imaging.BitmapImage? Base64ToImage(string base64String)
        {
            try
            {
                byte[] binaryData = Convert.FromBase64String(base64String);
                
                var bi = new System.Windows.Media.Imaging.BitmapImage();
                bi.BeginInit();
                bi.StreamSource = new System.IO.MemoryStream(binaryData);
                bi.CacheOption = System.Windows.Media.Imaging.BitmapCacheOption.OnLoad;
                bi.EndInit();
                bi.Freeze();
                return bi;
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
            txtIdentityPlace.Clear();
            txtEmail.Clear();
            txtSocialInsurance.Clear();
            txtBirthPlace.Clear();
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
            dpPositionCalculationDate.SelectedDate = null;
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
            txtSalaryHistoryLog.Clear();

            // Tab 5
            // Tab 5
            txtTotalAnnualLeave.Text = "12";
            txtAnnualLeaveTaken.Text = "0";
            txtAnnualLeaveRemaining.Text = "12";
            txtSickLeaveTaken.Text = "0";
            txtUnpaidLeaveTaken.Text = "0";
            txtMaternityLeaveTaken.Text = "0";
            
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
               foreach(var child in sp.Children)
               {
                   if (child is TextBlock tb)
                   {
                       tb.Text = "Thêm vào bảng";
                       break;
                   }
               }
            }

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

        private void NumberValidationTextBox(object? sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex("[^0-9]+");
            e.Handled = regex.IsMatch(e.Text);
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

        private void txtPhone_TextChanged(object? sender, TextChangedEventArgs e)
        {
            if (_isFormatting || _isRefreshing) return;

            _isFormatting = true;
            try
            {
                string text = txtPhone.Text;
                string digits = new string(text.Where(char.IsDigit).ToArray());

                // Giới hạn 10 số
                if (digits.Length > 10) digits = digits.Substring(0, 10);

                // Định dạng: 0XXX.XXX.XXX
                string formatted = "";
                if (digits.Length > 0)
                {
                    formatted = digits.Substring(0, Math.Min(4, digits.Length));
                    if (digits.Length > 4)
                    {
                        formatted += "." + digits.Substring(4, Math.Min(3, digits.Length - 4));
                        if (digits.Length > 7)
                        {
                            formatted += "." + digits.Substring(7, Math.Min(3, digits.Length - 7));
                        }
                    }
                }

                if (text != formatted)
                {
                    txtPhone.Text = formatted;
                    txtPhone.CaretIndex = formatted.Length;
                }

                // Hiển thị lỗi nếu không bắt đầu bằng 0
                if (digits.Length > 0 && !digits.StartsWith("0"))
                {
                    lblPhoneError.Text = "SĐT phải bắt đầu bằng số 0";
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
        private void DatePicker_PreviewKeyUp(object? sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.OriginalSource is TextBox textBox)
            {
                // Only process numeric keys to avoid duplicating slashes or messing up manual input
                bool isNumeric = (e.Key >= System.Windows.Input.Key.D0 && e.Key <= System.Windows.Input.Key.D9) ||
                                 (e.Key >= System.Windows.Input.Key.NumPad0 && e.Key <= System.Windows.Input.Key.NumPad9);

                if (isNumeric)
                {
                    string text = textBox.Text;
                    
                    // If length is exactly 2 (e.g. "12") or 5 (e.g. "12/04"), append "/" safely
                    if ((text.Length == 2 || text.Length == 5) && !text.EndsWith("/"))
                    {
                        textBox.Text = text + "/";
                        textBox.CaretIndex = textBox.Text.Length; // Move caret to end
                    }
                }
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

                // If endDate comes before startDate
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

        private void CalculateAnnualLeave(DateTime now)
        {
            // Rule: Total Annual Leave = (Years Worked / 5) + 12
            int totalLeave = 12;

            if (dpStartDate.SelectedDate.HasValue)
            {
                DateTime start = dpStartDate.SelectedDate.Value;
                
                // Calculate full years worked
                // Simple approximation or strict year difference?
                // Using strict year diff logic:
                int years = now.Year - start.Year;
                if (now < start.AddYears(years)) years--;
                
                if (years < 0) years = 0;

                int bonusDays = years / 5;
                totalLeave = 12 + bonusDays;
            }

            // Update Model and UI
            if (_personnel != null) _personnel.TotalAnnualLeaveDays = totalLeave;
            txtTotalAnnualLeave.Text = totalLeave.ToString();
            UpdateLeaveStatistics(totalLeave); // Pass calculated total explicitly
        }
        
        // ... (Skipping unrelated methods if any between) ... 

        // Important: UpdateLeaveStatistics is further down.
        // Since replace_file_content cannot skip large chunks easily without context, 
        // I will do TWO separate replaces or one big one if close.
        // They are far apart (Lines 1142 vs 1464).
        // I will split this into two calls.
        
        // Changing strategy: Only update CalculateAnnualLeave in this call.
        

        private string CalculateDetailedDateDifference(DateTime startDate, DateTime endDate)
        {
            if (startDate > endDate) return "0 năm 0 tháng 0 ngày";

            DateTime tempDate = startDate;
            int years = 0;
            int months = 0;
            int days = 0;

            // Count Years
            while (tempDate.AddYears(1) <= endDate)
            {
                years++;
                tempDate = tempDate.AddYears(1);
            }

            // Count Months
            while (tempDate.AddMonths(1) <= endDate)
            {
                months++;
                tempDate = tempDate.AddMonths(1);
            }

            // Count Days
            days = (endDate - tempDate).Days;

            return $"{years} năm {months} tháng {days} ngày";
        }
        
        // Tab 5: Leave Logic
        private void btnAddLeave_Click(object? sender, RoutedEventArgs e)
        {
            // Validate input
            if (cboLeaveType.SelectedIndex == -1)
            {
                new WarningWindow("Vui lòng chọn loại nghỉ phép!", "Thiếu thông tin").ShowDialog();
                return;
            }
            if (dpLeaveStartDate.SelectedDate == null || dpLeaveEndDate.SelectedDate == null)
            {
                new WarningWindow("Vui lòng nhập đầy đủ ngày bắt đầu và kết thúc!", "Thiếu thông tin").ShowDialog();
                return;
            }
            if (dpLeaveEndDate.SelectedDate < dpLeaveStartDate.SelectedDate)
            {
                new WarningWindow("Ngày kết thúc không thể nhỏ hơn ngày bắt đầu!", "Lỗi thời gian").ShowDialog();
                return;
            }

            var type = (cboLeaveType.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "";
            var start = dpLeaveStartDate.SelectedDate.Value;
            var end = dpLeaveEndDate.SelectedDate.Value;

            // Check for overlap
            if (_personnel != null && _personnel.LeaveHistories != null)
            {
                foreach (var history in _personnel.LeaveHistories)
                {
                    // Skip self if editing
                    if (_editingLeaveHistory != null && history == _editingLeaveHistory) continue;

                    if (start <= history.EndDate && end >= history.StartDate)
                    {
                        var msg = $"Khoảng thời gian này đã trùng với đợt nghỉ phép: {history.LeaveType} \n(Từ {history.StartDate:dd/MM/yyyy} đến {history.EndDate:dd/MM/yyyy}).\nVui lòng kiểm tra lại!";
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
                    duration = CalculateWorkingDays(start, end);
                }
                else
                {
                    duration = (end - start).TotalDays + 1;
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

                // Reset Edit Mode
                _editingLeaveHistory = null;
                dgLeaveHistory.SelectedItem = null;
                
                // Reset Button Text
                // Find TextBlock inside button
                if (btnAddLeave.Content is StackPanel sp)
                {
                   foreach(var child in sp.Children)
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
                                LeaveYear = previousYear
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
                                LeaveYear = currentYear
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
                    LeaveYear = (cboLeaveType.SelectedItem as ComboBoxItem)?.Content?.ToString() == "Phép năm" ? (int?)cboLeaveYear.SelectedItem : null
                };
                
                if (_personnel == null) _personnel = new Personnel();
                if (_personnel.LeaveHistories == null) _personnel.LeaveHistories = new System.Collections.Generic.List<LeaveHistory>();
                
                _personnel.LeaveHistories.Add(newItem);
            }
            
            // Refresh Grid
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
        }

        private void OnLeaveDateChanged(object? sender, SelectionChangedEventArgs e)
        {
            CalculateMaternityEndDate();
            CalculateMaternityEndDate();
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

        private void CalculateMaternityEndDate()
        {
            if (_isRefreshing) return;

            var type = (cboLeaveType.SelectedItem as ComboBoxItem)?.Content?.ToString();
            if (type != "Thai sản") return;

            if (dpLeaveStartDate.SelectedDate == null) return;

            DateTime start = dpLeaveStartDate.SelectedDate.Value;
            DateTime cutoff = new DateTime(2026, 7, 1);
            DateTime end;

            if (start < cutoff)
            {
                end = start.AddMonths(6).AddDays(-1);
            }
            else
            {
                end = start.AddMonths(7).AddDays(-1);
            }

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
                return;
            }

            var start = dpLeaveStartDate.SelectedDate.Value;
            var end = dpLeaveEndDate.SelectedDate.Value;

            if (end < start)
            {
                txtLeaveDuration.Text = "";
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
        }

        private double CalculateWorkingDays(DateTime start, DateTime end)
        {
            double days = 0;
            for (var date = start; date <= end; date = date.AddDays(1))
            {
                if (date.DayOfWeek != DayOfWeek.Saturday && date.DayOfWeek != DayOfWeek.Sunday)
                {
                    days++;
                }
            }
            return days;
        }

        private LeaveHistory? _editingLeaveHistory = null;

        private void dgLeaveHistory_SelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (dgLeaveHistory.SelectedItem is LeaveHistory item)
            {
                _editingLeaveHistory = item;
                
                // Populate fields
                foreach(ComboBoxItem cbi in cboLeaveType.Items)
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

                // Change Button Text
                // Find TextBlock inside button
                if (btnAddLeave.Content is StackPanel sp)
                {
                   foreach(var child in sp.Children)
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

            double annualTaken = 0;
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

                // Logic: Count if LeaveYear matches, OR if LeaveYear is null and StartDate.Year matches
                int itemYear = item.LeaveYear ?? item.StartDate.Year;

                if (itemYear == yearToCalculate)
                {
                    if (item.LeaveType == "Phép năm") annualTaken += item.DurationDays;
                    else if (item.LeaveType == "Nghỉ ốm") sickTaken += item.DurationDays;
                    // We still track year-specific unpaid leave if needed, but the UI should show the total for salary delay accuracy
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
            txtAnnualLeaveTaken.Text = annualTaken.ToString();
            txtAnnualLeaveRemaining.Text = (totalAnnual - annualTaken).ToString();
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

            // Sort by StartDate
            var sortedList = _personnel.LeaveHistories.OrderBy(x => x.StartDate).ToList();

            // Set ShowDeleteButton logic
            var processedLinkIds = new HashSet<string>();
            foreach (var item in sortedList)
            {
                if (!string.IsNullOrEmpty(item.LinkId))
                {
                    if (processedLinkIds.Contains(item.LinkId))
                    {
                        item.ShowDeleteButton = false; // Hide if link seen before
                        item.IsLinkedGroupHead = false;
                    }
                    else
                    {
                        item.ShowDeleteButton = true; // Show for first instance
                        // Check if there are actually other records with this LinkId
                        int count = sortedList.Count(x => x.LinkId == item.LinkId);
                        item.IsLinkedGroupHead = count > 1; // It is a head of a group only if count > 1
                        
                        processedLinkIds.Add(item.LinkId);
                    }
                }
                else
                {
                    item.ShowDeleteButton = true; // Always show for unlinked
                    item.IsLinkedGroupHead = false;
                }
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
            if (double.TryParse(txtUnpaidLeaveTaken.Text, out double unpaidDays) && unpaidDays > 0)
            {
                expectedDate = expectedDate.AddDays(unpaidDays);
            }

            // Set Result
            if (dpExpectedSalaryIncrease.SelectedDate != expectedDate)
            {
                dpExpectedSalaryIncrease.SelectedDate = expectedDate;
            }
        }
    }
}
