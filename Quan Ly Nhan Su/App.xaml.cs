using System.Windows;
using TaxPersonnelManagement.Data;
using TaxPersonnelManagement.Views;
using Microsoft.EntityFrameworkCore;
using AutoUpdaterDotNET;
using System;
using System.IO;
using System.Linq;
using System.Globalization;
using TaxPersonnelManagement.Models;





namespace TaxPersonnelManagement
{
    public partial class App : Application
    {
        public static Models.User? CurrentUser { get; set; }

        public App()
        {
            DebugLog("App Constructor Started");
            this.DispatcherUnhandledException += App_DispatcherUnhandledException;
            try 
            {
                InitializeComponent(); 
            }
            catch (Exception ex)
            {
                DebugLog("Error in Constructor: " + ex.ToString());
            }
        }

        private void App_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            DebugLog($"Unhandled Exception: {e.Exception.Message}\n{e.Exception.StackTrace}");
            MessageBox.Show($"Lỗi không mong muốn: {e.Exception.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            e.Handled = true;
        }

        private void Application_Startup(object sender, StartupEventArgs e)
        {
            // Cấu hình Tự động cập nhật (Auto-Update)
            // Đăng ký sự kiện để tự thiết kế giao diện thông báo
            AutoUpdater.CheckForUpdateEvent += (args) =>
            {
                if (args.Error == null)
                {
                    if (args.IsUpdateAvailable)
                    {
                        var updateWindow = new UpdateNotificationWindow(args);
                        updateWindow.ShowDialog();
                    }
                }
            };

            // Thêm ?t=... để tránh bị lưu bộ nhớ đệm (Cache) của GitHub
            AutoUpdater.Start($"https://raw.githubusercontent.com/htcomputercantho-afk/Quan-Ly-Nhan-Su-Thue/main/update.xml?t={System.DateTime.Now.Ticks}");
            
            // Force Standard Vietnamese Culture (Ignore Local OS Overrides)
            var culture = new System.Globalization.CultureInfo("vi-VN", false);
            culture.DateTimeFormat.ShortDatePattern = "dd/MM/yyyy";
            culture.DateTimeFormat.DateSeparator = "/";
            
            System.Threading.Thread.CurrentThread.CurrentCulture = culture;
            System.Threading.Thread.CurrentThread.CurrentUICulture = culture;
            FrameworkElement.LanguageProperty.OverrideMetadata(
                typeof(FrameworkElement),
                new FrameworkPropertyMetadata(
                    System.Windows.Markup.XmlLanguage.GetLanguage(culture.IetfLanguageTag)));

            // Workaround for Unikey/WPF bug where Vietnamese typing gets disabled
            EventManager.RegisterClassHandler(typeof(System.Windows.Controls.TextBox), UIElement.GotKeyboardFocusEvent, new System.Windows.Input.KeyboardFocusChangedEventHandler((s, args) =>
            {
                if (s is System.Windows.Controls.TextBox tb)
                {
                    System.Windows.Input.InputMethod.SetIsInputMethodEnabled(tb, true);
                    System.Windows.Input.InputMethod.SetIsInputMethodSuspended(tb, false);
                }
            }));
            
            EventManager.RegisterClassHandler(typeof(System.Windows.Controls.PasswordBox), UIElement.GotKeyboardFocusEvent, new System.Windows.Input.KeyboardFocusChangedEventHandler((s, args) =>
            {
                if (s is System.Windows.Controls.PasswordBox pb)
                {
                    System.Windows.Input.InputMethod.SetIsInputMethodEnabled(pb, true);
                    System.Windows.Input.InputMethod.SetIsInputMethodSuspended(pb, false);
                }
            }));

            DebugLog("Application_Startup Fired");
            try
            {
                // Initialize Database
                using (var context = new AppDbContext())
                {
                    DebugLog("Ensuring Database Created...");
                    context.Database.EnsureCreated();

                    // Detect if migration is pending and perform auto-backup
                    if (IsSchemaUpdatePending(context))
                    {
                        PerformBackup("before_migration");
                    }

                    
                    // Manual Migration for Departments table
                    context.Database.ExecuteSqlRaw(@"
                        CREATE TABLE IF NOT EXISTS Departments (
                            Id INTEGER PRIMARY KEY AUTOINCREMENT,
                            Name TEXT NOT NULL
                        );");

                    // Manual Migration for Positions table
                    context.Database.ExecuteSqlRaw(@"
                        CREATE TABLE IF NOT EXISTS Positions (
                            Id INTEGER PRIMARY KEY AUTOINCREMENT,
                            Name TEXT NOT NULL
                        );");

                    // Manual Migration for Ranks table
                    context.Database.ExecuteSqlRaw(@"
                        CREATE TABLE IF NOT EXISTS Ranks (
                            Id INTEGER PRIMARY KEY AUTOINCREMENT,
                            Code TEXT NOT NULL,
                            Name TEXT NOT NULL
                        );");

                    // Manual Migration for AvatarBase64 (Try-Catch to ignore if exists)
                    try
                    {
                        context.Database.ExecuteSqlRaw("ALTER TABLE Personnel ADD COLUMN AvatarBase64 TEXT");
                    }
                    catch { /* Column likely exists */ }

                    // Manual Migration for Tab 2 Fields (Position History)
                    try { context.Database.ExecuteSqlRaw("ALTER TABLE Personnel ADD COLUMN PositionDecisionDate TEXT"); } catch { }
                    try { context.Database.ExecuteSqlRaw("ALTER TABLE Personnel ADD COLUMN PositionCalculationDate TEXT"); } catch { }
                    try { context.Database.ExecuteSqlRaw("ALTER TABLE Personnel ADD COLUMN PositionYear TEXT"); } catch { }
                    try { context.Database.ExecuteSqlRaw("ALTER TABLE Personnel ADD COLUMN DetailedWorkHistory TEXT"); } catch { }
                    
                    // Manual Migration for Tab 3 Fields (Retirement Info)
                    try { context.Database.ExecuteSqlRaw("ALTER TABLE Personnel ADD COLUMN RetirementDate TEXT"); } catch { }

                    // Manual Migration for Tab 4 Fields (Party Info)
                    try { context.Database.ExecuteSqlRaw("ALTER TABLE Personnel ADD COLUMN PartyEntryDate TEXT"); } catch { }
                    try { context.Database.ExecuteSqlRaw("ALTER TABLE Personnel ADD COLUMN PartyOfficialDate TEXT"); } catch { }

                    // Manual Migration for Tab 5 Fields (Leave Info)
                    try { context.Database.ExecuteSqlRaw("ALTER TABLE Personnel ADD COLUMN TotalAnnualLeaveDays INTEGER DEFAULT 12"); } catch { }
                    context.Database.ExecuteSqlRaw(@"
                        CREATE TABLE IF NOT EXISTS LeaveHistories (
                            Id INTEGER PRIMARY KEY AUTOINCREMENT,
                            PersonnelId INTEGER NOT NULL,
                            LeaveType TEXT NOT NULL,
                            StartDate TEXT NOT NULL,
                            EndDate TEXT NOT NULL,
                            DurationDays REAL NOT NULL,
                            Reason TEXT,
                            SystemNote TEXT,
                            FOREIGN KEY (PersonnelId) REFERENCES Personnel(Id) ON DELETE CASCADE

                        );");

                    // Manual Migration for LeaveYear (Feature: Year Selection)
                    try { context.Database.ExecuteSqlRaw("ALTER TABLE LeaveHistories ADD COLUMN LeaveYear INTEGER"); } catch { }
                    try { context.Database.ExecuteSqlRaw("ALTER TABLE LeaveHistories ADD COLUMN SystemNote TEXT"); } catch { }

                    // Manual Migration for Tab 6 Fields (Salary Info)
                    try { context.Database.ExecuteSqlRaw("ALTER TABLE Personnel ADD COLUMN CurrentSalaryStep TEXT"); } catch { }
                    try { context.Database.ExecuteSqlRaw("ALTER TABLE Personnel ADD COLUMN CurrentSalaryCoefficient REAL DEFAULT 0"); } catch { }
                    try { context.Database.ExecuteSqlRaw("ALTER TABLE Personnel ADD COLUMN ExceedFramePercent REAL DEFAULT 0"); } catch { }
                    try { context.Database.ExecuteSqlRaw("ALTER TABLE Personnel ADD COLUMN PositionAllowance TEXT"); } catch { }
                    try { context.Database.ExecuteSqlRaw("ALTER TABLE Personnel ADD COLUMN SalaryReservationDeadline TEXT"); } catch { }
                    try { context.Database.ExecuteSqlRaw("ALTER TABLE Personnel ADD COLUMN NextSalaryStepDate TEXT"); } catch { }
                    try { context.Database.ExecuteSqlRaw("ALTER TABLE Personnel ADD COLUMN SalaryIncreaseDelayType TEXT"); } catch { }
                    try { context.Database.ExecuteSqlRaw("ALTER TABLE Personnel ADD COLUMN ExpectedSalaryIncreaseDate TEXT"); } catch { }
                    try { context.Database.ExecuteSqlRaw("ALTER TABLE Personnel ADD COLUMN SalaryHistoryLog TEXT"); } catch { }

                    // Manual Migration for Tab 7 Fields (Reward Info)
                    try { context.Database.ExecuteSqlRaw("ALTER TABLE Personnel ADD COLUMN EmulationTitles TEXT"); } catch { }
                    try { context.Database.ExecuteSqlRaw("ALTER TABLE Personnel ADD COLUMN RewardForms TEXT"); } catch { }

                    // Manual Migration for Tab 8 Fields (Discipline Info)
                    try { context.Database.ExecuteSqlRaw("ALTER TABLE Personnel ADD COLUMN DisciplineType TEXT"); } catch { }
                    try { context.Database.ExecuteSqlRaw("ALTER TABLE Personnel ADD COLUMN DisciplineDecisionNumber TEXT"); } catch { }
                    try { context.Database.ExecuteSqlRaw("ALTER TABLE Personnel ADD COLUMN DisciplineDecisionDate TEXT"); } catch { }
                    try { context.Database.ExecuteSqlRaw("ALTER TABLE Personnel ADD COLUMN DisciplineReason TEXT"); } catch { }

                    // Manual Migration for Feature: Salary Config Popup
                    context.Database.ExecuteSqlRaw(@"
                        CREATE TABLE IF NOT EXISTS RankSalarySpecs (
                            Id INTEGER PRIMARY KEY AUTOINCREMENT,
                            RankCode TEXT NOT NULL,
                            SalaryStep TEXT NOT NULL,
                            Coefficient REAL NOT NULL
                        );");

                    // Manual Migration for Salary Delay Reasons
                    context.Database.ExecuteSqlRaw(@"
                        CREATE TABLE IF NOT EXISTS SalaryDelayReasons (
                            Id INTEGER PRIMARY KEY AUTOINCREMENT,
                            Name TEXT NOT NULL
                        );");

                    // Manual Migration for IncomeRecords
                    context.Database.ExecuteSqlRaw(@"
                        CREATE TABLE IF NOT EXISTS IncomeRecords (
                            Id INTEGER PRIMARY KEY AUTOINCREMENT,
                            PersonnelId INTEGER NOT NULL,
                            Year INTEGER NOT NULL,
                            Month INTEGER NOT NULL,
                            IncomeType TEXT NOT NULL,
                            Amount TEXT NOT NULL,
                            Note TEXT,
                            FOREIGN KEY (PersonnelId) REFERENCES Personnel(Id) ON DELETE CASCADE
                        );");


                    // Manual Migration for PublicHolidays
                    context.Database.ExecuteSqlRaw(@"
                        CREATE TABLE IF NOT EXISTS PublicHolidays (
                            Id INTEGER PRIMARY KEY AUTOINCREMENT,
                            Name TEXT NOT NULL,
                            Date TEXT NOT NULL,
                            Note TEXT
                        );");

                    // Ensure holidays for current and next year
                    EnsureHolidaysSeeded(context);



                    // Seed Salary Delay Reasons if empty
                    var delayCount = context.Database.ExecuteSqlRaw("SELECT COUNT(*) FROM SalaryDelayReasons"); // This returns -1 for SELECT usually with ExecuteSqlRaw, need scalar
                    // Actually, ExecuteSqlRaw returns interactions.
                    // Let's just try insert ignoring conflicts or check via EF.
                    
                    // Better verify if empty using EF since schema is now ensured
                    if (!context.SalaryDelayReasons.Any())
                    {
                         context.Database.ExecuteSqlRaw("INSERT INTO SalaryDelayReasons (Id, Name) VALUES (1, 'Lùi 3 tháng (Khiển trách)')");
                         context.Database.ExecuteSqlRaw("INSERT INTO SalaryDelayReasons (Id, Name) VALUES (2, 'Lùi 6 tháng (Cảnh cáo)')");
                         context.Database.ExecuteSqlRaw("INSERT INTO SalaryDelayReasons (Id, Name) VALUES (3, 'Lùi 12 tháng (Giáng chức/Cách chức)')");
                         context.Database.ExecuteSqlRaw("INSERT INTO SalaryDelayReasons (Id, Name) VALUES (4, 'Nghỉ không lương')");
                    }

                    // Manual Migration for Nullable Dates (Schema Fix)
                    bool needDateFix = false;
                    using (var command = context.Database.GetDbConnection().CreateCommand())
                    {
                         command.CommandText = "PRAGMA table_info(Personnel)";
                         context.Database.OpenConnection();
                         using (var result = command.ExecuteReader())
                         {
                             while (result.Read())
                             {
                                 if (result["name"].ToString() == "DateOfBirth" && Convert.ToInt32(result["notnull"]) == 1)
                                 {
                                     needDateFix = true;
                                     break;
                                 }
                             }
                         }
                         context.Database.CloseConnection();
                    }

                    // Manual Migration for Nullable EndDate in LeaveHistories
                    bool needLeaveDateFix = false;
                    using (var command = context.Database.GetDbConnection().CreateCommand())
                    {
                         command.CommandText = "PRAGMA table_info(LeaveHistories)";
                         context.Database.OpenConnection();
                         using (var result = command.ExecuteReader())
                         {
                             while (result.Read())
                             {
                                 if (result["name"].ToString() == "EndDate" && Convert.ToInt32(result["notnull"]) == 1)
                                 {
                                     needLeaveDateFix = true;
                                     break;
                                 }
                             }
                         }
                         context.Database.CloseConnection();
                    }

                    if (needLeaveDateFix)
                    {
                        DebugLog("Migrating LeaveHistories table to allow nullable EndDate...");
                        using (var transaction = context.Database.BeginTransaction())
                        {
                            try
                            {
                                context.Database.ExecuteSqlRaw("ALTER TABLE LeaveHistories RENAME TO LeaveHistories_Old");
                                context.Database.ExecuteSqlRaw(@"
                                    CREATE TABLE LeaveHistories (
                                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                                        PersonnelId INTEGER NOT NULL,
                                        LeaveType TEXT NOT NULL,
                                        StartDate TEXT NOT NULL,
                                        EndDate TEXT, -- NULLABLE NOW
                                        DurationDays REAL NOT NULL,
                                        Reason TEXT,
                                        LeaveYear INTEGER,
                                        SystemNote TEXT,
                                        FOREIGN KEY (PersonnelId) REFERENCES Personnel(Id) ON DELETE CASCADE
                                    );");
                                context.Database.ExecuteSqlRaw(@"
                                    INSERT INTO LeaveHistories (Id, PersonnelId, LeaveType, StartDate, EndDate, DurationDays, Reason, LeaveYear, SystemNote)
                                    SELECT Id, PersonnelId, LeaveType, StartDate, EndDate, DurationDays, Reason, LeaveYear, SystemNote FROM LeaveHistories_Old");
                                context.Database.ExecuteSqlRaw("DROP TABLE LeaveHistories_Old");
                                transaction.Commit();
                                DebugLog("LeaveHistories Migration Successful.");
                            }
                            catch (Exception ex)
                            {
                                transaction.Rollback();
                                DebugLog("LeaveHistories Migration Failed: " + ex.Message);
                                throw;
                            }
                        }
                    }

                    if (needDateFix)

                    {
                        DebugLog("Migrating Personnel table to allow nullable dates...");
                        using (var transaction = context.Database.BeginTransaction())
                        {
                            try
                            {
                                // 1. Rename old table
                                context.Database.ExecuteSqlRaw("ALTER TABLE Personnel RENAME TO Personnel_Old");

                                // 2. Create new table (Schema from EF Core but manually defined to ensure correctness)
                                // We use a broad schema that matches current Entity
                                context.Database.ExecuteSqlRaw(@"
                                    CREATE TABLE Personnel (
                                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                                        StaffId TEXT,
                                        FullName TEXT NOT NULL,
                                        DateOfBirth TEXT, -- Now Nullable
                                        Gender TEXT,
                                        PhoneNumber TEXT,
                                        IdentityCardNumber TEXT,
                                        IdentityCardPlace TEXT,
                                        SocialSecurityNumber TEXT,
                                        Email TEXT,
                                        BirthPlace TEXT,
                                        Ethnicity TEXT,
                                        Religion TEXT,
                                        Department TEXT,
                                        Position TEXT,
                                        RankCode TEXT,
                                        RankName TEXT,
                                        TaxAuthorityStartDate TEXT, -- Now Nullable
                                        StartDate TEXT, -- Now Nullable
                                        Status TEXT,
                                        EducationLevel TEXT,
                                        Major TEXT,
                                        University TEXT,
                                        StateManagementLevel TEXT,
                                        PoliticalTheoryLevel TEXT,
                                        ITSkillLevel TEXT,
                                        LanguageSkillLevel TEXT,
                                        AvatarBase64 TEXT,
                                        TotalAnnualLeaveDays INTEGER DEFAULT 12,
                                        CurrentSalaryStep TEXT,
                                        CurrentSalaryCoefficient REAL DEFAULT 0,
                                        ExceedFramePercent REAL DEFAULT 0,
                                        PositionAllowance TEXT,
                                        SalaryReservationDeadline TEXT,
                                        NextSalaryStepDate TEXT,
                                        SalaryIncreaseDelayType TEXT,
                                        ExpectedSalaryIncreaseDate TEXT,
                                        SalaryHistoryLog TEXT,
                                        EmulationTitles TEXT,
                                        RewardForms TEXT,
                                        DisciplineType TEXT,
                                        DisciplineDecisionNumber TEXT,
                                        DisciplineDecisionDate TEXT,
                                        DisciplineReason TEXT,
                                        PositionDecisionDate TEXT,
                                        PositionCalculationDate TEXT,
                                        PositionYear TEXT,
                                        DetailedWorkHistory TEXT,
                                        RetirementDate TEXT,
                                        PartyEntryDate TEXT,
                                        PartyOfficialDate TEXT
                                    )");

                                // 3. Copy Data
                                context.Database.ExecuteSqlRaw(@"
                                    INSERT INTO Personnel (
                                        Id, StaffId, FullName, DateOfBirth, Gender, PhoneNumber, IdentityCardNumber, IdentityCardPlace,
                                        SocialSecurityNumber, Email, BirthPlace, Ethnicity, Religion, Department, Position,
                                        RankCode, RankName, TaxAuthorityStartDate, StartDate, Status, EducationLevel, Major,
                                        University, StateManagementLevel, PoliticalTheoryLevel, ITSkillLevel, LanguageSkillLevel, AvatarBase64,
                                        TotalAnnualLeaveDays, CurrentSalaryStep, CurrentSalaryCoefficient, ExceedFramePercent,
                                        PositionAllowance, SalaryReservationDeadline, NextSalaryStepDate, SalaryIncreaseDelayType,
                                        ExpectedSalaryIncreaseDate, SalaryHistoryLog, EmulationTitles, RewardForms,
                                        DisciplineType, DisciplineDecisionNumber, DisciplineDecisionDate, DisciplineReason,
                                        PositionDecisionDate, PositionCalculationDate, PositionYear, DetailedWorkHistory,
                                        RetirementDate, PartyEntryDate, PartyOfficialDate
                                    )
                                    SELECT 
                                        Id, StaffId, FullName, DateOfBirth, Gender, PhoneNumber, IdentityCardNumber, IdentityCardPlace,
                                        SocialSecurityNumber, Email, BirthPlace, Ethnicity, Religion, Department, Position,
                                        RankCode, RankName, TaxAuthorityStartDate, StartDate, Status, EducationLevel, Major,
                                        University, StateManagementLevel, PoliticalTheoryLevel, ITSkillLevel, LanguageSkillLevel, AvatarBase64,
                                        TotalAnnualLeaveDays, CurrentSalaryStep, CurrentSalaryCoefficient, ExceedFramePercent,
                                        PositionAllowance, SalaryReservationDeadline, NextSalaryStepDate, SalaryIncreaseDelayType,
                                        ExpectedSalaryIncreaseDate, SalaryHistoryLog, EmulationTitles, RewardForms,
                                        DisciplineType, DisciplineDecisionNumber, DisciplineDecisionDate, DisciplineReason,
                                        PositionDecisionDate, PositionCalculationDate, PositionYear, DetailedWorkHistory,
                                        RetirementDate, PartyEntryDate, PartyOfficialDate
                                    FROM Personnel_Old");

                                // 4. Drop Old Table
                                context.Database.ExecuteSqlRaw("DROP TABLE Personnel_Old");

                                transaction.Commit();
                                DebugLog("Migration Successful.");
                            }
                            catch (Exception ex)
                            {
                                transaction.Rollback();
                                DebugLog("Migration Failed: " + ex.Message);
                                throw; 
                            }
                        }
                    }

                    DebugLog("Database Created/Updated.");
                }

                // Show Login View
                DebugLog("Showing LoginView...");
                var loginView = new LoginView();
                loginView.Show();
                DebugLog("LoginView Shown.");
            }
            catch (System.Exception ex)
            {
                DebugLog($"Error: {ex.Message}\n{ex.StackTrace}");
                MessageBox.Show($"Lỗi khởi động: {ex.Message}\n{ex.StackTrace}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                Shutdown();
            }
        }

        public static void DebugLog(string message)
        {
            try
            {
                System.IO.File.AppendAllText("debug.log", $"{System.DateTime.Now}: {message}\n");
            }
            catch { }
        }

        /// <summary>
        /// Thực hiện sao lưu cơ sở dữ liệu vào thư mục backup.
        /// </summary>
        /// <param name="reason">Lý do sao lưu để gắn vào tên file.</param>
        public static void PerformBackup(string reason = "manual")
        {
            try
            {
                string dbPath = Path.Combine(System.AppContext.BaseDirectory, "tax_personnel.db");
                if (!File.Exists(dbPath))
                {
                    DebugLog($"Backup skipped: DB file not found at {dbPath}");
                    return;
                }

                string backupDir = Path.Combine(System.AppContext.BaseDirectory, "backup");
                if (!Directory.Exists(backupDir))
                {
                    Directory.CreateDirectory(backupDir);
                }

                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string backupPath = Path.Combine(backupDir, $"tax_personnel_{reason}_{timestamp}.db");

                File.Copy(dbPath, backupPath, true);
                DebugLog($"Backup successful: {backupPath} (Reason: {reason})");

                // Tự động dọn dẹp: Chỉ giữ lại 5 bản sao lưu gần nhất
                try
                {
                    int maxBackups = 5;
                    var directory = new DirectoryInfo(backupDir);
                    var files = directory.GetFiles("tax_personnel_*.db")
                                         .OrderByDescending(f => f.CreationTime)
                                         .ToList();

                    if (files.Count > maxBackups)
                    {
                        for (int i = maxBackups; i < files.Count; i++)
                        {
                            files[i].Delete();
                            DebugLog($"Auto-cleanup: Deleted old backup {files[i].Name}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    DebugLog($"Auto-cleanup failed: {ex.Message}");
                }

            }
            catch (Exception ex)
            {
                DebugLog($"Backup failed: {ex.Message}");
            }
        }

        private bool IsSchemaUpdatePending(AppDbContext context)
        {
            try
            {
                using (var command = context.Database.GetDbConnection().CreateCommand())
                {
                    // 1. Kiểm tra cột EndDate trong LeaveHistories (đã đổi thành nullable chưa?)
                    command.CommandText = "PRAGMA table_info(LeaveHistories)";
                    context.Database.OpenConnection();
                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            if (reader["name"].ToString() == "EndDate" && Convert.ToInt32(reader["notnull"]) == 1)
                                return true;
                        }
                    }
                    context.Database.CloseConnection();

                    // 2. Kiểm tra xem bảng IncomeRecords đã tồn tại chưa
                    command.CommandText = "SELECT name FROM sqlite_master WHERE type='table' AND name='IncomeRecords'";
                    context.Database.OpenConnection();
                    var result = command.ExecuteScalar();
                    context.Database.CloseConnection();
                    if (result == null) return true;
                }
            }
            catch (Exception ex)
            {
                DebugLog($"CheckSchemaUpdate failed: {ex.Message}");
            }
            return false;
        }

        private void EnsureHolidaysSeeded(AppDbContext context)
        {
            int currentYear = DateTime.Now.Year;
            SeedHolidaysForYear(context, currentYear - 1); // Ensure last year
            SeedHolidaysForYear(context, currentYear);
            SeedHolidaysForYear(context, currentYear + 1);
        }

        private void SeedHolidaysForYear(AppDbContext context, int year)
        {
            // Check if already seeded for this year
            if (context.PublicHolidays.Any(h => h.Date.Year == year)) return;

            var holidays = new System.Collections.Generic.List<PublicHoliday>();
            
            // Solar Fixed
            holidays.Add(new PublicHoliday { Name = "Tết Dương lịch", Date = new DateTime(year, 1, 1) });
            holidays.Add(new PublicHoliday { Name = "Ngày Giải phóng", Date = new DateTime(year, 4, 30) });
            holidays.Add(new PublicHoliday { Name = "Ngày Quốc tế Lao động", Date = new DateTime(year, 5, 1) });
            holidays.Add(new PublicHoliday { Name = "Ngày Quốc khánh", Date = new DateTime(year, 9, 2) });
            holidays.Add(new PublicHoliday { Name = "Ngày Quốc khánh (bổ sung)", Date = new DateTime(year, 9, 1) });

            // Lunar Holidays using ChineseLunisolarCalendar
            ChineseLunisolarCalendar lunar = new ChineseLunisolarCalendar();

            // 1. Giỗ tổ Hùng Vương (10/03 Âm lịch)
            try {
                DateTime hungKings = GetSolarFromLunar(lunar, year, 3, 10);
                holidays.Add(new PublicHoliday { Name = "Giỗ tổ Hùng Vương", Date = hungKings });
            } catch { }

            // 2. Tết Nguyên Đán (Từ 29/12 hoặc 30/12 đến mùng 4 Tết)
            try {
                // Mùng 1, 2, 3, 4 Tết
                holidays.Add(new PublicHoliday { Name = "Tết Nguyên Đán", Date = GetSolarFromLunar(lunar, year, 1, 1) });
                holidays.Add(new PublicHoliday { Name = "Tết Nguyên Đán", Date = GetSolarFromLunar(lunar, year, 1, 2) });
                holidays.Add(new PublicHoliday { Name = "Tết Nguyên Đán", Date = GetSolarFromLunar(lunar, year, 1, 3) });
                holidays.Add(new PublicHoliday { Name = "Tết Nguyên Đán", Date = GetSolarFromLunar(lunar, year, 1, 4) });

                // Đêm Giao thừa (Ngày cuối cùng của năm trước)
                // Phải tính dựa trên năm hiện tại trừ 1 để ra ngày cuối cùng
                DateTime firstDayOfNewYear = GetSolarFromLunar(lunar, year, 1, 1);
                holidays.Add(new PublicHoliday { Name = "Tết Nguyên Đán (Giao thừa)", Date = firstDayOfNewYear.AddDays(-1) });
            } catch { }

            context.PublicHolidays.AddRange(holidays);
            context.SaveChanges();
            DebugLog($"Auto-seeded holidays for year {year}.");
        }

        private DateTime GetSolarFromLunar(ChineseLunisolarCalendar lunar, int year, int lMonth, int lDay)
        {
            // ChineseLunisolarCalendar year cycle is different, we need to find the solar year
            // that contains this lunar date.
            // For a given solar year, the Lunar New Year usually falls in Jan or Feb.
            
            // Try to find the solar date in the range [year, year+1]
            for (int y = year - 1; y <= year + 1; y++)
            {
                try {
                    int monthsInYear = lunar.GetMonthsInYear(y);
                    for (int m = 1; m <= monthsInYear; m++)
                    {
                        // Handle leap months: we usually only care about the non-leap month for holidays
                        // but GetMonth returns 1..13 if there is a leap month.
                        // Simplified check:
                        if (lunar.GetDayOfMonth(lunar.ToDateTime(y, m, lDay, 0, 0, 0, 0)) == lDay)
                        {
                           // This is complex. Let's use a simpler approach.
                        }
                    }
                } catch { }
            }

            // Simplified accurate enough for standard years:
            // The ToDateTime(year, month, day, ...) uses Lunar parameters.
            // We need to map the "Solar Year" to the "Lunar Year" which is tricky.
            // Actually, for Tet of solar year Y, the Lunar Year is usually Y or Y-1.
            
            // Standard approach for Tet in Solar Year Y:
            // Tet usually starts in Jan/Feb of Year Y.
            // So we try Lunar Year Y or Y-1.
            
            DateTime date1 = lunar.ToDateTime(year, lMonth, lDay, 0, 0, 0, 0);
            if (date1.Year == year || (lMonth == 1 && date1.Year == year)) return date1;
            
            DateTime date2 = lunar.ToDateTime(year - 1, lMonth, lDay, 0, 0, 0, 0);
            return date2;
        }
    }
}



