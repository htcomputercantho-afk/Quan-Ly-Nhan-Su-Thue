using System;



using System.Collections.Generic;



using System.Linq;



using System.Windows;



using System.Windows.Controls;



using System.Windows.Media;



using System.Windows.Navigation;



using Microsoft.EntityFrameworkCore;



using TaxPersonnelManagement.Data;



using TaxPersonnelManagement.Models;



using TaxPersonnelManagement.Services;



using TaxPersonnelManagement.Views;







namespace TaxPersonnelManagement



{



    public partial class MainWindow : Window



    {



        private User? _currentUser;



        private PersonnelDetailView? _personnelDetailCache;



        private DashboardView? _dashboardCache;



        private StatisticsView? _statisticsCache;



        private PlanningManagementView? _planningCache;



        private RotationManagementView? _rotationCache;



        private DutyAssignmentView? _dutyAssignmentCache;



        private SalaryListView? _salaryCache;



        private AnnualIncomeView? _annualIncomeCache;



        private LeaveDetailView? _leaveDetailCache;



        private PositionDurationView? _positionDurationCache;



        private EmulationRewardView? _emulationRewardCache;



        private EvaluationListView? _evaluationCache;



        private TrainingListView? _trainingCache;



        private UserListView? _userListCache;



        private BackupRestoreView? _backupRestoreCache;



        private RoleManagementView? _roleManagementCache;



        private System.Windows.Controls.Button? _activeMenuButton;



        private object? _currentView;











        /// <summary>



        /// Khởi tạo cửa sổ chính sau khi đăng nhập thành công.



        /// </summary>



        /// <param name="user">Thông tin người dùng hiện tại.</param>



        public MainWindow(User user)



        {



            App.DebugLog("MainWindow Constructor Entry");



            InitializeComponent();



            UpdateWindowStateUI();



            _currentUser = user;



            App.CurrentUser = user; // Lưu thông tin người dùng vào biến toàn cục của Ứng dụng



            txtWelcome.Text = _currentUser.FullName; // Hiển thị tên người dùng trên thanh tiêu đề







            // Áp dụng phân quyền chi tiết (RBAC) cho thanh menu điều hướng



            ApplyUserPermissions();







            // Điều hướng mặc định tới màn hình Tổng quan (Dashboard) khi mở app



            NavigateDashboard(null, null);



            SetVersionInfo();







            // Auto-collapse sidebar khi app khởi động trên màn hình nhỏ



            this.Loaded += (s, e) =>



            {



                if (this.ActualWidth < 1400)



                    CollapseSidebar();







                // Đợi MainWindow hiển thị và vẽ hoàn tất trên màn hình rồi mới mở các popup cảnh báo / đồng bộ



                Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Background, new Action(() =>



                {



                    // Kiểm tra và hiển thị cảnh báo công chức sắp đi làm lại



                    CheckLeaveReturnAlerts();







                    // Kiểm tra và hiển thị cảnh báo đến hạn (nâng lương / nghỉ hưu / bổ nhiệm)



                    CheckDeadlineAlerts();







                    // Tự động kiểm tra bản cập nhật CSDL trên Google Drive khi mở app



                    _ = CheckGoogleDriveStartupSyncAsync();



                    PrewarmViewsInBackground();



                }));



            };



        }







        private bool _isClosingHandled = false;



        private bool _isLoggingOut = false;







        /// <summary>



        /// Tự động kiểm tra xem trên Google Drive có bản sao lưu CSDL mới hơn máy hiện tại không khi mở app.



        /// Nếu có, hỏi người dùng có muốn tải về để đồng bộ không.



        /// </summary>



        private async System.Threading.Tasks.Task CheckGoogleDriveStartupSyncAsync()



        {



            try



            {



                if (!App.DriveSync.HasSavedToken()) return;







                if (!App.DriveSync.IsConnected)



                {



                    bool connected = await App.DriveSync.TryLoadSavedTokenAsync();



                    if (!connected) return;



                }







                var cloudInfo = await App.DriveSync.GetCloudFileInfoAsync();



                if (cloudInfo == null || !cloudInfo.ModifiedTime.HasValue) return;







                App.DriveSync.InitialCloudModifiedTime = cloudInfo.ModifiedTime;







                if (App.DriveSync.ShouldPromptStartupDownload(cloudInfo, out string reason))



                {



                    var localTime = App.DriveSync.GetLocalDbLastWriteTime();



                    string message = $"Phát hiện bản sao lưu trên Google Drive MỚI HƠN dữ liệu trên máy này:\n\n" +



                                     $"• Trên Google Drive: {cloudInfo.ModifiedTime.Value:dd/MM/yyyy HH:mm:ss}\n" +



                                     $"• Trên máy tính này: {localTime?.ToString("dd/MM/yyyy HH:mm:ss") ?? "Không rõ"}\n\n" +



                                     $"Bạn có muốn tải dữ liệu mới nhất từ Google Drive về máy không?";







                    var confirmWin = new ConfirmWindow(message, "Đồng Bộ Dữ Liệu Từ Google Drive");



                    confirmWin.Owner = this;



                    if (confirmWin.ShowDialog() == true)



                    {



                        var syncDialog = new SyncOnCloseWindow();



                        syncDialog.Owner = this;



                        syncDialog.Show();







                        try



                        {



                            bool pullSuccess = await App.DriveSync.PullAsync();



                            syncDialog.Close();







                            if (pullSuccess)



                            {



                                App.IsDataDirty = false;



                                // Làm mới lại giao diện hiển thị và xóa toàn bộ các view cache



                                _dashboardCache = null;



                                _personnelDetailCache = null;



                                _statisticsCache = null;



                                _planningCache = null;



                                _rotationCache = null;



                                _dutyAssignmentCache = null;



                                InvalidateDataCaches();



                                NavigateDashboard(null, null);







                                var successWin = new SuccessWindow("Đã tải và cập nhật dữ liệu mới nhất từ Google Drive thành công!", "Đồng Bộ Thành Công");



                                successWin.Owner = this;



                                successWin.ShowDialog();



                            }



                            else



                            {



                                var warnWin = new WarningWindow("Không thể tải dữ liệu từ Google Drive. Vui lòng thử lại sau!", "Lỗi Đồng Bộ");



                                warnWin.Owner = this;



                                warnWin.ShowDialog();



                            }



                        }



                        catch (Exception ex)



                        {



                            syncDialog.Close();



                            App.DebugLog($"Pull on startup error: {ex.Message}");



                        }



                    }



                }



                else



                {



                    App.DebugLog($"Startup sync check: Skipped download prompt ({reason}).");



                }



            }



            catch (Exception ex)



            {



                App.DebugLog($"CheckGoogleDriveStartupSyncAsync error: {ex.Message}");



            }



        }







        /// <summary>



        /// Kiểm tra xem có công chức nào sắp hết nghỉ (thai sản / ốm / phép) trong 7 ngày tới không.



        /// Nếu có, hiển thị popup cảnh báo để nhắc nhở.



        /// </summary>



        private void CheckLeaveReturnAlerts()



        {



            try



            {



                var today = DateTime.Today;



                var cutoff = today.AddDays(7);







                using (var db = new AppDbContext())



                {



                    var allPersonnel = db.Personnel



                        .Include(p => p.LeaveHistories)



                        .Where(p => string.IsNullOrEmpty(p.Status) || p.Status == "Đang công tác")



                        .ToList();







                    var alerts = new List<LeaveReturnInfo>();







                    foreach (var p in allPersonnel)



                    {



                        if (p.LeaveHistories == null || p.LeaveHistories.Count == 0) continue;







                        // Sắp hết kỳ nghỉ trong 7 ngày tới



                        var upcomingLeave = p.LeaveHistories



                            .Where(l =>



                                l.EndDate.HasValue &&



                                l.StartDate.Date <= today &&



                                l.EndDate.Value.Date >= today &&



                                l.EndDate.Value.Date <= cutoff)



                            .OrderBy(l => l.EndDate)



                            .FirstOrDefault();







                        if (upcomingLeave != null)



                        {



                            var returnDate = upcomingLeave.EndDate!.Value.Date.AddDays(1);



                            alerts.Add(new LeaveReturnInfo



                            {



                                AlertType  = "upcoming",



                                FullName   = p.FullName,



                                Department = p.Department,



                                LeaveType  = upcomingLeave.LeaveType,



                                ReturnDate = returnDate,



                                DaysLeft   = (returnDate - today).Days



                            });



                            continue;



                        }







                        // Đợt nghỉ chưa có ngày kết thúc (thai sản, ốm dài ngày, không lương, dưỡng sức)



                        var openLeave = p.LeaveHistories



                            .Where(l =>



                                !l.EndDate.HasValue &&



                                l.StartDate.Date <= today &&



                                (l.LeaveType.Contains("Thai sản") ||



                                 l.LeaveType.Contains("Nghỉ ốm") ||



                                 l.LeaveType.Contains("Nghỉ thai sản") ||



                                 l.LeaveType.Contains("Không lương") ||



                                 l.LeaveType.Contains("Dưỡng sức")))



                            .OrderByDescending(l => l.StartDate)



                            .FirstOrDefault();







                        if (openLeave != null)



                        {



                            alerts.Add(new LeaveReturnInfo



                            {



                                AlertType  = "open",



                                FullName   = p.FullName,



                                Department = p.Department,



                                LeaveType  = openLeave.LeaveType,



                                ReturnDate = null,



                                DaysLeft   = 0



                            });



                        }



                    }







                    if (alerts.Count > 0)



                    {



                        var alertWindow = new LeaveReturnAlertWindow(alerts);



                        alertWindow.Owner = this;



                        alertWindow.ShowDialog();



                    }



                }



            }



            catch (Exception ex)



            {



                App.DebugLog("CheckLeaveReturnAlerts error: " + ex.Message);



            }



        }







        private void CheckDeadlineAlerts()



        {



            try



            {



                var today = DateTime.Today;



                const int SalaryAlertDays      = 60;



                const int RetirementAlertDays  = 90;



                const int AppointmentAlertDays = 90;



                const int AllowanceAlertDays   = 30;



                const int AppointmentTermYears = 5;







                using (var db = new AppDbContext())



                {



                    var personnelList = db.Personnel



                        .Where(p => string.IsNullOrEmpty(p.Status) || p.Status == "Đang công tác")



                        .ToList();







                    var alerts = new List<DeadlineAlertInfo>();







                    foreach (var p in personnelList)



                    {



                        // 1. Thời hạn bảo lưu phụ cấp chức vụ



                        if (p.SalaryReservationDeadline.HasValue)



                        {



                            var allowanceDate = p.SalaryReservationDeadline.Value.Date;



                            int days = (allowanceDate - today).Days;



                            if (days >= 0 && days <= AllowanceAlertDays)



                            {



                                alerts.Add(new DeadlineAlertInfo



                                {



                                    AlertCategory = "allowance",



                                    FullName      = p.FullName,



                                    Department    = p.Department,



                                    DeadlineDate  = allowanceDate,



                                    DaysLeft      = days



                                });



                            }



                        }







                        // 2. Nâng lương định kỳ



                        if (p.ExpectedSalaryIncreaseDate.HasValue)



                        {



                            var salaryDate = p.ExpectedSalaryIncreaseDate.Value.Date;



                            int days = (salaryDate - today).Days;



                            if (days >= 0 && days <= SalaryAlertDays)



                            {



                                alerts.Add(new DeadlineAlertInfo



                                {



                                    AlertCategory = "salary",



                                    FullName      = p.FullName,



                                    Department    = p.Department,



                                    DeadlineDate  = salaryDate,



                                    DaysLeft      = days



                                });



                            }



                        }







                        // 3. Nghỉ hưu



                        if (p.RetirementDate.HasValue)



                        {



                            var retDate = p.RetirementDate.Value.Date;



                            int days = (retDate - today).Days;



                            if (days >= 0 && days <= RetirementAlertDays)



                            {



                                alerts.Add(new DeadlineAlertInfo



                                {



                                    AlertCategory = "retirement",



                                    FullName      = p.FullName,



                                    Department    = p.Department,



                                    DeadlineDate  = retDate,



                                    DaysLeft      = days



                                });



                            }



                        }







                        // 4. Bổ nhiệm / tái cử (Trưởng / Phó)



                        if (p.PositionDecisionDate.HasValue && !string.IsNullOrWhiteSpace(p.Position))



                        {



                            string pos = p.Position.ToLower();



                            bool isLeader = (pos.Contains("trưởng") || pos.Contains("phó"))



                                         && !pos.Contains("công chức");



                            if (isLeader)



                            {



                                var apptEndDate = p.PositionDecisionDate.Value



                                                    .AddYears(AppointmentTermYears).Date;



                                int days = (apptEndDate - today).Days;



                                if (days >= 0 && days <= AppointmentAlertDays)



                                {



                                    alerts.Add(new DeadlineAlertInfo



                                    {



                                        AlertCategory = "appointment",



                                        FullName      = p.FullName,



                                        Department    = p.Department,



                                        DeadlineDate  = apptEndDate,



                                        DaysLeft      = days



                                    });



                                }



                            }



                        }



                    }







                    if (alerts.Count > 0)



                    {



                        var alertWin = new DeadlineAlertWindow(alerts);



                        alertWin.Owner = this;



                        alertWin.ShowDialog();



                    }



                }



            }



            catch (Exception ex)



            {



                App.DebugLog("CheckDeadlineAlerts error: " + ex.Message);



            }



        }







        /// <summary>



        /// Bắt sự kiện người dùng bấm nút [X] đóng ứng dụng.



        /// Cơ chế Smart Cloud Sync:



        /// - Nếu không có dữ liệu nào bị sửa đổi (IsDataDirty == false): Đóng app ngay, không đẩy đè CSDL.



        /// - Nếu có dữ liệu sửa đổi (IsDataDirty == true): Kiểm tra xung đột với Drive và đẩy lên an toàn.



        /// </summary>



        protected override async void OnClosing(System.ComponentModel.CancelEventArgs e)



        {



            if (!_isClosingHandled && App.DriveSync.HasSavedToken())



            {



                // 1. Nếu người dùng chỉ mở app xem, không sửa bất kỳ dữ liệu nào -> BỎ QUA PUSH để tránh ghi đè dữ liệu mới trên Drive



                if (!App.IsDataDirty)



                {



                    App.DebugLog("OnClosing: IsDataDirty is false. Skipping auto-push to Google Drive to preserve cloud data.");



                    base.OnClosing(e);



                    return;



                }







                // 2. Có thay đổi dữ liệu -> Tạm dừng đóng để xử lý đồng bộ



                e.Cancel = true;



                _isClosingHandled = true;







                try



                {



                    if (!App.DriveSync.IsConnected)



                    {



                        await App.DriveSync.TryLoadSavedTokenAsync();



                    }







                    if (App.DriveSync.IsConnected)



                    {



                        // Kiểm tra xung đột: Drive có bị ai khác sửa trong lúc phiên làm việc này đang mở không?



                        var cloudInfo = await App.DriveSync.GetCloudFileInfoAsync();



                        var initialCloudTime = App.DriveSync.InitialCloudModifiedTime;







                        bool hasConflict = false;



                        if (cloudInfo?.ModifiedTime.HasValue == true && initialCloudTime.HasValue)



                        {



                            // Nếu Drive bị sửa đổi trong lúc phiên này đang mở (mới hơn thời điểm mở app ít nhất 10 giây)



                            if (cloudInfo.ModifiedTime.Value > initialCloudTime.Value.AddSeconds(10))



                            {



                                hasConflict = true;



                            }



                        }



                        else if (cloudInfo?.ModifiedTime.HasValue == true && !initialCloudTime.HasValue)



                        {



                            // Fallback nếu lúc mở app chưa lấy được cloudInfo: so sánh với SessionStartTime



                            if (cloudInfo.ModifiedTime.Value > App.SessionStartTime.AddMinutes(2))



                            {



                                hasConflict = true;



                            }



                        }







                        if (hasConflict)



                        {



                            string message = $"CẢNH BÁO XUNG ĐỘT DỮ LIỆU!\n\n" +



                                             $"Dữ liệu trên Google Drive đã được cập nhật từ máy khác vào lúc: {cloudInfo?.ModifiedTime:dd/MM/yyyy HH:mm:ss}.\n\n" +



                                             $"Nếu tiếp tục đẩy lên, bạn sẽ GHI ĐÈ và làm mất dữ liệu đó.\n\n" +



                                             $"Bạn có chắc chắn muốn đẩy dữ liệu từ máy này lên Google Drive không?";







                            var confirm = new ConfirmWindow(message, "Cảnh Báo Ghi Đè Dữ Liệu");



                            confirm.Owner = this;



                            if (confirm.ShowDialog() != true)



                            {



                                // Người dùng hủy -> Không đẩy lên, đóng app



                                this.Close();



                                return;



                            }



                        }







                        // Tiến hành Push



                        var syncDialog = new SyncOnCloseWindow();



                        syncDialog.Owner = this;



                        syncDialog.Show();







                        try



                        {



                            App.DebugLog("OnClosing: Starting PushAsync...");



                            var pushTask = App.DriveSync.PushAsync();



                            var completedTask = await System.Threading.Tasks.Task.WhenAny(pushTask, System.Threading.Tasks.Task.Delay(60000));



                            if (completedTask == pushTask)



                            {



                                bool pushSuccess = await pushTask;



                                App.DebugLog($"OnClosing: PushAsync finished. Result = {pushSuccess}");



                            }



                            else



                            {



                                App.DebugLog("OnClosing: PushAsync TIMED OUT after 60s.");



                            }



                        }



                        finally



                        {



                            syncDialog.Close();



                        }



                    }



                }



                catch (Exception ex)



                {



                    App.DebugLog($"OnClosing sync error: {ex.Message}");



                }



                finally



                {



                    this.Close(); // Tiếp tục đóng ứng dụng



                }



                return;



            }







            base.OnClosing(e);



        }







        private void SetVersionInfo()



        {



            var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;



            txtVersion.Text = $"Version {version?.Major}.{version?.Minor}.{version?.Build}.{version?.Revision}";



        }







        // Constructor trống dành cho hỗ trợ Visual Designer (không dùng trong runtime)



        public MainWindow()



        {



            InitializeComponent();



            UpdateWindowStateUI();



        }







        /// <summary>



        /// Hiển thị một View trên khung giao diện chính với cơ chế chuyển đổi Visibility (0ms).



        /// Toàn bộ Visual Tree của các view đã xem được giữ nguyên trong MainContentGrid,



        /// giúp loại bỏ hoàn toàn độ trễ khi chuyển đổi giữa các menu có dữ liệu lớn.



        /// </summary>



        private void ShowView(object? view)



        {



            if (view == null || MainContentGrid == null) return;



            _currentView = view;







            FrameworkElement? targetElement = null;







            foreach (UIElement child in MainContentGrid.Children)



            {



                if (child is FrameworkElement fe && fe.Tag == view)



                {



                    targetElement = fe;



                    break;



                }



            }







            if (targetElement == null)



            {



                if (view is Page page)



                {



                    var frame = new Frame



                    {



                        Content = page,



                        NavigationUIVisibility = NavigationUIVisibility.Hidden,



                        JournalOwnership = JournalOwnership.UsesParentJournal,



                        Tag = view



                    };



                    targetElement = frame;



                }



                else if (view is FrameworkElement fe)



                {



                    fe.Tag = view;



                    targetElement = fe;



                }



                else



                {



                    targetElement = new ContentControl { Content = view, Tag = view };



                }







                MainContentGrid.Children.Add(targetElement);



            }







            // Hiện view được chọn, ẩn toàn bộ các view khác tức thì (0ms)



            foreach (UIElement child in MainContentGrid.Children)



            {



                if (child is FrameworkElement fe)



                {



                    fe.Visibility = (fe == targetElement) ? Visibility.Visible : Visibility.Collapsed;



                }



            }



        }







        private void EnsureViewHosted(object? view)



        {



            if (view == null || MainContentGrid == null) return;







            foreach (UIElement child in MainContentGrid.Children)



            {



                if (child is FrameworkElement fe && fe.Tag == view)



                {



                    return; // Đã có trong visual tree



                }



            }







            FrameworkElement targetElement;



            if (view is Page page)



            {



                targetElement = new Frame



                {



                    Content = page,



                    NavigationUIVisibility = NavigationUIVisibility.Hidden,



                    JournalOwnership = JournalOwnership.UsesParentJournal,



                    Tag = view,



                    Visibility = Visibility.Collapsed



                };



            }



            else if (view is FrameworkElement fe)



            {



                fe.Tag = view;



                fe.Visibility = Visibility.Collapsed;



                targetElement = fe;



            }



            else



            {



                targetElement = new ContentControl { Content = view, Tag = view, Visibility = Visibility.Collapsed };



            }







            MainContentGrid.Children.Add(targetElement);



        }







        private void RemoveViewHost(object? view)



        {



            if (view == null || MainContentGrid == null) return;



            for (int i = MainContentGrid.Children.Count - 1; i >= 0; i--)



            {



                if (MainContentGrid.Children[i] is FrameworkElement fe && fe.Tag == view)



                {



                    MainContentGrid.Children.RemoveAt(i);



                    break;



                }



            }



        }







        /// <summary>



        /// Điều hướng tới màn hình Tổng quan



        /// </summary>



        public void NavigateToDashboard(int? targetPersonnelId = null)



        {



            if (_activeMenuButton == btnDashboard && !targetPersonnelId.HasValue && _currentView == _dashboardCache)



                return;







            UpdateMenuState(btnDashboard); // Cập nhật trạng thái hiển thị của nút menu







            if (_dashboardCache == null)



            {



                _dashboardCache = new DashboardView(targetPersonnelId);



            }



            else if (targetPersonnelId.HasValue)



            {



                _dashboardCache.PersonnelList.TargetPersonnelId = targetPersonnelId;



                _dashboardCache.PersonnelList.LoadData();



            }







            ShowView(_dashboardCache);



        }







        private void NavigateDashboard(object? sender, RoutedEventArgs? e)



        {



            NavigateToDashboard();



        }







        private void NavigateStatistics(object sender, RoutedEventArgs e)



        {



            if (_activeMenuButton == btnStatistics && _currentView == _statisticsCache) return;



            UpdateMenuState(btnStatistics);



            _statisticsCache ??= new StatisticsView();



            ShowView(_statisticsCache);



        }







        /// <summary>



        /// Điều hướng tới màn hình Chi tiết hồ sơ cán bộ.



        /// </summary>



        /// <param name="p">Hồ sơ cán bộ cần xem/sửa (null nếu muốn thêm mới).</param>



        public void NavigateToPersonnelDetail(Personnel? p, int activeTab = 0)



        {



            UpdateMenuState(btnPersonnel);







            // Nếu là thêm mới, sử dụng cache nếu có



            if (p == null)



            {



                if (_personnelDetailCache == null)



                {



                    _personnelDetailCache = new PersonnelDetailView(null, activeTab);



                }



                ShowView(_personnelDetailCache);



            }



            else



            {



                // Nếu là chỉnh sửa nhân sự cụ thể, gỡ host cũ và tạo view mới



                RemoveViewHost(_personnelDetailCache);



                _personnelDetailCache = new PersonnelDetailView(p, activeTab);



                ShowView(_personnelDetailCache);



            }



        }







        public void ClearPersonnelCache()



        {



            RemoveViewHost(_personnelDetailCache);



            _personnelDetailCache = null;



            InvalidateDataCaches();



        }







        public void InvalidateDataCaches()



        {



            MainContentGrid?.Children.Clear();



            _currentView = null;



            _dashboardCache = null;



            _statisticsCache = null;



            _personnelDetailCache = null;



            _salaryCache = null;



            _annualIncomeCache = null;



            _leaveDetailCache = null;



            _positionDurationCache = null;



            _emulationRewardCache = null;



            _evaluationCache = null;



            _trainingCache = null;



            _planningCache = null;



            _rotationCache = null;



            _dutyAssignmentCache = null;



            _userListCache = null;



            _backupRestoreCache = null;



            _roleManagementCache = null;



        }







        private void NavigatePersonnel(object sender, RoutedEventArgs e)



        {



            NavigateToPersonnelDetail(null);



            // Ghi chú: Dùng màn hình Tổng quan để xem danh sách nhân sự.



        }







        private void NavigateSalary(object sender, RoutedEventArgs e)



        {



            if (_activeMenuButton == btnSalary && _currentView == _salaryCache) return;



            UpdateMenuState(btnSalary);



            _salaryCache ??= new SalaryListView();



            ShowView(_salaryCache);



        }







        private void NavigateAnnualIncome(object sender, RoutedEventArgs e)



        {



            if (_activeMenuButton == btnAnnualIncome && _currentView == _annualIncomeCache) return;



            UpdateMenuState(btnAnnualIncome);



            _annualIncomeCache ??= new AnnualIncomeView();



            ShowView(_annualIncomeCache);



        }







        private void NavigateLeaveDetail(object sender, RoutedEventArgs e)



        {



            if (_activeMenuButton == btnLeaveDetail && _currentView == _leaveDetailCache) return;



            UpdateMenuState(btnLeaveDetail);



            _leaveDetailCache ??= new LeaveDetailView();



            ShowView(_leaveDetailCache);



        }







        private void NavigateEmulationReward(object sender, RoutedEventArgs e)



        {



            if (_activeMenuButton == btnEmulationReward && _currentView == _emulationRewardCache) return;



            UpdateMenuState(btnEmulationReward);



            _emulationRewardCache ??= new EmulationRewardView();



            ShowView(_emulationRewardCache);



        }







        private void NavigateEvaluation(object sender, RoutedEventArgs e)



        {



            if (_activeMenuButton == btnEvaluation && _currentView == _evaluationCache) return;



            UpdateMenuState(btnEvaluation);



            _evaluationCache ??= new EvaluationListView();



            ShowView(_evaluationCache);



        }







        private void NavigateTraining(object sender, RoutedEventArgs e)



        {



            if (_activeMenuButton == btnTraining && _currentView == _trainingCache) return;



            UpdateMenuState(btnTraining);



            _trainingCache ??= new TrainingListView();



            ShowView(_trainingCache);



        }







        private void NavigatePlanning(object sender, RoutedEventArgs e)



        {



            if (_activeMenuButton == btnPlanning && _currentView == _planningCache) return;



            UpdateMenuState(btnPlanning);



            _planningCache ??= new PlanningManagementView();



            ShowView(_planningCache);



        }







        private void NavigateRotation(object sender, RoutedEventArgs e)



        {



            if (_activeMenuButton == btnRotation && _currentView == _rotationCache) return;



            UpdateMenuState(btnRotation);



            _rotationCache ??= new RotationManagementView();



            ShowView(_rotationCache);



        }







        private void NavigateDutyAssignment(object sender, RoutedEventArgs e)



        {



            if (_activeMenuButton == btnDutyAssignment && _currentView == _dutyAssignmentCache) return;



            UpdateMenuState(btnDutyAssignment);



            _dutyAssignmentCache ??= new DutyAssignmentView();



            ShowView(_dutyAssignmentCache);



        }







        private void NavigatePositionDuration(object sender, RoutedEventArgs e)



        {



            if (_activeMenuButton == btnPositionDuration && _currentView == _positionDurationCache) return;



            UpdateMenuState(btnPositionDuration);



            _positionDurationCache ??= new PositionDurationView();



            ShowView(_positionDurationCache);



        }







        private void NavigateRoles(object sender, RoutedEventArgs e)



        {



            if (!PermissionService.HasPermission(PermissionRegistry.RoleManagement_Manage))



            {



                var w = new WarningWindow("Từ chối truy cập", "Tài khoản của bạn không được phân quyền truy cập chức năng này.\n\nVui lòng liên hệ Quản trị viên để được cấp quyền."); w.Owner = this; w.ShowDialog();



                return;



            }



            if (_activeMenuButton == btnRoles && _currentView == _roleManagementCache) return;



            UpdateMenuState(btnRoles);



            _roleManagementCache ??= new RoleManagementView();



            ShowView(_roleManagementCache);



        }







        private void NavigateUsers(object sender, RoutedEventArgs e)



        {



            if (!PermissionService.HasPermission(PermissionRegistry.UserManagement_Manage))



            {



                var w = new WarningWindow("Từ chối truy cập", "Tài khoản của bạn không được phân quyền truy cập chức năng này.\n\nVui lòng liên hệ Quản trị viên để được cấp quyền."); w.Owner = this; w.ShowDialog();



                return;



            }



            if (_activeMenuButton == btnUsers && _currentView == _userListCache) return;



            UpdateMenuState(btnUsers);



            _userListCache ??= new UserListView();



            ShowView(_userListCache);



        }







        private void NavigateBackupRestore(object sender, RoutedEventArgs e)



        {



            if (!PermissionService.HasPermission(PermissionRegistry.BackupRestore_Manage))



            {



                var w = new WarningWindow("Từ chối truy cập", "Tài khoản của bạn không được phân quyền truy cập chức năng này.\n\nVui lòng liên hệ Quản trị viên để được cấp quyền."); w.Owner = this; w.ShowDialog();



                return;



            }



            if (_activeMenuButton == btnBackupRestore && _currentView == _backupRestoreCache) return;



            UpdateMenuState(btnBackupRestore);



            _backupRestoreCache ??= new BackupRestoreView();



            ShowView(_backupRestoreCache);



        }







        public void ApplyUserPermissions()



        {



            PermissionService.ReloadCurrentUserPermissions();







            btnDashboard.Visibility = PermissionService.HasPermission(PermissionRegistry.Menu_Dashboard) ? Visibility.Visible : Visibility.Collapsed;



            btnStatistics.Visibility = PermissionService.HasPermission(PermissionRegistry.Menu_Statistics) ? Visibility.Visible : Visibility.Collapsed;



            btnPersonnel.Visibility = PermissionService.HasPermission(PermissionRegistry.Menu_Personnel) ? Visibility.Visible : Visibility.Collapsed;



            btnSalary.Visibility = PermissionService.HasPermission(PermissionRegistry.Menu_Salary) ? Visibility.Visible : Visibility.Collapsed;



            btnAnnualIncome.Visibility = PermissionService.HasPermission(PermissionRegistry.Menu_AnnualIncome) ? Visibility.Visible : Visibility.Collapsed;



            btnLeaveDetail.Visibility = PermissionService.HasPermission(PermissionRegistry.Menu_LeaveDetail) ? Visibility.Visible : Visibility.Collapsed;



            btnPositionDuration.Visibility = PermissionService.HasPermission(PermissionRegistry.Menu_PositionDuration) ? Visibility.Visible : Visibility.Collapsed;



            btnEmulationReward.Visibility = PermissionService.HasPermission(PermissionRegistry.Menu_EmulationReward) ? Visibility.Visible : Visibility.Collapsed;



            btnEvaluation.Visibility = PermissionService.HasPermission(PermissionRegistry.Menu_Evaluation) ? Visibility.Visible : Visibility.Collapsed;



            btnTraining.Visibility = PermissionService.HasPermission(PermissionRegistry.Menu_Training) ? Visibility.Visible : Visibility.Collapsed;



            btnPlanning.Visibility = PermissionService.HasPermission(PermissionRegistry.Menu_Planning) ? Visibility.Visible : Visibility.Collapsed;



            btnRotation.Visibility = PermissionService.HasPermission(PermissionRegistry.Menu_Rotation) ? Visibility.Visible : Visibility.Collapsed;



            btnDutyAssignment.Visibility = PermissionService.HasPermission(PermissionRegistry.Menu_DutyAssignment) ? Visibility.Visible : Visibility.Collapsed;



            btnRoles.Visibility = PermissionService.HasPermission(PermissionRegistry.RoleManagement_Manage) ? Visibility.Visible : Visibility.Collapsed;



            btnUsers.Visibility = PermissionService.HasPermission(PermissionRegistry.UserManagement_Manage) ? Visibility.Visible : Visibility.Collapsed;



            btnBackupRestore.Visibility = PermissionService.HasPermission(PermissionRegistry.BackupRestore_Manage) ? Visibility.Visible : Visibility.Collapsed;



        }







        private void UpdateMenuState(System.Windows.Controls.Button activeButton)



        {



            if (_activeMenuButton == activeButton) return;



            _activeMenuButton = activeButton;







            // Đặt lại nền tất cả các nút menu về trong suốt



            var transparent = System.Windows.Media.Brushes.Transparent;



            btnDashboard.Background = transparent;



            btnStatistics.Background = transparent;



            btnPersonnel.Background = transparent;



            btnSalary.Background = transparent;



            btnAnnualIncome.Background = transparent;



            btnLeaveDetail.Background = transparent;



            btnPositionDuration.Background = transparent;



            btnEmulationReward.Background = transparent;



            btnEvaluation.Background = transparent;



            btnTraining.Background = transparent;



            btnPlanning.Background = transparent;



            btnRotation.Background = transparent;



            btnDutyAssignment.Background = transparent;



            btnRoles.Background = transparent;



            btnUsers.Background = transparent;



            btnBackupRestore.Background = transparent;







            // Đánh dấu nút đang active với nền trắng bán trong suốt (~30% opacity)



            activeButton.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromArgb(80, 255, 255, 255));



        }







        /// <summary>



        /// Khởi tạo trước (pre-warm) các màn hình menu trong thời gian rảnh của ứng dụng (ApplicationIdle)



        /// và nạp sẵn vào visual tree của MainContentGrid để khi click lần đầu tiên cũng chuyển tức thì (0ms).



        /// </summary>



        private void PrewarmViewsInBackground()



        {



            Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.ApplicationIdle, new Action(() =>



            {



                try { _statisticsCache ??= new StatisticsView(); EnsureViewHosted(_statisticsCache); } catch { }



            }));



            Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.ApplicationIdle, new Action(() =>



            {



                try { _salaryCache ??= new SalaryListView(); EnsureViewHosted(_salaryCache); } catch { }



            }));



            Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.ApplicationIdle, new Action(() =>



            {



                try { _leaveDetailCache ??= new LeaveDetailView(); EnsureViewHosted(_leaveDetailCache); } catch { }



            }));



            Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.ApplicationIdle, new Action(() =>



            {



                try { _positionDurationCache ??= new PositionDurationView(); EnsureViewHosted(_positionDurationCache); } catch { }



            }));



            Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.ApplicationIdle, new Action(() =>



            {



                try { _annualIncomeCache ??= new AnnualIncomeView(); EnsureViewHosted(_annualIncomeCache); } catch { }



            }));



            Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.ApplicationIdle, new Action(() =>



            {



                try { _emulationRewardCache ??= new EmulationRewardView(); EnsureViewHosted(_emulationRewardCache); } catch { }



            }));



            Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.ApplicationIdle, new Action(() =>



            {



                try { _evaluationCache ??= new EvaluationListView(); EnsureViewHosted(_evaluationCache); } catch { }



            }));



            Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.ApplicationIdle, new Action(() =>



            {



                try { _trainingCache ??= new TrainingListView(); EnsureViewHosted(_trainingCache); } catch { }



            }));



            Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.ApplicationIdle, new Action(() =>



            {



                try { _planningCache ??= new PlanningManagementView(); EnsureViewHosted(_planningCache); } catch { }



            }));



            Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.ApplicationIdle, new Action(() =>



            {



                try { _rotationCache ??= new RotationManagementView(); EnsureViewHosted(_rotationCache); } catch { }



            }));



            Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.ApplicationIdle, new Action(() =>



            {



                try { _dutyAssignmentCache ??= new DutyAssignmentView(); EnsureViewHosted(_dutyAssignmentCache); } catch { }



            }));



        }







        private bool _isSidebarExpanded = true;



        /// <summary>



        /// True nếu người dùng đã bấm toggle thủ công — khi đó SizeChanged không tự ghi đè trạng thái.



        /// </summary>



        private bool _manualToggle = false;







        private void CollapseSidebar()



        {



            _isSidebarExpanded = false;



            colSidebar.Width = new GridLength(70);



            txtLogo.Visibility = Visibility.Collapsed;



            txtOverview.Visibility = Visibility.Collapsed;



            txtStatistics.Visibility = Visibility.Collapsed;



            txtPersonnel.Visibility = Visibility.Collapsed;



            txtSalary.Visibility = Visibility.Collapsed;



            txtAnnualIncome.Visibility = Visibility.Collapsed;



            txtLeaveDetail.Visibility = Visibility.Collapsed;



            txtPositionDuration.Visibility = Visibility.Collapsed;



            txtEmulationReward.Visibility = Visibility.Collapsed;



            txtEvaluation.Visibility = Visibility.Collapsed;



            txtTraining.Visibility = Visibility.Collapsed;



            txtPlanning.Visibility = Visibility.Collapsed;



            txtRotation.Visibility = Visibility.Collapsed;



            txtDutyAssignment.Visibility = Visibility.Collapsed;



            txtUsers.Visibility = Visibility.Collapsed;



            txtBackupRestore.Visibility = Visibility.Collapsed;



            txtLogout.Visibility = Visibility.Collapsed;



            txtCopyright.Visibility = Visibility.Collapsed;



            txtVersion.Visibility = Visibility.Collapsed;



            imgLogo.Margin = new Thickness(0);







            var buttons = new[] { btnDashboard, btnStatistics, btnPersonnel, btnSalary, btnAnnualIncome, btnLeaveDetail, btnPositionDuration, btnEmulationReward, btnEvaluation, btnTraining, btnPlanning, btnRotation, btnDutyAssignment, btnUsers, btnBackupRestore, btnLogout };



            foreach (var btn in buttons)



            {



                btn.Padding = new Thickness(0);



                if (btn.Content is System.Windows.Controls.StackPanel sp && sp.Children.Count > 0)



                {



                    var icon = sp.Children[0] as FrameworkElement;



                    if (icon != null) icon.Margin = new Thickness(14, 0, 0, 0);



                }



            }



        }







        private void ExpandSidebar()



        {



            _isSidebarExpanded = true;



            colSidebar.Width = new GridLength(250);



            txtLogo.Visibility = Visibility.Visible;



            txtOverview.Visibility = Visibility.Visible;



            txtStatistics.Visibility = Visibility.Visible;



            txtPersonnel.Visibility = Visibility.Visible;



            txtSalary.Visibility = Visibility.Visible;



            txtAnnualIncome.Visibility = Visibility.Visible;



            txtLeaveDetail.Visibility = Visibility.Visible;



            txtPositionDuration.Visibility = Visibility.Visible;



            txtEmulationReward.Visibility = Visibility.Visible;



            txtEvaluation.Visibility = Visibility.Visible;



            txtTraining.Visibility = Visibility.Visible;



            txtPlanning.Visibility = Visibility.Visible;



            txtRotation.Visibility = Visibility.Visible;



            txtDutyAssignment.Visibility = Visibility.Visible;



            txtUsers.Visibility = Visibility.Visible;



            txtBackupRestore.Visibility = Visibility.Visible;



            txtLogout.Visibility = Visibility.Visible;



            txtCopyright.Visibility = Visibility.Visible;



            txtVersion.Visibility = Visibility.Visible;



            imgLogo.Margin = new Thickness(0, 0, 10, 0);







            var buttons = new[] { btnDashboard, btnStatistics, btnPersonnel, btnSalary, btnAnnualIncome, btnLeaveDetail, btnPositionDuration, btnEmulationReward, btnEvaluation, btnTraining, btnPlanning, btnRotation, btnDutyAssignment, btnUsers, btnBackupRestore, btnLogout };



            foreach (var btn in buttons)



            {



                btn.Padding = new Thickness(25, 0, 25, 0);



                if (btn.Content is System.Windows.Controls.StackPanel sp && sp.Children.Count > 0)



                {



                    var icon = sp.Children[0] as FrameworkElement;



                    if (icon != null) icon.Margin = new Thickness(0, 0, 15, 0);



                }



            }



        }







        private void btnToggleSidebar_Click(object sender, RoutedEventArgs e)



        {



            _manualToggle = true; // Người dùng chủ động bấm — không tự ghi đè nữa



            if (_isSidebarExpanded) CollapseSidebar();



            else ExpandSidebar();



        }







        /// <summary>



        /// Tự động thu gọn/mở rộng Sidebar khi cửa sổ thay đổi kích thước.



        /// Ngưỡng: &lt; 1400px → thu gọn; ≥ 1400px → mở rộng.



        /// Nếu người dùng đã bấm toggle thủ công thì không tự ghi đè.



        /// </summary>



        private void MainWindow_SizeChanged(object sender, SizeChangedEventArgs e)



        {



            if (_manualToggle) return;







            if (e.NewSize.Width < 1400 && _isSidebarExpanded)



                CollapseSidebar();



            else if (e.NewSize.Width >= 1400 && !_isSidebarExpanded)



                ExpandSidebar();



        }







        #region Window Caption & Title Bar Controls







        private void btnMinimize_Click(object sender, RoutedEventArgs e)



        {



            SystemCommands.MinimizeWindow(this);



        }







        private void btnMaximize_Click(object sender, RoutedEventArgs e)



        {



            if (this.WindowState == WindowState.Maximized)



            {



                SystemCommands.RestoreWindow(this);



            }



            else



            {



                SystemCommands.MaximizeWindow(this);



            }



        }







        private void btnClose_Click(object sender, RoutedEventArgs e)



        {



            SystemCommands.CloseWindow(this);



        }







        private void MainWindow_StateChanged(object? sender, EventArgs e)



        {



            UpdateWindowStateUI();



        }







        private void UpdateWindowStateUI()



        {



            if (this.WindowState == WindowState.Maximized)



            {



                // Bù trừ 7px viền mở rộng của Windows khi cửa sổ Maximized để tránh tràn màn hình



                RootWindowGrid.Margin = new Thickness(7);



                if (pathMaximize != null)



                {



                    pathMaximize.Data = Geometry.Parse("M 2.5,2.5 V 0.5 H 9.5 V 7.5 H 7.5 M 0.5,2.5 H 7.5 V 9.5 H 0.5 Z");



                    btnMaximize.ToolTip = "Khôi phục kích thước";



                }



            }



            else



            {



                RootWindowGrid.Margin = new Thickness(0);



                if (pathMaximize != null)



                {



                    pathMaximize.Data = Geometry.Parse("M 0.5,0.5 H 9.5 V 9.5 H 0.5 Z");



                    btnMaximize.ToolTip = "Phóng to";



                }



            }



        }







        #endregion







        private void Logout_Click(object sender, RoutedEventArgs e)



        {



            _isLoggingOut = true;



            Application.Current.ShutdownMode = ShutdownMode.OnExplicitShutdown;



            this.Close();



        }







        protected override void OnClosed(EventArgs e)



        {



            base.OnClosed(e);



            if (_isLoggingOut)



            {



                var login = new LoginView();



                Application.Current.MainWindow = login;



                Application.Current.ShutdownMode = ShutdownMode.OnMainWindowClose;



                login.Show();



            }



        }



    }



}