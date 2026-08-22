using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Microsoft.EntityFrameworkCore;
using TaxPersonnelManagement.Data;
using TaxPersonnelManagement.Models;
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


        /// <summary>
        /// Khởi tạo cửa sổ chính sau khi đăng nhập thành công.
        /// </summary>
        /// <param name="user">Thông tin người dùng hiện tại.</param>
        public MainWindow(User user)
        {
            App.DebugLog("MainWindow Constructor Entry");
            InitializeComponent();
            _currentUser = user;
            App.CurrentUser = user; // Lưu thông tin người dùng vào biến toàn cục của Ứng dụng
            txtWelcome.Text = _currentUser.FullName; // Hiển thị tên người dùng trên thanh tiêu đề

            // Ẩn menu 'Tài khoản' và 'Sao lưu' nếu người dùng không phải là Quản trị viên (Admin)
            if (_currentUser.Role != UserRole.Admin)
            {
                btnUsers.Visibility = Visibility.Collapsed;
                btnBackupRestore.Visibility = Visibility.Collapsed;
            }

            // Điều hướng mặc định tới màn hình Tổng quan (Dashboard) khi mở app
            NavigateDashboard(null, null);
            SetVersionInfo();

            // Auto-collapse sidebar khi app khởi động trên màn hình nhỏ
            this.Loaded += (s, e) =>
            {
                if (this.ActualWidth < 1400)
                    CollapseSidebar();

                // Kiểm tra và hiển thị cảnh báo công chức sắp đi làm lại
                CheckLeaveReturnAlerts();
            };
        }

        private bool _isClosingHandled = false;

        /// <summary>
        /// Kiểm tra xem có công chức nào sắp hết nghỉ (thai sản / ốm / phép) trong 7 ngày tới không.
        /// Nếu có, hiển thị popup cảnh báo để nhắc nhở.
        /// </summary>
        private void CheckLeaveReturnAlerts()
        {
            try
            {
                const int AlertDays = 7;
                var today = DateTime.Today;
                var cutoff = today.AddDays(AlertDays);

                using (var db = new AppDbContext())
                {
                    // Lấy tất cả personnel có lịch nghỉ kết thúc trong khoảng [hôm nay, hôm nay + 7 ngày]
                    var personnelList = db.Personnel
                        .Include(p => p.LeaveHistories)
                        .Where(p => p.LeaveHistories.Any(l =>
                            l.EndDate.HasValue &&
                            l.EndDate.Value.Date >= today &&
                            l.EndDate.Value.Date <= cutoff &&
                            l.StartDate.Date <= today))
                        .ToList();

                    var alerts = new List<LeaveReturnInfo>();

                    foreach (var p in personnelList)
                    {
                        // Lấy lịch nghỉ đang active và sắp kết thúc gần nhất
                        var activeLeave = p.LeaveHistories
                            .Where(l =>
                                l.EndDate.HasValue &&
                                l.EndDate.Value.Date >= today &&
                                l.EndDate.Value.Date <= cutoff &&
                                l.StartDate.Date <= today)
                            .OrderBy(l => l.EndDate)
                            .FirstOrDefault();

                        if (activeLeave == null) continue;

                        // Ngày đi làm lại = ngày kết thúc nghỉ + 1
                        var returnDate = activeLeave.EndDate!.Value.Date.AddDays(1);
                        int daysLeft = (returnDate - today).Days;

                        alerts.Add(new LeaveReturnInfo
                        {
                            FullName = p.FullName,
                            Department = p.Department,
                            LeaveType = activeLeave.LeaveType,
                            ReturnDate = returnDate,
                            DaysLeft = daysLeft
                        });
                    }

                    // Sắp xếp theo ngày đi làm lại gần nhất
                    alerts = alerts.OrderBy(a => a.ReturnDate).ToList();

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

        /// <summary>
        /// Bắt sự kiện người dùng bấm nút [X] đóng ứng dụng.
        /// Nếu đang kết nối Google Drive, hiển thị cửa sổ thông báo đang đồng bộ trước khi đóng hẳn.
        /// </summary>
        protected override async void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            if (!_isClosingHandled && App.DriveSync.HasSavedToken())
            {
                e.Cancel = true; // Tạm dừng đóng ứng dụng để chạy đồng bộ
                _isClosingHandled = true;

                var syncDialog = new SyncOnCloseWindow();
                syncDialog.Show();

                try
                {
                    // Đảm bảo DriveService đã nạp token
                    if (!App.DriveSync.IsConnected)
                    {
                        await App.DriveSync.TryLoadSavedTokenAsync();
                    }

                    if (App.DriveSync.IsConnected)
                    {
                        // Chạy PushAsync lên Google Drive (tối đa 10 giây)
                        var pushTask = App.DriveSync.PushAsync();
                        await System.Threading.Tasks.Task.WhenAny(pushTask, System.Threading.Tasks.Task.Delay(10000));
                    }
                }
                catch { }
                finally
                {
                    syncDialog.Close();
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
        }

        /// <summary>
        /// Điều hướng tới màn hình Tổng quan
        /// </summary>
        public void NavigateToDashboard(int? targetPersonnelId = null)
        {
            UpdateMenuState(btnDashboard); // Cập nhật trạng thái hiển thị của nút menu

            if (_dashboardCache == null)
            {
                _dashboardCache = new DashboardView(targetPersonnelId);
            }
            else
            {
                if (targetPersonnelId.HasValue)
                {
                    _dashboardCache.PersonnelList.TargetPersonnelId = targetPersonnelId;
                }
                // Làm mới dữ liệu để hiển thị thay đổi mới nhất, giữ nguyên bộ lọc đang chọn
                _dashboardCache.PersonnelList.LoadData();
            }

            MainFrame.Navigate(_dashboardCache); // Tải nội dung View vào Frame chính
        }

        private void NavigateDashboard(object? sender, RoutedEventArgs? e)
        {
            NavigateToDashboard();
        }

        private void NavigateStatistics(object sender, RoutedEventArgs e)
        {
            UpdateMenuState(btnStatistics);
            if (_statisticsCache == null)
            {
                _statisticsCache = new StatisticsView();
            }
            else
            {
                _statisticsCache.LoadStatistics();
            }
            MainFrame.Navigate(_statisticsCache);
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
                MainFrame.Navigate(_personnelDetailCache);
            }
            else
            {
                // Nếu là chỉnh sửa nhân sự cụ thể, tạo view mới (hoặc có thể cache theo ID nếu cần, 
                // nhưng hiện tại ưu tiên fix cho phần "Thêm mới" như yêu cầu)
                _personnelDetailCache = new PersonnelDetailView(p, activeTab);
                MainFrame.Navigate(_personnelDetailCache);
            }
        }

        public void ClearPersonnelCache()
        {
            _personnelDetailCache = null;
        }


        private void NavigatePersonnel(object sender, RoutedEventArgs e)
        {
            NavigateToPersonnelDetail(null);
            // Ghi chú: Dùng màn hình Tổng quan để xem danh sách nhân sự.
        }


        private void NavigateSalary(object sender, RoutedEventArgs e)
        {
            UpdateMenuState(btnSalary);
            MainFrame.Navigate(new SalaryListView());
        }

        private void NavigateAnnualIncome(object sender, RoutedEventArgs e)
        {
            UpdateMenuState(btnAnnualIncome);
            MainFrame.Navigate(new AnnualIncomeView());
        }

        private void NavigateLeaveDetail(object sender, RoutedEventArgs e)
        {
            UpdateMenuState(btnLeaveDetail);
            MainFrame.Navigate(new LeaveDetailView());
        }

        private void NavigateEmulationReward(object sender, RoutedEventArgs e)
        {
            UpdateMenuState(btnEmulationReward);
            MainFrame.Navigate(new EmulationRewardView());
        }

        private void NavigateEvaluation(object sender, RoutedEventArgs e)
        {
            UpdateMenuState(btnEvaluation);
            MainFrame.Navigate(new EvaluationListView());
        }

        private void NavigateTraining(object sender, RoutedEventArgs e)
        {
            UpdateMenuState(btnTraining);
            MainFrame.Navigate(new TrainingListView());
        }

        private void NavigatePlanning(object sender, RoutedEventArgs e)
        {
            UpdateMenuState(btnPlanning);
            if (_planningCache == null)
            {
                _planningCache = new PlanningManagementView();
            }
            else
            {
                _planningCache.LoadData();
            }
            MainFrame.Navigate(_planningCache);
        }

        private void NavigatePositionDuration(object sender, RoutedEventArgs e)
        {
            UpdateMenuState(btnPositionDuration);
            MainFrame.Navigate(new PositionDurationView());
        }

        private void NavigateUsers(object sender, RoutedEventArgs e)
        {
            // Chỉ Admin mới có quyền truy cập màn hình quản lý tài khoản
            if (_currentUser == null || _currentUser.Role != UserRole.Admin)
            {
                MessageBox.Show("Bạn không có quyền truy cập!", "Cảnh báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            UpdateMenuState(btnUsers);
            MainFrame.Navigate(new UserListView());
        }

        private void NavigateBackupRestore(object sender, RoutedEventArgs e)
        {
            if (_currentUser == null || _currentUser.Role != UserRole.Admin)
            {
                MessageBox.Show("Bạn không có quyền truy cập!", "Cảnh báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            UpdateMenuState(btnBackupRestore);
            MainFrame.Navigate(new BackupRestoreView());
        }

        private void UpdateMenuState(System.Windows.Controls.Button activeButton)
        {
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
            btnUsers.Background = transparent;
            btnBackupRestore.Background = transparent;

            // Đánh dấu nút đang active với nền trắng bán trong suốt (~30% opacity)
            activeButton.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromArgb(80, 255, 255, 255));
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
            txtUsers.Visibility = Visibility.Collapsed;
            txtBackupRestore.Visibility = Visibility.Collapsed;
            txtLogout.Visibility = Visibility.Collapsed;
            txtCopyright.Visibility = Visibility.Collapsed;
            txtVersion.Visibility = Visibility.Collapsed;
            imgLogo.Margin = new Thickness(0);

            var buttons = new[] { btnDashboard, btnStatistics, btnPersonnel, btnSalary, btnAnnualIncome, btnLeaveDetail, btnPositionDuration, btnEmulationReward, btnEvaluation, btnTraining, btnPlanning, btnUsers, btnBackupRestore, btnLogout };
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
            txtUsers.Visibility = Visibility.Visible;
            txtBackupRestore.Visibility = Visibility.Visible;
            txtLogout.Visibility = Visibility.Visible;
            txtCopyright.Visibility = Visibility.Visible;
            txtVersion.Visibility = Visibility.Visible;
            imgLogo.Margin = new Thickness(0, 0, 10, 0);

            var buttons = new[] { btnDashboard, btnStatistics, btnPersonnel, btnSalary, btnAnnualIncome, btnLeaveDetail, btnPositionDuration, btnEmulationReward, btnEvaluation, btnTraining, btnPlanning, btnUsers, btnBackupRestore, btnLogout };
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

        private void Logout_Click(object sender, RoutedEventArgs e)
        {
            LoginView login = new LoginView();
            login.Show();
            this.Close();
        }
    }
}