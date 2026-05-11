using System.Windows;
using TaxPersonnelManagement.Models;
using TaxPersonnelManagement.Views;

namespace TaxPersonnelManagement
{
    public partial class MainWindow : Window
    {
        private User? _currentUser;
        private PersonnelDetailView? _personnelDetailCache;


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
            txtWelcome.Text = _currentUser.FullName; // Updated to just set the Name part
            
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
            };
        }

        private void SetVersionInfo()
        {
            var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            txtVersion.Text = $"Version {version?.Major}.{version?.Minor}.{version?.Build}.{version?.Revision}";
        }

        // For Designer support
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
            MainFrame.Navigate(new DashboardView(targetPersonnelId)); // Tải nội dung View vào Frame chính
        }

        private void NavigateDashboard(object? sender, RoutedEventArgs? e)
        {
            NavigateToDashboard();
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
            // Note: Users should use "Overview" to see the list.
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

        private void NavigatePositionDuration(object sender, RoutedEventArgs e)
        {
            UpdateMenuState(btnPositionDuration);
            MainFrame.Navigate(new PositionDurationView());
        }
        
        private void NavigateUsers(object sender, RoutedEventArgs e)
        {
            // Only Admin
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
            // Reset all buttons to transparent
            var transparent = System.Windows.Media.Brushes.Transparent;
            btnDashboard.Background = transparent;
            btnPersonnel.Background = transparent;
            btnSalary.Background = transparent;
            btnAnnualIncome.Background = transparent;
            btnLeaveDetail.Background = transparent;
            btnPositionDuration.Background = transparent;
            btnEmulationReward.Background = transparent;
            btnUsers.Background = transparent;
            btnBackupRestore.Background = transparent;

            // Set active button background (Semi-transparent white, stronger than hover)
            // Using a new SolidColorBrush(Color.FromArgb(80, 255, 255, 255)) -> approx 30% opacity
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
            txtPersonnel.Visibility = Visibility.Collapsed;
            txtSalary.Visibility = Visibility.Collapsed;
            txtAnnualIncome.Visibility = Visibility.Collapsed;
            txtLeaveDetail.Visibility = Visibility.Collapsed;
            txtPositionDuration.Visibility = Visibility.Collapsed;
            txtEmulationReward.Visibility = Visibility.Collapsed;
            txtUsers.Visibility = Visibility.Collapsed;
            txtBackupRestore.Visibility = Visibility.Collapsed;
            txtLogout.Visibility = Visibility.Collapsed;
            txtCopyright.Visibility = Visibility.Collapsed;
            txtVersion.Visibility = Visibility.Collapsed;
            imgLogo.Margin = new Thickness(0);

            var buttons = new[] { btnDashboard, btnPersonnel, btnSalary, btnAnnualIncome, btnLeaveDetail, btnPositionDuration, btnEmulationReward, btnUsers, btnBackupRestore, btnLogout };
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
            txtPersonnel.Visibility = Visibility.Visible;
            txtSalary.Visibility = Visibility.Visible;
            txtAnnualIncome.Visibility = Visibility.Visible;
            txtLeaveDetail.Visibility = Visibility.Visible;
            txtPositionDuration.Visibility = Visibility.Visible;
            txtEmulationReward.Visibility = Visibility.Visible;
            txtUsers.Visibility = Visibility.Visible;
            txtBackupRestore.Visibility = Visibility.Visible;
            txtLogout.Visibility = Visibility.Visible;
            txtCopyright.Visibility = Visibility.Visible;
            txtVersion.Visibility = Visibility.Visible;
            imgLogo.Margin = new Thickness(0, 0, 10, 0);

            var buttons = new[] { btnDashboard, btnPersonnel, btnSalary, btnAnnualIncome, btnLeaveDetail, btnPositionDuration, btnEmulationReward, btnUsers, btnBackupRestore, btnLogout };
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