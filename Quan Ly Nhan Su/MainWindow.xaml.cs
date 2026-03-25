using System.Windows;
using TaxPersonnelManagement.Models;
using TaxPersonnelManagement.Views;

namespace TaxPersonnelManagement
{
    public partial class MainWindow : Window
    {
        private User? _currentUser;

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
        }

        // For Designer support
        public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Điều hướng tới màn hình Tổng quan
        /// </summary>
        public void NavigateToDashboard()
        {
            UpdateMenuState(btnDashboard); // Cập nhật trạng thái hiển thị của nút menu
            MainFrame.Navigate(new DashboardView()); // Tải nội dung View vào Frame chính
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
            MainFrame.Navigate(new PersonnelDetailView(p, activeTab));
        }

        private void NavigatePersonnel(object sender, RoutedEventArgs e)
        {
            UpdateMenuState(btnPersonnel);
            MainFrame.Navigate(new PersonnelDetailView(null));
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
            btnUsers.Background = transparent;
            btnBackupRestore.Background = transparent;

            // Set active button background (Semi-transparent white, stronger than hover)
            // Using a new SolidColorBrush(Color.FromArgb(80, 255, 255, 255)) -> approx 30% opacity
            activeButton.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromArgb(80, 255, 255, 255));
        }

        private bool _isSidebarExpanded = true;

        private void btnToggleSidebar_Click(object sender, RoutedEventArgs e)
        {
            _isSidebarExpanded = !_isSidebarExpanded;

            if (_isSidebarExpanded)
            {
                // Phóng to
                colSidebar.Width = new GridLength(250);
                txtLogo.Visibility = Visibility.Visible;
                txtOverview.Visibility = Visibility.Visible;
                txtPersonnel.Visibility = Visibility.Visible;
                txtSalary.Visibility = Visibility.Visible;
                txtAnnualIncome.Visibility = Visibility.Visible;
                txtLeaveDetail.Visibility = Visibility.Visible;
                txtUsers.Visibility = Visibility.Visible;
                txtBackupRestore.Visibility = Visibility.Visible;
                txtLogout.Visibility = Visibility.Visible;
                txtCopyright.Visibility = Visibility.Visible;

                imgLogo.Margin = new Thickness(0, 0, 10, 0);

                var buttons = new[] { btnDashboard, btnPersonnel, btnSalary, btnAnnualIncome, btnLeaveDetail, btnUsers, btnBackupRestore, btnLogout };
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
            else
            {
                // Thu gọn
                colSidebar.Width = new GridLength(70);
                txtLogo.Visibility = Visibility.Collapsed;
                txtOverview.Visibility = Visibility.Collapsed;
                txtPersonnel.Visibility = Visibility.Collapsed;
                txtSalary.Visibility = Visibility.Collapsed;
                txtAnnualIncome.Visibility = Visibility.Collapsed;
                txtLeaveDetail.Visibility = Visibility.Collapsed;
                txtUsers.Visibility = Visibility.Collapsed;
                txtBackupRestore.Visibility = Visibility.Collapsed;
                txtLogout.Visibility = Visibility.Collapsed;
                txtCopyright.Visibility = Visibility.Collapsed;

                imgLogo.Margin = new Thickness(0);

                var buttons = new[] { btnDashboard, btnPersonnel, btnSalary, btnAnnualIncome, btnLeaveDetail, btnUsers, btnBackupRestore, btnLogout };
                foreach (var btn in buttons)
                {
                    btn.Padding = new Thickness(0);
                    if (btn.Content is System.Windows.Controls.StackPanel sp && sp.Children.Count > 0)
                    {
                        var icon = sp.Children[0] as FrameworkElement;
                        // Canh giữa icon khi thu gọn (Width cột = 70, nút có Margin ngang = 10 -> Width nút = 50. Icon=22. Left Margin = (50-22)/2 = 14)
                        if (icon != null) icon.Margin = new Thickness(14, 0, 0, 0);
                    }
                }
            }
        }

        private void Logout_Click(object sender, RoutedEventArgs e)
        {
            LoginView login = new LoginView();
            login.Show();
            this.Close();
        }
    }
}