using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using TaxPersonnelManagement.Data;
using TaxPersonnelManagement.Models;
using System;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using MaterialDesignThemes.Wpf;

namespace TaxPersonnelManagement.Views
{
    public partial class LoginView : Window
    {
        private const string CredentialsFile = "user.config";

        public LoginView()
        {
            InitializeComponent();
            LoadCredentials();
            SetVersionInfo();
            UpdatePasswordPlaceholder();

            // Khi phóng to cửa sổ không viền, chừa lề để nội dung không bị tràn ra ngoài màn hình
            StateChanged += (s, e) =>
                RootGrid.Margin = WindowState == WindowState.Maximized ? new Thickness(7) : new Thickness(0);

            Loaded += (s, e) =>
            {
                UpdateResponsiveLayout(RootGrid.ActualWidth, RootGrid.ActualHeight);
                Dispatcher.BeginInvoke(new Action(UpdateGlass), System.Windows.Threading.DispatcherPriority.Loaded);
                if (string.IsNullOrEmpty(txtUsername.Text)) txtUsername.Focus();
                else txtPassword.Focus();
            };
        }

        private void SetVersionInfo()
        {
            var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            txtVersion.Text = $"Phiên bản {version?.Major}.{version?.Minor}.{version?.Build}.{version?.Revision}";
        }

        private class SavedCredentials
        {
            public string Username { get; set; } = string.Empty;
            public string Password { get; set; } = string.Empty;
        }

        private void LoadCredentials()
        {
            try
            {
                if (System.IO.File.Exists(CredentialsFile))
                {
                    App.DebugLog("Found saved credentials file. Loading...");
                    string json = System.IO.File.ReadAllText(CredentialsFile);
                    var creds = System.Text.Json.JsonSerializer.Deserialize<SavedCredentials>(json);
                    if (creds != null)
                    {
                        txtUsername.Text = creds.Username;
                        // Simple Base64 decode
                        try
                        {
                            byte[] data = Convert.FromBase64String(creds.Password);
                            txtPassword.Password = System.Text.Encoding.UTF8.GetString(data);
                        }
                        catch { txtPassword.Password = ""; }

                        chkRememberMe.IsChecked = true;
                        App.DebugLog("Credentials loaded.");
                    }
                }
                else
                {
                    App.DebugLog("No saved credentials found.");
                }
            }
            catch (Exception ex) { App.DebugLog($"Error loading credentials: {ex.Message}"); }
        }

        private void SaveCredentials(string username, string password)
        {
            try
            {
                if (chkRememberMe.IsChecked == true)
                {
                    // Simple Base64 encode for basic obfuscation
                    string encodedPass = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(password));
                    var creds = new SavedCredentials { Username = username, Password = encodedPass };
                    string json = System.Text.Json.JsonSerializer.Serialize(creds);
                    System.IO.File.WriteAllText(CredentialsFile, json);
                }
                else
                {
                    if (System.IO.File.Exists(CredentialsFile))
                        System.IO.File.Delete(CredentialsFile);
                }
            }
            catch (Exception ex)
            {
                App.DebugLog($"Error saving credentials: {ex.Message}");
            }
        }

        // ── Đăng nhập ─────────────────────────────────────────────────────
        private bool _isLoggingIn = false;

        private async void btnLogin_Click(object sender, RoutedEventArgs e) => await PerformLoginAsync();

        private async void Input_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                await PerformLoginAsync();
            }
        }

        private async Task PerformLoginAsync()
        {
            if (_isLoggingIn) return;
            HideBanner();

            string username = txtUsername.Text.Trim();
            string password = txtPassword.Password;

            if (string.IsNullOrWhiteSpace(username) || string.IsNullOrWhiteSpace(password))
            {
                ShowBanner(BannerType.Warning, "Vui lòng nhập đầy đủ tên đăng nhập và mật khẩu.");
                return;
            }

            SetLoadingState(true);
            try
            {
                User? user = await Task.Run(() =>
                {
                    using var context = new AppDbContext();
                    return context.Users.FirstOrDefault(u => u.Username == username && u.PasswordHash == password);
                });

                if (user != null)
                {
                    // Save credentials if checked
                    SaveCredentials(username, password);

                    // Login Success
                    App.DebugLog("Login successful. Creating MainWindow...");
                    MainWindow dashboard = new MainWindow(user);
                    App.DebugLog("MainWindow created. Showing...");
                    Application.Current.MainWindow = dashboard; // Prevent shutdown when Login closes
                    this.Hide(); // Ẩn LoginView ngay lập tức để người dùng thấy MainWindow vào việc
                    dashboard.Show();
                    App.DebugLog("MainWindow shown. Closing LoginView...");
                    this.Close();
                    return;
                }

                ShowBanner(BannerType.Error, "Tên đăng nhập hoặc mật khẩu không chính xác.");
            }
            catch (Exception ex)
            {
                App.DebugLog($"Login Error: {ex.Message}");
                ShowBanner(BannerType.Error, $"Lỗi đăng nhập: {ex.Message}");
            }
            finally
            {
                if (IsLoaded) SetLoadingState(false);
            }
        }

        private void SetLoadingState(bool isLoading)
        {
            _isLoggingIn = isLoading;
            btnLogin.IsEnabled = !isLoading;
            txtUsername.IsEnabled = !isLoading;
            txtPassword.IsEnabled = !isLoading;
            txtVisiblePassword.IsEnabled = !isLoading;

            LoginProgress.Visibility = isLoading ? Visibility.Visible : Visibility.Collapsed;
            LoginBtnIcon.Visibility = isLoading ? Visibility.Collapsed : Visibility.Visible;
            LoginButtonText.Text = isLoading ? "ĐANG XÁC THỰC..." : "ĐĂNG NHẬP";
        }

        private void btnExit_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        // ── Thanh tiêu đề tùy chỉnh ───────────────────────────────────────
        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2)
            {
                WindowState = WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized;
                return;
            }
            try { DragMove(); } catch { }
        }

        private void btnMinimize_Click(object sender, RoutedEventArgs e)
        {
            WindowState = WindowState.Minimized;
        }

        // ── Khung thông báo nội tuyến ─────────────────────────────────────
        private enum BannerType { Warning, Error, Success }

        private static SolidColorBrush Brush(string hex) => (SolidColorBrush)new BrushConverter().ConvertFromString(hex)!;

        private void ShowBanner(BannerType type, string message)
        {
            BannerMessage.Text = message;
            switch (type)
            {
                case BannerType.Warning:
                    ErrorBanner.Background = Brush("#FFF3CD");
                    ErrorBanner.BorderBrush = Brush("#FBBF24");
                    BannerMessage.Foreground = Brush("#785000");
                    BannerIcon.Kind = PackIconKind.AlertCircleOutline;
                    BannerIcon.Foreground = Brush("#C88200");
                    break;
                case BannerType.Error:
                    ErrorBanner.Background = Brush("#FEE2E2");
                    ErrorBanner.BorderBrush = Brush("#F87171");
                    BannerMessage.Foreground = Brush("#991B1B");
                    BannerIcon.Kind = PackIconKind.CloseCircleOutline;
                    BannerIcon.Foreground = Brush("#DC2626");
                    break;
                case BannerType.Success:
                    ErrorBanner.Background = Brush("#DCFCE7");
                    ErrorBanner.BorderBrush = Brush("#4ADE80");
                    BannerMessage.Foreground = Brush("#15803D");
                    BannerIcon.Kind = PackIconKind.CheckCircleOutline;
                    BannerIcon.Foreground = Brush("#16A34A");
                    break;
            }
            BannerCloseIcon.Foreground = BannerMessage.Foreground;
            ErrorBanner.Visibility = Visibility.Visible;
        }

        private void HideBanner() => ErrorBanner.Visibility = Visibility.Collapsed;

        private void btnDismissBanner_Click(object sender, RoutedEventArgs e) => HideBanner();

        // ── Hiệu ứng focus ô nhập ─────────────────────────────────────────
        private static void SetFieldFocused(System.Windows.Controls.Border border, bool focused)
        {
            border.BorderBrush = focused ? Brush("#1A84F8") : Brush("#7090CAF9");
            border.BorderThickness = new Thickness(focused ? 1.5 : 1);
            border.Background = focused ? Brush("#EBFFFFFF") : Brush("#C8FFFFFF");
        }

        private void Username_GotFocus(object sender, KeyboardFocusChangedEventArgs e) => SetFieldFocused(UsernameBorder, true);
        private void Username_LostFocus(object sender, KeyboardFocusChangedEventArgs e) => SetFieldFocused(UsernameBorder, false);
        private void Password_GotFocus(object sender, KeyboardFocusChangedEventArgs e) => SetFieldFocused(PasswordBorder, true);
        private void Password_LostFocus(object sender, KeyboardFocusChangedEventArgs e) => SetFieldFocused(PasswordBorder, false);

        // ── Ẩn / hiện mật khẩu ────────────────────────────────────────────
        private bool _isSyncing = false;
        private bool _showPassword = false;

        private void btnTogglePassword_Click(object sender, RoutedEventArgs e)
        {
            _showPassword = !_showPassword;
            _isSyncing = true;
            if (_showPassword)
            {
                txtVisiblePassword.Text = txtPassword.Password;
                txtVisiblePassword.Visibility = Visibility.Visible;
                txtPassword.Visibility = Visibility.Collapsed;
                txtVisiblePassword.Focus();
                txtVisiblePassword.CaretIndex = txtVisiblePassword.Text.Length;
                EyeIcon.Kind = PackIconKind.EyeOffOutline;
                EyeIcon.Foreground = Brush("#1890FF");
            }
            else
            {
                txtPassword.Password = txtVisiblePassword.Text;
                txtPassword.Visibility = Visibility.Visible;
                txtVisiblePassword.Visibility = Visibility.Collapsed;
                txtPassword.Focus();
                EyeIcon.Kind = PackIconKind.EyeOutline;
                EyeIcon.Foreground = Brush("#4A90C8");
            }
            _isSyncing = false;
            UpdatePasswordPlaceholder();
        }

        private void txtPassword_PasswordChanged(object sender, RoutedEventArgs e)
        {
            UpdatePasswordPlaceholder();
            if (_isSyncing) return;
            _isSyncing = true;
            txtVisiblePassword.Text = txtPassword.Password;
            _isSyncing = false;
        }

        private void txtVisiblePassword_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            if (_isSyncing) return;
            if (txtVisiblePassword.Visibility == Visibility.Visible)
            {
                _isSyncing = true;
                txtPassword.Password = txtVisiblePassword.Text;
                _isSyncing = false;
            }
            UpdatePasswordPlaceholder();
        }

        private void UpdatePasswordPlaceholder()
        {
            if (PasswordPlaceholder == null) return;
            // Khi đang hiện mật khẩu, TextBox tự có placeholder riêng
            PasswordPlaceholder.Visibility = (!_showPassword && string.IsNullOrEmpty(txtPassword.Password))
                ? Visibility.Visible : Visibility.Collapsed;
        }

        // ── Hiệu ứng kính mờ (mô phỏng Acrylic của WinUI) ─────────────────
        private void GlassHost_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            GlassHost.Clip = new RectangleGeometry(new Rect(0, 0, GlassHost.ActualWidth, GlassHost.ActualHeight), 22, 22);
            UpdateGlass();
        }

        private void ContentScrollViewer_ScrollChanged(object sender, System.Windows.Controls.ScrollChangedEventArgs e) => UpdateGlass();

        private void UpdateGlass()
        {
            if (GlassHost == null || BackgroundLayer == null || GlassHost.ActualWidth <= 0 || BackgroundLayer.ActualWidth <= 0) return;
            try
            {
                const double pad = 40; // khớp với Margin="-40" của GlassBlur
                Point origin = GlassHost.TranslatePoint(new Point(-pad, -pad), BackgroundLayer);
                GlassBrush.Viewbox = new Rect(origin, new Size(GlassHost.ActualWidth + pad * 2, GlassHost.ActualHeight + pad * 2));
            }
            catch { }
        }

        // ── Bố cục thích ứng (Responsive) ─────────────────────────────────
        private void RootGrid_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            UpdateResponsiveLayout(e.NewSize.Width, e.NewSize.Height);
            Dispatcher.BeginInvoke(new Action(UpdateGlass), System.Windows.Threading.DispatcherPriority.Render);
        }

        private void UpdateResponsiveLayout(double width, double height)
        {
            if (LoginCardContainer == null || SloganPanel == null) return;

            // 1. Theo chiều rộng
            if (width >= 900)
            {
                bool large = width >= 1200;
                SloganPanel.Visibility = Visibility.Visible;
                LeftGradientOverlay.Visibility = Visibility.Visible;
                ColLeft.Width = new GridLength(1, GridUnitType.Star);
                ColRight.Width = new GridLength(1, GridUnitType.Star);

                SloganPanel.Margin = large ? new Thickness(56, 0, 24, 0) : new Thickness(32, 0, 16, 0);
                double sloganSize = large ? 32 : 26;
                SloganLine1.FontSize = SloganLine2.FontSize = SloganLine3.FontSize = sloganSize;
                SloganSubtitle1.FontSize = SloganSubtitle2.FontSize = large ? 13 : 11.5;

                Grid.SetColumn(LoginCardContainer, 1);
                Grid.SetColumnSpan(LoginCardContainer, 1);
                LoginCardContainer.HorizontalAlignment = HorizontalAlignment.Center;
                LoginCardContainer.Width = large ? 450 : 430;
                LoginCardContainer.MaxWidth = large ? 460 : 445;
                LoginCardContainer.Margin = new Thickness(16, 10, 16, 10);
                LoginCard.Padding = large ? new Thickness(42, 24, 42, 20) : new Thickness(38, 22, 38, 18);
            }
            else
            {
                bool tablet = width >= 700;
                SloganPanel.Visibility = Visibility.Collapsed;
                LeftGradientOverlay.Visibility = Visibility.Collapsed;
                ColLeft.Width = new GridLength(0);
                ColRight.Width = new GridLength(0);

                Grid.SetColumn(LoginCardContainer, 0);
                Grid.SetColumnSpan(LoginCardContainer, 3);
                LoginCardContainer.HorizontalAlignment = tablet ? HorizontalAlignment.Center : HorizontalAlignment.Stretch;
                LoginCardContainer.Width = tablet ? 440 : double.NaN;
                LoginCardContainer.MaxWidth = tablet ? 450 : 440;
                LoginCardContainer.Margin = tablet ? new Thickness(16, 10, 16, 10) : new Thickness(16, 8, 16, 8);
                LoginCard.Padding = tablet ? new Thickness(38, 22, 38, 18) : new Thickness(28, 18, 28, 16);
            }

            // 2. Theo chiều cao
            bool compact = height < 640;
            MainContentGrid.MinHeight = compact ? 380 : 500;
            MainContentGrid.Margin = compact ? new Thickness(0, 16, 0, 6) : new Thickness(0, 32, 0, 16);
            LogoBorder.Width = LogoBorder.Height = compact ? 56 : 72;
            LogoBorder.CornerRadius = new CornerRadius(compact ? 28 : 36);
            LoginTitle.FontSize = compact ? 20 : 23;
            UsernameFieldGrid.Height = PasswordFieldGrid.Height = compact ? 42 : 46;
            btnLogin.Height = compact ? 41 : 44;
            btnExit.Height = compact ? 39 : 42;
            btnExit.Margin = new Thickness(0, 0, 0, compact ? 8 : 10);
            if (compact && width >= 700)
                LoginCard.Padding = new Thickness(LoginCard.Padding.Left, 14, LoginCard.Padding.Right, 12);
        }
    }
}
