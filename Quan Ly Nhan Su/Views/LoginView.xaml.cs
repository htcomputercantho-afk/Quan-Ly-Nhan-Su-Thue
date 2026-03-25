using System.Linq;
using System.Windows;
using TaxPersonnelManagement.Data;
using TaxPersonnelManagement.Models;
using System;
using System.IO;
using System.Text.Json;

namespace TaxPersonnelManagement.Views
{
    public partial class LoginView : Window
    {
        private const string CredentialsFile = "user.config";

        public LoginView()
        {
            InitializeComponent();
            LoadCredentials();
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
                        try {
                            byte[] data = Convert.FromBase64String(creds.Password);
                            txtPassword.Password = System.Text.Encoding.UTF8.GetString(data);
                        } catch { txtPassword.Password = ""; }
                        
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

        private void btnLogin_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string username = txtUsername.Text;
                string password = txtPassword.Password;

                if (string.IsNullOrWhiteSpace(username) || string.IsNullOrWhiteSpace(password))
                {
                    MessageBox.Show("Vui lòng nhập đầy đủ thông tin!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                using (var context = new AppDbContext())
                {
                    var user = context.Users.FirstOrDefault(u => u.Username == username && u.PasswordHash == password);
                    if (user != null)
                    {
                        // Save credentials if checked
                        SaveCredentials(username, password);

                        // Login Success
                        App.DebugLog("Login successful. Creating MainWindow...");
                        MainWindow dashboard = new MainWindow(user);
                        App.DebugLog("MainWindow created. Showing...");
                        Application.Current.MainWindow = dashboard; // Prevent shutdown when Login closes
                        dashboard.Show();
                        App.DebugLog("MainWindow shown. Closing LoginView...");
                        this.Close();
                    }
                    else
                    {
                        MessageBox.Show("Sai tên đăng nhập hoặc mật khẩu!", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                 App.DebugLog($"Login Error: {ex.Message}");
                 MessageBox.Show($"Lỗi đăng nhập: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void btnExit_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        private bool _isSyncing = false;

        private void btnTogglePassword_Click(object sender, RoutedEventArgs e)
        {
            if (sender is not System.Windows.Controls.Primitives.ToggleButton toggleButton) return;
            var icon = toggleButton.Content as MaterialDesignThemes.Wpf.PackIcon;

            if (toggleButton.IsChecked == true)
            {
                // Show Password
                _isSyncing = true;
                txtVisiblePassword.Text = txtPassword.Password;
                _isSyncing = false;
                
                txtVisiblePassword.Visibility = Visibility.Visible;
                txtPassword.Visibility = Visibility.Collapsed;
                if (icon != null) icon.Kind = MaterialDesignThemes.Wpf.PackIconKind.EyeOff;
            }
            else
            {
                // Hide Password
                _isSyncing = true;
                txtPassword.Password = txtVisiblePassword.Text;
                _isSyncing = false;

                txtPassword.Visibility = Visibility.Visible;
                txtVisiblePassword.Visibility = Visibility.Collapsed;
                if (icon != null) icon.Kind = MaterialDesignThemes.Wpf.PackIconKind.Eye;
            }
        }

        private void txtPassword_PasswordChanged(object sender, RoutedEventArgs e)
        {
             if (_isSyncing) return;
             
             // If typing in PasswordBox (Hidden Mode), update hidden TextBox value
             // (though TextBox is collapsed, we keep it ready or just sync on toggle).
             // Syncing on toggle is safer to avoid string mismanagement, 
             // but if we want 'real-time' state, we can sync. 
             // Currently, we only sync ON TOGGLE in the click handler, which is sufficient for this simple use case.
             // But if specific requirement needs typing in visible mode:
             
             if (txtPassword.Visibility == Visibility.Visible)
             {
                 _isSyncing = true;
                 txtVisiblePassword.Text = txtPassword.Password;
                 _isSyncing = false;
             }
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
        }
    }
}
