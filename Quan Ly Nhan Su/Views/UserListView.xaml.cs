using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using TaxPersonnelManagement.Data;
using TaxPersonnelManagement.Models;
using System.Collections.Generic;

namespace TaxPersonnelManagement.Views
{
    public partial class UserListView : UserControl
    {
        private List<User> _allUsers = new List<User>();

        public UserListView()
        {
            InitializeComponent();
            LoadData();
        }

        private void LoadData()
        {
            try
            {
                using (var context = new AppDbContext())
                {
                    _allUsers = context.Users.ToList();
                    dgUsers.ItemsSource = _allUsers;
                }
            }
            catch (Exception ex)
            {
                var warning = new WarningWindow("Lỗi hệ thống", "Lỗi tải dữ liệu tài khoản: " + ex.Message);
                warning.Owner = Window.GetWindow(this);
                warning.ShowDialog();
            }
        }

        private void TxtSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            string keyword = txtSearch.Text.ToLower();
            dgUsers.ItemsSource = _allUsers.Where(u => 
                (u.Username != null && u.Username.ToLower().Contains(keyword)) ||
                (u.FullName != null && u.FullName.ToLower().Contains(keyword))).ToList();
        }

        private void BtnAddUser_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new UserEditDialog();
            dialog.Owner = Window.GetWindow(this);
            dialog.ShowDialog();

            if (dialog.IsSuccess)
            {
                LoadData();
                var success = new SuccessWindow("Đã thêm tài khoản mới thành công!");
                success.Owner = Window.GetWindow(this);
                success.ShowDialog();
            }
        }

        private void BtnEditUser_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is int userId)
            {
                var dialog = new UserEditDialog(userId);
                dialog.Owner = Window.GetWindow(this);
                dialog.ShowDialog();

                if (dialog.IsSuccess)
                {
                    LoadData();
                    var success = new SuccessWindow("Cập nhật thông tin thành công!");
                    success.Owner = Window.GetWindow(this);
                    success.ShowDialog();
                }
            }
        }

        private void BtnChangePassword_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is int userId)
            {
                var user = _allUsers.FirstOrDefault(u => u.Id == userId);
                if (user != null)
                {
                    var dialog = new ChangePasswordDialog(user.Id, user.Username);
                    dialog.Owner = Window.GetWindow(this);
                    dialog.ShowDialog();

                    if (dialog.IsSuccess)
                    {
                        var success = new SuccessWindow("Đổi mật khẩu thành công!");
                        success.Owner = Window.GetWindow(this);
                        success.ShowDialog();
                    }
                }
            }
        }

        private void BtnDeleteUser_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is int userId)
            {
                var userToDelete = _allUsers.FirstOrDefault(u => u.Id == userId);
                if (userToDelete != null)
                {
                    if (userToDelete.Username.ToLower() == "admin")
                    {
                        var warningMsg = new WarningWindow("Từ chối truy cập", "Không thể xóa tài khoản Quản trị viên hệ thống (admin).");
                        warningMsg.Owner = Window.GetWindow(this);
                        warningMsg.ShowDialog();
                        return;
                    }

                    var confirm = new ConfirmDialog($"Bạn có chắc chắn muốn xóa tài khoản '{userToDelete.Username}'?");
                    confirm.Owner = Window.GetWindow(this);
                    
                    if (confirm.ShowDialog() == true)
                    {
                        try
                        {
                            using (var db = new AppDbContext())
                            {
                                var user = db.Users.Find(userId);
                                if (user != null)
                                {
                                    db.Users.Remove(user);
                                    db.SaveChanges();
                                    
                                    LoadData();
                                    txtSearch.Text = string.Empty; // Reset search

                                    var success = new SuccessWindow("Đã xóa tài khoản thành công!");
                                    success.Owner = Window.GetWindow(this);
                                    success.ShowDialog();
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            var err = new WarningWindow("Lỗi", "Không thể xóa tài khoản: " + ex.Message);
                            err.Owner = Window.GetWindow(this);
                            err.ShowDialog();
                        }
                    }
                }
            }
        }
    }
}
