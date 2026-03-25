using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using TaxPersonnelManagement.Data;
using TaxPersonnelManagement.Models;

namespace TaxPersonnelManagement.Views
{
    public partial class AddDepartmentDialog : Window
    {
        public ObservableCollection<Department> Departments { get; set; } = new ObservableCollection<Department>();
        public string? SelectedDepartment { get; private set; }

        public AddDepartmentDialog()
        {
            InitializeComponent();
            LoadDepartments();
        }

        private void LoadDepartments()
        {
            using (var context = new AppDbContext())
            {
                // Ensure table exists just in case (hacky migration)
                try { context.Database.EnsureCreated(); } catch { }
                
                var list = context.Departments.OrderBy(d => d.Name).ToList();
                Departments = new ObservableCollection<Department>(list);
                lstDepartments.ItemsSource = Departments;
            }
        }

        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            var newDeptName = txtDepartmentName.Text.Trim();
            if (string.IsNullOrWhiteSpace(newDeptName))
            {
                MessageBox.Show("Vui lòng nhập tên phòng ban!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            using (var context = new AppDbContext())
            {
                if (_editingDepartment == null)
                {
                    if (context.Departments.Any(d => d.Name == newDeptName))
                    {
                        MessageBox.Show("Phòng ban này đã tồn tại!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    var newDept = new Department { Name = newDeptName };
                    context.Departments.Add(newDept);
                }
                else
                {
                    // Update
                    var dept = context.Departments.Find(_editingDepartment.Id);
                    if (dept != null)
                    {
                        dept.Name = newDeptName;
                    }
                    _editingDepartment = null;
                }
                context.SaveChanges();
            }

            txtDepartmentName.Clear();
            txtDepartmentName.Focus();
            // Reset button visual
             btnAdd.Background = (System.Windows.Media.Brush)Application.Current.Resources["PrimaryHueMidBrush"];
             btnAdd.Content = new MaterialDesignThemes.Wpf.PackIcon { Kind = MaterialDesignThemes.Wpf.PackIconKind.Plus, Width = 24, Height = 24 };
             btnAdd.ToolTip = "Thêm phòng ban";
             
            LoadDepartments();
        }

        private Department? _editingDepartment = null;

        private void btnEdit_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is Department dept)
            {
                _editingDepartment = dept;
                txtDepartmentName.Text = dept.Name;
                txtDepartmentName.Focus();
                
                // Change button to indicate Update
                btnAdd.Content = new MaterialDesignThemes.Wpf.PackIcon { Kind = MaterialDesignThemes.Wpf.PackIconKind.ContentSave, Width = 24, Height = 24 };
                btnAdd.ToolTip = "Lưu thay đổi";
                btnAdd.Background = System.Windows.Media.Brushes.Orange;
            }
        }

        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is Department dept)
            {
                var confirm = new ConfirmWindow($"Bạn có chắc muốn xóa phòng ban '{dept.Name}'?", "Xác nhận xóa");
                if (confirm.ShowDialog() == true)
                {
                     using (var context = new AppDbContext())
                    {
                        var d = context.Departments.Find(dept.Id);
                        if (d != null)
                        {
                            context.Departments.Remove(d);
                            context.SaveChanges();
                        }
                    }
                    LoadDepartments();
                }
            }
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }

        private void ListBoxItem_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (lstDepartments.SelectedItem is Department selected)
            {
                SelectedDepartment = selected.Name;
                DialogResult = true;
                Close();
            }
        }
    }
}
