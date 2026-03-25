using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using TaxPersonnelManagement.Data;
using TaxPersonnelManagement.Models;
using System.Collections.ObjectModel;
using Microsoft.EntityFrameworkCore;

namespace TaxPersonnelManagement.Views
{
    public partial class DisciplineConfigDialog : Window
    {
        private ObservableCollection<DisciplineType> _items = new ObservableCollection<DisciplineType>();
        private DisciplineType? _editingItem = null;

        public DisciplineConfigDialog()
        {
            InitializeComponent();
            LoadData();
            lstItems.ItemsSource = _items;
        }

        private void LoadData()
        {
            _items.Clear();
            
            try 
            {
                using (var context = new AppDbContext())
                {
                    // Ensure table exists (just in case migration failed or direct SQL needed)
                    // But usually App.xaml.cs handles migration. We'll trust it or use simple check.
                    
                    var list = context.DisciplineTypes.OrderBy(x => x.Id).ToList();
                    
                    // Seed if empty
                    if (list.Count == 0)
                    {
                        var defaults = new[] 
                        {
                            "Khiển trách",
                            "Cảnh cáo",
                            "Hạ bậc lương",
                            "Giáng chức",
                            "Cách chức",
                            "Buộc thôi việc"
                        };

                        foreach (var name in defaults)
                        {
                            context.DisciplineTypes.Add(new DisciplineType { Name = name });
                        }
                        context.SaveChanges();
                        list = context.DisciplineTypes.OrderBy(x => x.Id).ToList();
                    }

                    foreach (var item in list)
                    {
                        _items.Add(item);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi tải dữ liệu: {ex.Message}", "Lỗi Hệ Thống", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtName.Text))
            {
                MessageBox.Show("Vui lòng nhập tên hình thức kỷ luật!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                using (var context = new AppDbContext())
                {
                    if (_editingItem == null)
                    {
                        // Add new
                        var newItem = new DisciplineType { Name = txtName.Text.Trim() };
                        context.DisciplineTypes.Add(newItem);
                        context.SaveChanges();
                    }
                    else
                    {
                        // Update
                        var itemToUpdate = context.DisciplineTypes.Find(_editingItem.Id);
                        if (itemToUpdate != null)
                        {
                            itemToUpdate.Name = txtName.Text.Trim();
                            context.SaveChanges();
                        }
                        
                        // Reset UI
                        _editingItem = null;
                        if (btnAdd.Content is StackPanel sp && sp.Children[1] is TextBlock tb)
                        {
                            tb.Text = "Thêm";
                        }
                    }
                }

                txtName.Clear();
                LoadData(); // reload list
                txtName.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi lưu dữ liệu: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void btnEdit_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is int id)
            {
                 var item = _items.FirstOrDefault(x => x.Id == id);
                 if (item != null)
                 {
                     _editingItem = item;
                     txtName.Text = item.Name;
                     
                     // Change button text to "Lưu"
                     if (btnAdd.Content is StackPanel sp && sp.Children.Count > 1 && sp.Children[1] is TextBlock tb)
                     {
                         tb.Text = "Lưu";
                     }
                     txtName.Focus();
                 }
            }
        }

        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is int id)
            {
                var dialog = new ConfirmDialog("Bạn có chắc chắn muốn xóa dòng này?");
                if (dialog.ShowDialog() == true)
                {
                    using (var context = new AppDbContext())
                    {
                        var item = context.DisciplineTypes.Find(id);
                        if (item != null)
                        {
                            context.DisciplineTypes.Remove(item);
                            context.SaveChanges();
                            LoadData();
                        }
                    }
                }
            }
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
