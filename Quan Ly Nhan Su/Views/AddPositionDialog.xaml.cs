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
    public partial class AddPositionDialog : Window
    {
        public ObservableCollection<Position> Positions { get; set; } = new ObservableCollection<Position>();
        public string? SelectedPosition { get; private set; }

        public AddPositionDialog()
        {
            InitializeComponent();
            LoadDepartments();
            LoadPositions();
        }

        private void LoadDepartments()
        {
            var deptList = new List<string> {
                "-- Tất cả bộ phận --",
                "Ban lãnh đạo",
                "Các tổ"
            };

            using (var context = new AppDbContext())
            {
                try { context.Database.EnsureCreated(); } catch { }

                var dbDepts = context.Departments
                                     .Select(d => d.Name)
                                     .Where(n => !string.IsNullOrEmpty(n))
                                     .Distinct()
                                     .ToList();

                foreach (var d in dbDepts)
                {
                    if (!deptList.Contains(d, System.StringComparer.OrdinalIgnoreCase))
                    {
                        deptList.Add(d);
                    }
                }
            }

            cboDepartment.ItemsSource = deptList;
            cboDepartment.SelectedIndex = 0;
        }

        private void LoadPositions()
        {
            var posOrder = new System.Collections.Generic.List<string> {
                "Trưởng Thuế cơ sở",
                "Quyền Trưởng Thuế cơ sở",
                "Phó Trưởng Thuế cơ sở",
                "Tổ trưởng",
                "Phó Tổ trưởng",
                "Công chức"
            };

            using (var context = new AppDbContext())
            {
                // Ensure table exists just in case
                try { context.Database.EnsureCreated(); } catch { }

                var list = context.Positions.ToList()
                                  .OrderBy(x =>
                                  {
                                      int idx = posOrder.FindIndex(p => p.Equals(x.Name, System.StringComparison.OrdinalIgnoreCase));
                                      return idx == -1 ? 999 : idx;
                                  })
                                  .ThenBy(x => x.Name)
                                  .ToList();

                Positions = new ObservableCollection<Position>(list);
                lstPositions.ItemsSource = Positions;
            }
        }

        private Position? _editingPosition = null;

        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            var newName = txtPositionName.Text.Trim();
            if (string.IsNullOrWhiteSpace(newName))
            {
                new WarningWindow("Vui lòng nhập tên chức vụ!", "Thông báo").ShowDialog();
                return;
            }

            string? selectedDept = cboDepartment.SelectedItem?.ToString();
            if (selectedDept == "-- Tất cả bộ phận --" || string.IsNullOrEmpty(selectedDept))
            {
                selectedDept = null;
            }

            using (var context = new AppDbContext())
            {
                if (_editingPosition == null)
                {
                    bool exists = context.Positions.AsEnumerable().Any(d => 
                        d.Name.Equals(newName, System.StringComparison.OrdinalIgnoreCase) && 
                        (d.DepartmentName ?? "").Equals(selectedDept ?? "", System.StringComparison.OrdinalIgnoreCase));
                    if (exists)
                    {
                        new WarningWindow("Chức vụ này thuộc bộ phận đã chọn đã tồn tại!", "Thông báo").ShowDialog();
                        return;
                    }

                    var newPos = new Position { Name = newName, DepartmentName = selectedDept };
                    context.Positions.Add(newPos);
                }
                else
                {
                    var p = context.Positions.Find(_editingPosition.Id);
                    if (p != null)
                    {
                        p.Name = newName;
                        p.DepartmentName = selectedDept;
                    }
                    _editingPosition = null;
                }
                context.SaveChanges();
            }

            txtPositionName.Clear();
            cboDepartment.SelectedIndex = 0;
            txtPositionName.Focus();

            // Reset button visual
            btnAdd.Background = (System.Windows.Media.Brush)Application.Current.Resources["PrimaryHueMidBrush"];
            btnAdd.Content = new MaterialDesignThemes.Wpf.PackIcon { Kind = MaterialDesignThemes.Wpf.PackIconKind.Plus, Width = 24, Height = 24 };
            btnAdd.ToolTip = "Thêm chức vụ";

            LoadPositions();
        }

        private void btnEdit_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is Position pos)
            {
                _editingPosition = pos;
                txtPositionName.Text = pos.Name;
                txtPositionName.Focus();

                if (string.IsNullOrEmpty(pos.DepartmentName))
                {
                    cboDepartment.SelectedIndex = 0;
                }
                else
                {
                    bool found = false;
                    foreach (var item in cboDepartment.Items)
                    {
                        if (item?.ToString()?.Equals(pos.DepartmentName, System.StringComparison.OrdinalIgnoreCase) == true)
                        {
                            cboDepartment.SelectedItem = item;
                            found = true;
                            break;
                        }
                    }
                    if (!found)
                    {
                        if (cboDepartment.ItemsSource is List<string> deptsList)
                        {
                            var newList = new List<string>(deptsList) { pos.DepartmentName };
                            cboDepartment.ItemsSource = newList;
                            cboDepartment.SelectedItem = pos.DepartmentName;
                        }
                    }
                }

                // Change button to indicate Update
                btnAdd.Content = new MaterialDesignThemes.Wpf.PackIcon { Kind = MaterialDesignThemes.Wpf.PackIconKind.ContentSave, Width = 24, Height = 24 };
                btnAdd.ToolTip = "Lưu thay đổi";
                btnAdd.Background = System.Windows.Media.Brushes.Orange;
            }
        }

        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is Position pos)
            {
                var confirm = new ConfirmWindow($"Bạn có chắc muốn xóa chức vụ '{pos.Name}'?", "Xác nhận xóa");
                if (confirm.ShowDialog() == true)
                {
                    using (var context = new AppDbContext())
                    {
                        var p = context.Positions.Find(pos.Id);
                        if (p != null)
                        {
                            context.Positions.Remove(p);
                            context.SaveChanges();
                        }
                    }
                    LoadPositions();
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
            if (lstPositions.SelectedItem is Position selected)
            {
                SelectedPosition = selected.Name;
                DialogResult = true;
                Close();
            }
        }
    }
}
