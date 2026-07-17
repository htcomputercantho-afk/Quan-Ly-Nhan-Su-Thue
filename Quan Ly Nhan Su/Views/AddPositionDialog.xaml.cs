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

        private readonly bool _isPartyMode;

        public AddPositionDialog(bool isPartyMode = false)
        {
            InitializeComponent();
            _isPartyMode = isPartyMode;
            if (_isPartyMode)
            {
                txtTitle.Text = "Quản lý Chức danh Đảng";
                MaterialDesignThemes.Wpf.HintAssist.SetHint(txtPositionName, "Nhập tên chức danh Đảng mới...");
                btnAdd.ToolTip = "Thêm chức danh";
            }
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
            using (var context = new AppDbContext())
            {
                // Ensure table exists just in case
                try { context.Database.EnsureCreated(); } catch { }

                List<Position> list;
                if (_isPartyMode)
                {
                    // Check if there are any Party positions, if not, seed defaults
                    bool hasPartyPos = context.Positions.Any(x => x.DepartmentName == "__ĐẢNG__");
                    if (!hasPartyPos)
                    {
                        var defaultPartyPositions = new List<string> {
                            "Bí thư Chi bộ",
                            "Phó Bí thư Chi bộ",
                            "Chi ủy viên",
                            "Bí thư Đảng bộ",
                            "Phó Bí thư Đảng bộ",
                            "Ủy viên Ban thường vụ Đảng bộ",
                            "Ủy viên Ban chấp hành Đảng bộ",
                            "Bí thư Đảng ủy",
                            "Phó Bí thư Đảng ủy",
                            "Ủy viên Ban thường vụ Đảng ủy",
                            "Ủy viên Ban chấp hành Đảng ủy",
                            "Tổ trưởng Tổ Đảng",
                            "Đảng viên"
                        };
                        foreach (var name in defaultPartyPositions.Distinct())
                        {
                            context.Positions.Add(new Position { Name = name, DepartmentName = "__ĐẢNG__" });
                        }
                        context.SaveChanges();
                    }

                    list = context.Positions
                                  .Where(x => x.DepartmentName == "__ĐẢNG__")
                                  .ToList()
                                  .OrderBy(x => x.Name)
                                  .ToList();
                }
                else
                {
                    var posOrder = new System.Collections.Generic.List<string> {
                        "Chi cục trưởng",
                        "Quyền Chi cục trưởng",
                        "Phó Chi cục trưởng",
                        "Trưởng Thuế cơ sở",
                        "Quyền Trưởng Thuế cơ sở",
                        "Phó Trưởng Thuế cơ sở",
                        "Đội trưởng",
                        "Trưởng phòng",
                        "Phó Đội trưởng",
                        "Phó Trưởng phòng",
                        "Tổ trưởng",
                        "Phó Tổ trưởng",
                        "Công chức",
                        "Nhân viên"
                    };

                    list = context.Positions.ToList()
                                      .Where(x => x.DepartmentName != "__ĐẢNG__" && !string.IsNullOrEmpty(x.Name))
                                      .OrderBy(x =>
                                      {
                                          int idx = posOrder.FindIndex(p => p.Equals(x.Name, System.StringComparison.OrdinalIgnoreCase));
                                          return idx == -1 ? 999 : idx;
                                      })
                                      .ThenBy(x => x.Name)
                                      .ToList();
                }

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
                new WarningWindow(_isPartyMode ? "Vui lòng nhập tên chức danh!" : "Vui lòng nhập tên chức vụ!", "Thông báo").ShowDialog();
                return;
            }

            string? selectedDept = _isPartyMode ? "__ĐẢNG__" : cboDepartment.SelectedItem?.ToString();
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
                        new WarningWindow(_isPartyMode ? "Chức danh này đã tồn tại!" : "Chức vụ này thuộc bộ phận đã chọn đã tồn tại!", "Thông báo").ShowDialog();
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
                var msg = _isPartyMode ? $"Bạn có chắc muốn xóa chức danh '{pos.Name}'?" : $"Bạn có chắc muốn xóa chức vụ '{pos.Name}'?";
                var confirm = new ConfirmWindow(msg, "Xác nhận xóa");
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
