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
        private readonly bool _isPlanningMode;
        private bool _isInitializing = true;

        public AddPositionDialog(bool isPartyMode = false, bool isPlanningMode = false)
        {
            InitializeComponent();
            _isPartyMode = isPartyMode;
            _isPlanningMode = isPlanningMode;

            if (_isPartyMode)
            {
                txtTitle.Text = "Quản lý Chức danh Đảng";
                MaterialDesignThemes.Wpf.HintAssist.SetHint(txtPositionName, "Nhập tên chức danh Đảng mới...");
                btnAdd.ToolTip = "Thêm chức danh";
            }
            else if (_isPlanningMode)
            {
                txtTitle.Text = "Quản lý Chức danh Quy hoạch";
                MaterialDesignThemes.Wpf.HintAssist.SetHint(txtPositionName, "Nhập tên chức danh quy hoạch mới...");
                btnAdd.ToolTip = "Thêm chức danh quy hoạch";
            }
            else
            {
                pnlGroupType.Visibility = Visibility.Visible;
            }

            LoadDepartments();
            LoadPositions();
            _isInitializing = false;
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
                try { context.Database.EnsureCreated(); } catch { }

                // 1. Tự động dọn dẹp các bản ghi "test"
                var testPositions = context.Positions.Where(x => x.Name.ToLower() == "test").ToList();
                if (testPositions.Any())
                {
                    context.Positions.RemoveRange(testPositions);
                    try { context.SaveChanges(); } catch { }
                }

                // 2. Migration một lần: Chuyển các chức danh quy hoạch cũ (__QUY_HOẠCH__) sang bảng PlanningPositions
                var oldPlanningPositions = context.Positions.Where(p => p.DepartmentName == "__QUY_HOẠCH__").ToList();
                if (oldPlanningPositions.Any())
                {
                    foreach (var oldP in oldPlanningPositions)
                    {
                        if (!context.PlanningPositions.Any(x => x.Name.ToLower() == oldP.Name.ToLower()))
                        {
                            context.PlanningPositions.Add(new PlanningPosition { Name = oldP.Name });
                        }
                    }
                    context.Positions.RemoveRange(oldPlanningPositions);
                    try { context.SaveChanges(); } catch { }
                }

                // 3. Migration một lần: Gán GroupType mặc định cho các chức vụ cũ nếu GroupType đang null
                var leadershipNames = new HashSet<string>(System.StringComparer.OrdinalIgnoreCase) {
                    "Chi cục trưởng", "Quyền Chi cục trưởng", "Phó Chi cục trưởng",
                    "Trưởng Thuế cơ sở", "Quyền Trưởng Thuế cơ sở", "Phó Trưởng Thuế cơ sở",
                    "Cục trưởng", "Phó Cục trưởng"
                };
                var unitNames = new HashSet<string>(System.StringComparer.OrdinalIgnoreCase) {
                    "Đội trưởng", "Trưởng phòng", "Phó Đội trưởng", "Phó Trưởng phòng",
                    "Tổ trưởng", "Phó Tổ trưởng", "Công chức", "Nhân viên"
                };

                var unclassified = context.Positions.Where(p => p.DepartmentName != "__ĐẢNG__" && string.IsNullOrEmpty(p.GroupType)).ToList();
                if (unclassified.Any())
                {
                    foreach (var p in unclassified)
                    {
                        if (leadershipNames.Contains(p.Name)) p.GroupType = "Ban lãnh đạo";
                        else if (unitNames.Contains(p.Name)) p.GroupType = "Các tổ";
                    }
                    try { context.SaveChanges(); } catch { }
                }

                List<Position> list;
                if (_isPartyMode)
                {
                    list = context.Positions
                                  .Where(x => x.DepartmentName == "__ĐẢNG__")
                                  .OrderBy(x => x.Name)
                                  .ToList();
                }
                else if (_isPlanningMode)
                {
                    var dbPlanning = context.PlanningPositions.OrderBy(x => x.Name).ToList();
                    list = dbPlanning.Select(p => new Position { Id = p.Id, Name = p.Name, DepartmentName = "__QUY_HOẠCH__" }).ToList();
                }
                else
                {
                    var posOrder = new List<string> {
                        "Chi cục trưởng", "Quyền Chi cục trưởng", "Phó Chi cục trưởng",
                        "Trưởng Thuế cơ sở", "Quyền Trưởng Thuế cơ sở", "Phó Trưởng Thuế cơ sở",
                        "Cục trưởng", "Phó Cục trưởng",
                        "Đội trưởng", "Trưởng phòng", "Phó Đội trưởng", "Phó Trưởng phòng",
                        "Tổ trưởng", "Phó Tổ trưởng", "Công chức", "Nhân viên"
                    };

                    var rawList = context.Positions.ToList()
                                      .Where(x => x.DepartmentName != "__ĐẢNG__" && x.DepartmentName != "__QUY_HOẠCH__" && !string.IsNullOrEmpty(x.Name));

                    if (rdoGroupLeadership != null && rdoGroupLeadership.IsChecked == true)
                    {
                        rawList = rawList.Where(x => x.GroupType == "Ban lãnh đạo");
                    }
                    else if (rdoGroupUnits != null && rdoGroupUnits.IsChecked == true)
                    {
                        rawList = rawList.Where(x => x.GroupType == "Các Tổ" || x.GroupType == "Các tổ");
                    }

                    list = rawList.OrderBy(x =>
                                      {
                                          if (x.GroupType == "Ban lãnh đạo") return 0;
                                          if (x.GroupType == "Các Tổ" || x.GroupType == "Các tổ") return 1;
                                          return 2;
                                      })
                                  .ThenBy(x =>
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

        private void rdoGroup_Checked(object sender, RoutedEventArgs e)
        {
            if (_isInitializing) return;
            LoadPositions();
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

            string? selectedGroup = null;
            if (!_isPartyMode && !_isPlanningMode)
            {
                if (rdoGroupLeadership.IsChecked == true) selectedGroup = "Ban lãnh đạo";
                else if (rdoGroupUnits.IsChecked == true) selectedGroup = "Các Tổ";
            }

            using (var context = new AppDbContext())
            {
                if (_isPlanningMode)
                {
                    if (_editingPosition == null)
                    {
                        bool exists = context.PlanningPositions.AsEnumerable().Any(d => d.Name.Equals(newName, System.StringComparison.OrdinalIgnoreCase));
                        if (exists)
                        {
                            new WarningWindow("Chức danh quy hoạch này đã tồn tại!", "Thông báo").ShowDialog();
                            return;
                        }
                        context.PlanningPositions.Add(new PlanningPosition { Name = newName });
                    }
                    else
                    {
                        var item = context.PlanningPositions.Find(_editingPosition.Id);
                        if (item != null) item.Name = newName;
                        _editingPosition = null;
                    }
                }
                else
                {
                    string? selectedDept = _isPartyMode ? "__ĐẢNG__" : null;

                    if (_editingPosition == null)
                    {
                        bool exists = context.Positions.AsEnumerable().Any(d => 
                            d.Name.Equals(newName, System.StringComparison.OrdinalIgnoreCase) && 
                            (d.DepartmentName ?? "").Equals(selectedDept ?? "", System.StringComparison.OrdinalIgnoreCase));
                        if (exists)
                        {
                            new WarningWindow(_isPartyMode ? "Chức danh này đã tồn tại!" : "Chức vụ này đã tồn tại!", "Thông báo").ShowDialog();
                            return;
                        }

                        var newPos = new Position { Name = newName, DepartmentName = selectedDept, GroupType = selectedGroup };
                        context.Positions.Add(newPos);
                    }
                    else
                    {
                        var p = context.Positions.Find(_editingPosition.Id);
                        if (p != null)
                        {
                            p.Name = newName;
                            p.DepartmentName = selectedDept;
                            p.GroupType = selectedGroup;
                        }
                        _editingPosition = null;
                    }
                }

                context.SaveChanges();
            }

            txtPositionName.Clear();
            rdoGroupAll.IsChecked = true;
            txtPositionName.Focus();

            // Reset button visual
            btnAdd.Background = (System.Windows.Media.Brush)Application.Current.Resources["PrimaryHueMidBrush"];
            btnAdd.Content = new MaterialDesignThemes.Wpf.PackIcon { Kind = MaterialDesignThemes.Wpf.PackIconKind.Plus, Width = 24, Height = 24 };
            btnAdd.ToolTip = _isPlanningMode ? "Thêm chức danh quy hoạch" : "Thêm chức vụ";

            LoadPositions();
        }

        private void btnEdit_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is Position pos)
            {
                _editingPosition = pos;
                txtPositionName.Text = pos.Name;
                txtPositionName.Focus();

                if (pos.GroupType == "Ban lãnh đạo") rdoGroupLeadership.IsChecked = true;
                else if (pos.GroupType == "Các Tổ" || pos.GroupType == "Các tổ") rdoGroupUnits.IsChecked = true;
                else rdoGroupAll.IsChecked = true;

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
                var msg = _isPartyMode ? $"Bạn có chắc muốn xóa chức danh '{pos.Name}'?" : 
                          (_isPlanningMode ? $"Bạn có chắc muốn xóa chức danh quy hoạch '{pos.Name}'?" : $"Bạn có chắc muốn xóa chức vụ '{pos.Name}'?");
                var confirm = new ConfirmWindow(msg, "Xác nhận xóa");
                if (confirm.ShowDialog() == true)
                {
                    using (var context = new AppDbContext())
                    {
                        if (_isPlanningMode)
                        {
                            var item = context.PlanningPositions.Find(pos.Id);
                            if (item != null)
                            {
                                context.PlanningPositions.Remove(item);
                                context.SaveChanges();
                            }
                        }
                        else
                        {
                            var p = context.Positions.Find(pos.Id);
                            if (p != null)
                            {
                                context.Positions.Remove(p);
                                context.SaveChanges();
                            }
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
