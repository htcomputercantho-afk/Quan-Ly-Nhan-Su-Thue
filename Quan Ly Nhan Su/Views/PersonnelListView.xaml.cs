using System.Linq;
using System.Windows;
using System.Windows.Controls;
using TaxPersonnelManagement.Data;
using TaxPersonnelManagement.Models;
using System.Collections.ObjectModel;
using Microsoft.EntityFrameworkCore;

namespace TaxPersonnelManagement.Views
{
    public partial class PersonnelListView : UserControl
    {
        private string _currentCardFilter = "All";
        
        // Pagination
        private List<Personnel> _fullFilteredList = new List<Personnel>();
        private int _currentPage = 1;
        private const int PageSize = 20;
        private int _totalPages = 1;

        public PersonnelListView()
        {
            InitializeComponent();
            ApplyAuthorization();
            LoadDepartments();
            LoadData();
        }

        private void Grid_PreviewMouseWheel(object sender, System.Windows.Input.MouseWheelEventArgs e)
        {
            if (!e.Handled)
            {
                e.Handled = true;
                var eventArg = new System.Windows.Input.MouseWheelEventArgs(e.MouseDevice, e.Timestamp, e.Delta);
                eventArg.RoutedEvent = UIElement.MouseWheelEvent;
                eventArg.Source = sender;
                var parent = ((FrameworkElement)sender).Parent as UIElement;
                parent?.RaiseEvent(eventArg);
            }
        }

        /// <summary>
        /// Ẩn nút Thêm mới nếu người dùng không phải Admin.
        /// </summary>
        private void ApplyAuthorization()
        {
            if (App.CurrentUser?.Role == UserRole.Staff)
            {
                btnAdd.Visibility = Visibility.Collapsed;
            }
        }

        private void AdminOnly_Loaded(object sender, RoutedEventArgs e)
        {
            if (App.CurrentUser?.Role == UserRole.Staff && sender is FrameworkElement element)
            {
                element.Visibility = Visibility.Collapsed;
            }
        }

        private void LoadDepartments()
        {
            using (var context = new AppDbContext())
            {
                var dbDepts = context.Departments.Select(d => d.Name).ToList();
                var usedDepts = context.Personnel
                                   .Where(p => !string.IsNullOrEmpty(p.Department))
                                   .Select(p => p.Department)
                                   .Distinct()
                                   .ToList();

                var allDepts = dbDepts.Union(usedDepts!)
                                      .Where(x => !string.IsNullOrEmpty(x))
                                      .Distinct()
                                      .OrderBy(d => d)
                                      .ToList();
                
                allDepts.Insert(0, "-- Tất cả bộ phận --");
                cbDepartmentFilter.ItemsSource = allDepts;
                cbDepartmentFilter.SelectedIndex = 0;
            }
        }

        private void LoadData(string search = "")
        {
            using (var context = new AppDbContext())
            {
                var query = context.Personnel.Include("LeaveHistories").AsQueryable();

                if (!string.IsNullOrEmpty(search))
                {
                    string s = search.ToLower();
                    query = query.Where(p => (p.FullName != null && p.FullName.ToLower().Contains(s)) || 
                                             (p.Department != null && p.Department.ToLower().Contains(s)));
                }

                if (cbDepartmentFilter.SelectedItem is string dept && dept != "-- Tất cả bộ phận --")
                {
                    query = query.Where(p => p.Department == dept);
                }
                
                var deptOrder = new List<string> {
                    "Ban lãnh đạo",
                    "Tổ Hành chính, tổng hợp",
                    "Tổ Kiểm tra số 1",
                    "Tổ Kiểm tra số 2",
                    "Tổ Kiểm tra số 3",
                    "Tổ Nghiệp vụ, dự toán, pháp chế",
                    "Tổ Quản lý các khoản thu khác",
                    "Tổ Quản lý, hỗ trợ cá nhân, hộ kinh doanh số 1",
                    "Tổ Quản lý, hỗ trợ cá nhân, hộ kinh doanh số 2",
                    "Tổ Quản lý, hỗ trợ doanh nghiệp số 1",
                    "Tổ Quản lý, hỗ trợ doanh nghiệp số 2"
                };

                var list = query.AsEnumerable()
                                .OrderBy(p => {
                                    string dept2 = (p.Department ?? "").Trim();
                                    int index = deptOrder.FindIndex(d => d.Equals(dept2, StringComparison.OrdinalIgnoreCase));
                                    return index == -1 ? 999 : index;
                                })
                                .ThenBy(p => {
                                    string pos = p.Position?.ToLower() ?? "";
                                    string dept3 = (p.Department ?? "").ToLower();

                                    if (dept3.Contains("lãnh đạo"))
                                    {
                                        // 1. Trưởng Thuế cơ sở, 2. Quyền Trưởng Thuế cơ sở, 3. Phó Trưởng Thuế cơ sở
                                        if (pos.Contains("trưởng") && !pos.Contains("phó") && !pos.Contains("quyền")) return 1;
                                        if (pos.Contains("quyền")) return 2;
                                        if (pos.Contains("phó")) return 3;
                                    }
                                    else
                                    {
                                        // 1. Tổ trưởng, 2. Phó Tổ trưởng, 3. Công chức
                                        if (pos.Contains("tổ trưởng") && !pos.Contains("phó")) return 1;
                                        if (pos.Contains("phó")) return 2;
                                        if (pos.Contains("công chức")) return 3;
                                    }
                                    return 99;
                                })
                                .ThenBy(p => p.FullName)
                                .ToList();

                // Dashboard stats (always on full list)
                txtTotalCount.Text = list.Count.ToString();

                var now = DateTime.Now.Date;
                
                int maternity = list.Count(p => p.LeaveHistories != null && 
                                                p.LeaveHistories.Any(l => (l.LeaveType == "Thai sản" || l.LeaveType == "Nghỉ thai sản") && 
                                                                          l.StartDate <= now && l.EndDate >= now));
                txtMaternityCount.Text = maternity.ToString();

                int sick = list.Count(p => p.LeaveHistories != null && 
                                           p.LeaveHistories.Any(l => (l.LeaveType == "Nghỉ ốm" || l.LeaveType.Contains("ốm")) && 
                                                                     l.StartDate <= now && l.EndDate >= now));
                txtSickCount.Text = sick.ToString();

                // Apply card filter
                var displayList = list.AsEnumerable();
                if (_currentCardFilter == "Maternity")
                {
                    displayList = displayList.Where(p => p.LeaveHistories != null && 
                                                    p.LeaveHistories.Any(l => (l.LeaveType == "Thai sản" || l.LeaveType == "Nghỉ thai sản") && 
                                                                              l.StartDate <= now && l.EndDate >= now));
                }
                else if (_currentCardFilter == "Sick")
                {
                    displayList = displayList.Where(p => p.LeaveHistories != null && 
                                                    p.LeaveHistories.Any(l => (l.LeaveType == "Nghỉ ốm" || l.LeaveType.Contains("ốm")) && 
                                                                              l.StartDate <= now && l.EndDate >= now));
                }

                _fullFilteredList = displayList.ToList();
                _currentPage = 1;
                ApplyPagination();
            }
        }

        private void ApplyPagination()
        {
            int totalItems = _fullFilteredList.Count;
            _totalPages = Math.Max(1, (int)Math.Ceiling((double)totalItems / PageSize));
            
            if (_currentPage > _totalPages) _currentPage = _totalPages;
            if (_currentPage < 1) _currentPage = 1;

            var pageItems = _fullFilteredList
                .Skip((_currentPage - 1) * PageSize)
                .Take(PageSize)
                .ToList();

            // Re-number STT across pages
            int startIndex = (_currentPage - 1) * PageSize;
            for (int i = 0; i < pageItems.Count; i++)
            {
                pageItems[i].STT = startIndex + i + 1;
            }

            dgPersonnel.ItemsSource = pageItems;

            // Update footer
            if (totalItems == 0)
            {
                txtPagingInfo.Text = "Không có dữ liệu";
                txtPageInfo.Text = "0 / 0";
            }
            else
            {
                int from = startIndex + 1;
                int to = startIndex + pageItems.Count;
                txtPagingInfo.Text = $"Hiển thị {from} - {to} trên {totalItems} nhân viên";
                txtPageInfo.Text = $"{_currentPage} / {_totalPages}";
            }

            btnPrevPage.IsEnabled = _currentPage > 1;
            btnNextPage.IsEnabled = _currentPage < _totalPages;
        }

        private void btnPrevPage_Click(object sender, RoutedEventArgs e)
        {
            if (_currentPage > 1)
            {
                _currentPage--;
                ApplyPagination();
            }
        }

        private void btnNextPage_Click(object sender, RoutedEventArgs e)
        {
            if (_currentPage < _totalPages)
            {
                _currentPage++;
                ApplyPagination();
            }
        }

        private void cbDepartmentFilter_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            LoadData(txtSearch.Text);
        }

        /// <summary>
        /// Cập nhật hiệu ứng trực quan của các thẻ Dashboard.
        /// Thẻ đang được chọn sẽ sáng (Opacity = 1.0), các thẻ khác sẽ mờ (Opacity = 0.6).
        /// </summary>
        private void UpdateCardVisuals()
        {
            cardTotal.Opacity = _currentCardFilter == "All" ? 1.0 : 0.6;
            cardMaternity.Opacity = _currentCardFilter == "Maternity" ? 1.0 : 0.6;
            cardSick.Opacity = _currentCardFilter == "Sick" ? 1.0 : 0.6;
        }

        /// <summary>Bấm thẻ "Tổng nhân sự" → hiển thị toàn bộ danh sách.</summary>
        private void cardTotal_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            _currentCardFilter = "All";
            UpdateCardVisuals();
            LoadData(txtSearch.Text);
        }

        /// <summary>Bấm thẻ "Nghỉ thai sản" → chỉ hiển thị công chức đang nghỉ thai sản.</summary>
        private void cardMaternity_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            _currentCardFilter = "Maternity";
            UpdateCardVisuals();
            LoadData(txtSearch.Text);
        }

        /// <summary>Bấm thẻ "Nghỉ ốm" → chỉ hiển thị công chức đang nghỉ ốm.</summary>
        private void cardSick_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            _currentCardFilter = "Sick";
            UpdateCardVisuals();
            LoadData(txtSearch.Text);
        }

        private void txtSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            LoadData(txtSearch.Text);
        }

        /// <summary>Thêm mới công chức → mở trang chi tiết với hồ sơ trống.</summary>
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
             if (Application.Current.MainWindow is MainWindow mw)
             {
                 mw.NavigateToPersonnelDetail(null);
             }
        }

        /// <summary>Chỉnh sửa hồ sơ công chức → mở trang chi tiết với dữ liệu hiện có.</summary>
        private void btnEdit_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is int id)
            {
                using (var context = new AppDbContext())
                {
                    var p = context.Personnel.Include("LeaveHistories").FirstOrDefault(x => x.Id == id);
                    if (p != null)
                    {
                        if (Application.Current.MainWindow is MainWindow mw)
                        {
                            mw.NavigateToPersonnelDetail(p);
                        }
                    }
                }
            }
        }

        /// <summary>Xóa công chức sau khi xác nhận.</summary>
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is int id)
            {
                var dialog = new ConfirmDialog("Bạn có chắc muốn xóa nhân sự này?");
                if (dialog.ShowDialog() == true)
                {
                    using (var context = new AppDbContext())
                    {
                        var p = context.Personnel.Find(id);
                        if (p != null)
                        {
                            context.Personnel.Remove(p);
                            context.SaveChanges();
                            LoadData();
                            LoadDepartments();
                        }
                    }
                }
            }
        }

        /// <summary>Xem hồ sơ công chức dạng popup (profile dialog).</summary>
        private void btnView_Click(object sender, RoutedEventArgs e)
        {
             if (sender is Button btn && btn.Tag is int id)
             {
                 using (var context = new AppDbContext())
                 {
                     var p = context.Personnel.Include("LeaveHistories").FirstOrDefault(x => x.Id == id);
                     if (p != null)
                     {
                         var dialog = new PersonnelProfileDialog(p);
                         dialog.ShowDialog();
                     }
                 }
             }
        }

        /// <summary>
        /// Xuất danh sách nhân sự ra file Excel.
        /// Tên file tự động thay đổi theo bộ lọc thẻ Dashboard đang chọn (Thai sản / Ốm / Tất cả).
        /// Dữ liệu xuất ra chính là dữ liệu đang hiển thị trên bảng (đã qua lọc).
        /// </summary>
        private void btnExportExcel_Click(object sender, RoutedEventArgs e)
        {
             if (_fullFilteredList != null && _fullFilteredList.Any())
             {
                  // Tên file tự động theo bộ lọc: NghiThaiSan / NghiOm / NhanSu_Full
                  var dlg = new Microsoft.Win32.SaveFileDialog();
                  dlg.FileName = _currentCardFilter switch
                  {
                      "Maternity" => $"DanhSach_NghiThaiSan_{DateTime.Now:yyyyMMdd}",
                      "Sick" => $"DanhSach_NghiOm_{DateTime.Now:yyyyMMdd}",
                      _ => $"DanhSach_NhanSu_Full_{DateTime.Now:yyyyMMdd}"
                  };
                  dlg.DefaultExt = ".xlsx";
                  dlg.Filter = "Excel Documents (.xlsx)|*.xlsx";

                  if (dlg.ShowDialog() == true)
                  {
                      try 
                      {
                          TaxPersonnelManagement.Services.ExcelExporter.Export(_fullFilteredList, dlg.FileName);
                                                    var success = new SuccessWindow("Xuất Excel thành công!", dlg.FileName);
                           if (Window.GetWindow(this) is Window parent) success.Owner = parent;
                           success.ShowDialog();
                      }
                      catch (Exception ex)
                      {
                          MessageBox.Show("Lỗi xuất Excel: " + ex.Message, "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                      }
                  }
             }
        }
    }
}
