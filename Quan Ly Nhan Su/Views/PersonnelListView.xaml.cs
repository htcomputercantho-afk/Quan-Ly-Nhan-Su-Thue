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
        /// <summary>
        /// Bộ lọc thẻ Dashboard hiện tại: "All" = Tất cả, "Maternity" = Nghỉ thai sản, "Sick" = Nghỉ ốm.
        /// Khi người dùng bấm vào thẻ trên Dashboard, giá trị này thay đổi và danh sách sẽ được lọc tương ứng.
        /// </summary>
        private string _currentCardFilter = "All";

        public PersonnelListView()
        {
            InitializeComponent();
            ApplyAuthorization();
            LoadDepartments();
            LoadData();
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

        /// <summary>
        /// Tải danh sách phòng ban từ CSDL để hiển thị trong bộ lọc ComboBox.
        /// Bao gồm cả phòng ban từ bảng Departments và phòng ban đang dùng trong bảng Personnel.
        /// </summary>
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

        /// <summary>
        /// Tải danh sách nhân sự từ CSDL, áp dụng bộ lọc tìm kiếm, phòng ban, và thẻ Dashboard.
        /// Đồng thời cập nhật số liệu thống kê trên các thẻ Dashboard (tổng, thai sản, nghỉ ốm).
        /// </summary>
        private void LoadData(string search = "")
        {
            using (var context = new AppDbContext())
            {
                var query = context.Personnel.Include("LeaveHistories").AsQueryable();

                // 1. Lọc theo từ khóa tìm kiếm (Họ tên hoặc Phòng ban)
                if (!string.IsNullOrEmpty(search))
                {
                    string s = search.ToLower();
                    query = query.Where(p => (p.FullName != null && p.FullName.ToLower().Contains(s)) || 
                                             (p.Department != null && p.Department.ToLower().Contains(s)));
                }

                // 2. Lọc theo Phòng ban được chọn
                if (cbDepartmentFilter.SelectedItem is string dept && dept != "-- Tất cả bộ phận --")
                {
                    query = query.Where(p => p.Department == dept);
                }
                
                // Thứ tự sắp xếp phòng ban theo cấu trúc tổ chức
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

                // Sắp xếp theo phòng ban → chức vụ
                var list = query.AsEnumerable()
                                .OrderBy(p => {
                                    string dept2 = (p.Department ?? "").Trim();
                                    int index = deptOrder.FindIndex(d => d.Equals(dept2, StringComparison.OrdinalIgnoreCase));
                                    return index == -1 ? 999 : index;
                                })
                                .ThenBy(p => {
                                    string pos = p.Position?.ToLower() ?? "";
                                    if (pos.Contains("chi cục trưởng") && !pos.Contains("phó")) return 1;
                                    if (pos.Contains("phó")) return 2;
                                    if (pos.Contains("tổ trưởng") || pos.Contains("đội trưởng")) return 3;
                                    return 4;
                                })
                                .ToList();

                // Đánh số thứ tự
                for (int i = 0; i < list.Count; i++)
                {
                    list[i].STT = i + 1;
                }

                // 3. Cập nhật số liệu thống kê trên các thẻ Dashboard (luôn tính trên toàn bộ danh sách)
                txtTotalCount.Text = list.Count.ToString();

                var now = DateTime.Now.Date;
                
                // Đếm số công chức đang nghỉ thai sản
                int maternity = list.Count(p => p.LeaveHistories != null && 
                                                p.LeaveHistories.Any(l => (l.LeaveType == "Thai sản" || l.LeaveType == "Nghỉ thai sản") && 
                                                                          l.StartDate <= now && l.EndDate >= now));
                txtMaternityCount.Text = maternity.ToString();

                // Đếm số công chức đang nghỉ ốm
                int sick = list.Count(p => p.LeaveHistories != null && 
                                           p.LeaveHistories.Any(l => (l.LeaveType == "Nghỉ ốm" || l.LeaveType.Contains("ốm")) && 
                                                                     l.StartDate <= now && l.EndDate >= now));
                txtSickCount.Text = sick.ToString();

                // 4. Áp dụng bộ lọc thẻ Dashboard để chỉ hiển thị nhóm công chức được chọn
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

                dgPersonnel.ItemsSource = displayList.ToList();
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
             if (dgPersonnel.ItemsSource is System.Collections.Generic.IEnumerable<TaxPersonnelManagement.Models.Personnel> list)
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
                          TaxPersonnelManagement.Services.ExcelExporter.Export(list, dlg.FileName);
                          
                          var success = new SuccessWindow("Xuất Excel thành công!");
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
