using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Microsoft.EntityFrameworkCore;
using TaxPersonnelManagement.Data;
using TaxPersonnelManagement.Models;

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
            try
            {
                InitializeComponent();
                ApplyAuthorization();
                LoadDepartments();
                LoadData();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                if (ex.InnerException != null) msg += "\nInner: " + ex.InnerException.Message;
                MessageBox.Show("Lỗi khởi tạo PersonnelListView:\n" + msg, "Lỗi Debug", MessageBoxButton.OK, MessageBoxImage.Error);
                throw;
            }
        }

        private void Grid_PreviewMouseWheel(object sender, System.Windows.Input.MouseWheelEventArgs e)
        {
            if (cbDepartmentFilter.IsDropDownOpen)
            {
                return;
            }

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

        private void ApplyAuthorization()
        {
            if (App.CurrentUser?.Role == UserRole.Staff)
            {
                btnAdd.Visibility = Visibility.Collapsed;
                btnImportExcel.Visibility = Visibility.Collapsed;
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
            var deptOrder = new System.Collections.Generic.List<string> {
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

            using (var context = new AppDbContext())
            {
                var allDepts = context.Departments
                                      .Select(d => d.Name)
                                      .Where(x => !string.IsNullOrEmpty(x))
                                      .Distinct()
                                      .ToList()
                                      .OrderBy(x => {
                                          int idx = deptOrder.FindIndex(d => d.Equals(x, StringComparison.OrdinalIgnoreCase));
                                          return idx == -1 ? 999 : idx;
                                      })
                                      .ThenBy(x => x)
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
                var query = context.Personnel
                                   .Include(p => p.LeaveHistories)
                                   .Include(p => p.SalaryRecords)
                                   .AsQueryable();

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
                                        if (pos.Contains("trưởng") && !pos.Contains("phó") && !pos.Contains("quyền")) return 1;
                                        if (pos.Contains("quyền")) return 2;
                                        if (pos.Contains("phó")) return 3;
                                    }
                                    else
                                    {
                                        if (pos.Contains("tổ trưởng") && !pos.Contains("phó")) return 1;
                                        if (pos.Contains("phó")) return 2;
                                        if (pos.Contains("công chức")) return 3;
                                    }
                                    return 99;
                                })
                                .ThenBy(p => p.FullName)
                                .ToList();

                txtTotalCount.Text = list.Count.ToString();

                var now = DateTime.Now.Date;
                
                int maternity = list.Count(p => p.LeaveHistories != null && 
                                                p.LeaveHistories.Any(l => (l.LeaveType == "Thai sản" || l.LeaveType == "Nghỉ thai sản") && 
                                                                          l.StartDate <= now && (l.EndDate == null || l.EndDate >= now)));
                txtMaternityCount.Text = maternity.ToString();

                int sick = list.Count(p => p.LeaveHistories != null && 
                                           p.LeaveHistories.Any(l => (l.LeaveType == "Nghỉ ốm" || l.LeaveType.Contains("ốm")) && 
                                                                      l.StartDate <= now && (l.EndDate == null || l.EndDate >= now)));
                txtSickCount.Text = sick.ToString();

                var displayList = list.AsEnumerable();
                if (_currentCardFilter == "Maternity")
                {
                    displayList = displayList.Where(p => p.LeaveHistories != null && 
                                                    p.LeaveHistories.Any(l => (l.LeaveType == "Thai sản" || l.LeaveType == "Nghỉ thai sản") && 
                                                                              l.StartDate <= now && (l.EndDate == null || l.EndDate >= now)));
                }
                else if (_currentCardFilter == "Sick")
                {
                    displayList = displayList.Where(p => p.LeaveHistories != null && 
                                                    p.LeaveHistories.Any(l => (l.LeaveType == "Nghỉ ốm" || l.LeaveType.Contains("ốm")) && 
                                                                              l.StartDate <= now && (l.EndDate == null || l.EndDate >= now)));
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

            int startIndex = (_currentPage - 1) * PageSize;
            for (int i = 0; i < pageItems.Count; i++)
            {
                pageItems[i].STT = startIndex + i + 1;
            }

            dgPersonnel.ItemsSource = pageItems;

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

        private void UpdateCardVisuals()
        {
            cardTotal.Opacity = _currentCardFilter == "All" ? 1.0 : 0.6;
            cardMaternity.Opacity = _currentCardFilter == "Maternity" ? 1.0 : 0.6;
            cardSick.Opacity = _currentCardFilter == "Sick" ? 1.0 : 0.6;
        }

        private void cardTotal_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            _currentCardFilter = "All";
            UpdateCardVisuals();
            LoadData(txtSearch.Text);
        }

        private void cardMaternity_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            _currentCardFilter = "Maternity";
            UpdateCardVisuals();
            LoadData(txtSearch.Text);
        }

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

        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
             if (Application.Current.MainWindow is MainWindow mw)
             {
                 mw.NavigateToPersonnelDetail(null);
             }
        }

        private void btnEdit_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is int id)
            {
                using (var context = new AppDbContext())
                {
                    var p = context.Personnel.Include(p => p.LeaveHistories).Include(p => p.SalaryRecords).FirstOrDefault(x => x.Id == id);
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

        private void btnView_Click(object sender, RoutedEventArgs e)
        {
             if (sender is Button btn && btn.Tag is int id)
             {
                  using (var context = new AppDbContext())
                  {
                      var p = context.Personnel.Include(p => p.LeaveHistories).Include(p => p.SalaryRecords).FirstOrDefault(x => x.Id == id);
                      if (p != null)
                      {
                          var dialog = new PersonnelProfileDialog(p);
                          dialog.ShowDialog();
                      }
                  }
             }
        }

        private void btnExportExcel_Click(object sender, RoutedEventArgs e)
        {
             if (_fullFilteredList != null && _fullFilteredList.Any())
             {
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
                          var success = new SuccessWindow("Xuất Excel thành công!", null, dlg.FileName, true);
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

        private async void btnImportExcel_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.DefaultExt = ".xlsx";
            dlg.Filter = "Excel Documents (.xlsx)|*.xlsx";

            if (dlg.ShowDialog() == true)
            {
                try
                {
                    // Hiện hiệu ứng loading
                    LoadingOverlay.Visibility = Visibility.Visible;

                    // Đọc file Excel ở luồng nền (background thread)
                    var importedData = await Task.Run(() => TaxPersonnelManagement.Services.ExcelImporter.Import(dlg.FileName));

                    // Ẩn hiệu ứng loading
                    LoadingOverlay.Visibility = Visibility.Collapsed;

                    if (importedData != null && importedData.Any())
                    {
                        // Lấy danh sách CCCD đã tồn tại trong DB để kiểm tra trùng
                        List<string> existingIds;
                        using (var context = new AppDbContext())
                        {
                            existingIds = context.Personnel
                                .Where(p => !string.IsNullOrEmpty(p.IdentityCardNumber))
                                .Select(p => p.IdentityCardNumber!)
                                .ToList();
                        }

                        // Lọc những người chưa có trong DB (dựa trên CCCD)
                        var newData = importedData.Where(p => 
                            string.IsNullOrEmpty(p.IdentityCardNumber) || 
                            !existingIds.Contains(p.IdentityCardNumber)
                        ).ToList();

                        int skippedCount = importedData.Count - newData.Count;

                        if (newData.Count == 0)
                        {
                            var info = new SuccessWindow($"Toàn bộ {importedData.Count} người trong file đã có trong hệ thống.", "Không có dữ liệu mới");
                            if (Window.GetWindow(this) is Window p1) info.Owner = p1;
                            info.ShowDialog();
                            return;
                        }

                        var confirmMsg = $"Tìm thấy {importedData.Count} nhân sự. Trong đó:\n" +
                                        $"- {newData.Count} nhân sự mới\n" +
                                        $"- {skippedCount} nhân sự đã tồn tại (sẽ bỏ qua)\n\n" +
                                        "Bạn có muốn tiếp tục nhập không?";

                        var confirmDialog = new ConfirmDialog(confirmMsg);
                        if (confirmDialog.ShowDialog() == true)
                        {
                            LoadingOverlay.Visibility = Visibility.Visible;
                            
                            await Task.Run(() => {
                                using (var context = new AppDbContext())
                                {
                                    context.Personnel.AddRange(newData);
                                    context.SaveChanges();
                                }
                            });

                            LoadingOverlay.Visibility = Visibility.Collapsed;

                            LoadData();
                            LoadDepartments();

                            var success = new SuccessWindow($"Đã nhập thành công {newData.Count} nhân sự mới!", skippedCount > 0 ? $"Đã bỏ qua {skippedCount} người trùng." : "");
                            if (Window.GetWindow(this) is Window parent) success.Owner = parent;
                            success.ShowDialog();
                        }
                    }
                    else
                    {
                        MessageBox.Show("Không tìm thấy dữ liệu hợp lệ trong file Excel.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
                catch (Exception ex)
                {
                    LoadingOverlay.Visibility = Visibility.Collapsed;
                    MessageBox.Show("Lỗi nhập Excel: " + ex.Message, "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
    }
}
