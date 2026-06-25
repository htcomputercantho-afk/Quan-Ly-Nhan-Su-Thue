using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using TaxPersonnelManagement.Data;

namespace TaxPersonnelManagement.Views
{
    public partial class PersonnelSelectorDialog : Window
    {
        private List<PersonnelSelectorItem> _allPersonnel = new();
        private bool _isSingleSelect = false;
        public List<int> SelectedPersonnelIds { get; private set; } = new();

        public PersonnelSelectorDialog(List<int> excludedIds, bool isSingleSelect = false)
        {
            _isSingleSelect = isSingleSelect;
            InitializeComponent();
            
            if (_isSingleSelect)
            {
                lblSubtitle.Text = "Chọn một cán bộ công chức từ danh sách dưới đây để tiếp tục.";
                btnConfirm.Content = "XÁC NHẬN";
                this.Loaded += PersonnelSelectorDialog_Loaded;
                dgPersonnelSelector.SelectionChanged += DgPersonnelSelector_SelectionChanged;
                dgPersonnelSelector.MouseDoubleClick += DgPersonnelSelector_MouseDoubleClick;
            }

            LoadDepartments();
            LoadPersonnel(excludedIds);
        }

        private void PersonnelSelectorDialog_Loaded(object sender, RoutedEventArgs e)
        {
            if (_isSingleSelect && dgPersonnelSelector.Columns.Count > 0)
            {
                dgPersonnelSelector.Columns[0].Visibility = Visibility.Collapsed;
            }
        }

        private void DgPersonnelSelector_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isSingleSelect)
            {
                UpdateCount();
            }
        }

        private void DgPersonnelSelector_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (_isSingleSelect && dgPersonnelSelector.SelectedItem is PersonnelSelectorItem selectedItem)
            {
                SelectedPersonnelIds = new List<int> { selectedItem.Id };
                DialogResult = true;
                Close();
            }
        }

        /// <summary>
        /// Tải danh sách phòng ban lên bộ lọc ComboBox
        /// </summary>
        private void LoadDepartments()
        {
            try
            {
                using (var db = new AppDbContext())
                {
                    var departments = db.Departments
                                        .OrderBy(d => d.Name)
                                        .Select(d => d.Name)
                                        .ToList();

                    var filterList = new List<string> { "Tất cả tổ/bộ phận" };
                    filterList.AddRange(departments);

                    cboDepartmentFilter.ItemsSource = filterList;
                    cboDepartmentFilter.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi tải danh mục tổ/bộ phận: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Tải danh sách cán bộ chưa tham gia lớp học
        /// </summary>
        private void LoadPersonnel(List<int> excludedIds)
        {
            try
            {
                using (var db = new AppDbContext())
                {
                    // Lấy toàn bộ cán bộ, loại trừ những người đã có trong lớp
                    var query = db.Personnel
                                  .Where(p => !excludedIds.Contains(p.Id))
                                  .OrderBy(p => p.Department)
                                  .ThenBy(p => p.FullName)
                                  .ToList();

                    _allPersonnel = query.Select(p => new PersonnelSelectorItem
                    {
                        Id = p.Id,
                        FullName = p.FullName,
                        IdentityCardNumber = p.IdentityCardNumber ?? "Không có",
                        Department = p.Department ?? "Không có bộ phận",
                        IsSelected = false
                    }).ToList();

                    ApplyFilter();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi tải danh sách cán bộ công chức: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Áp dụng bộ lọc tổ và từ khóa tìm kiếm
        /// </summary>
        private void ApplyFilter()
        {
            if (_allPersonnel == null) return;

            string selectedDept = cboDepartmentFilter.SelectedItem as string ?? "Tất cả tổ/bộ phận";
            string search = txtSearchQuery.Text.Trim();

            var filtered = _allPersonnel.AsEnumerable();

            if (selectedDept != "Tất cả tổ/bộ phận")
            {
                filtered = filtered.Where(p => p.Department == selectedDept);
            }

            if (!string.IsNullOrEmpty(search))
            {
                string searchUnsigned = RemoveSign(search);
                filtered = filtered.Where(p => p.FullName.Contains(search, StringComparison.OrdinalIgnoreCase) ||
                                               RemoveSign(p.FullName).Contains(searchUnsigned, StringComparison.OrdinalIgnoreCase) ||
                                               p.IdentityCardNumber.Contains(search, StringComparison.OrdinalIgnoreCase));
            }

            var list = filtered.ToList();

            // Đánh số thứ tự (STT) cho các mục hiển thị
            int stt = 1;
            foreach (var item in list)
            {
                item.STT = stt++;
            }

            dgPersonnelSelector.ItemsSource = null;
            dgPersonnelSelector.ItemsSource = list;

            UpdateCount();
        }

        /// <summary>
        /// Cập nhật số lượng đã chọn hiển thị trên UI và trạng thái CheckBox chọn tất cả
        /// </summary>
        private void UpdateCount()
        {
            if (_allPersonnel == null) return;

            if (_isSingleSelect)
            {
                var selected = dgPersonnelSelector.SelectedItem as PersonnelSelectorItem;
                var list = dgPersonnelSelector.ItemsSource as List<PersonnelSelectorItem>;
                int displayedCount = list?.Count ?? 0;
                
                lblSelectionCount.Text = selected != null 
                    ? $"Đang chọn: {selected.FullName} (Bộ phận: {selected.Department}) | Hiển thị: {displayedCount}"
                    : $"Chưa chọn cán bộ nào (Hiển thị: {displayedCount})";
            }
            else
            {
                int selectedCount = _allPersonnel.Count(p => p.IsSelected);
                var list = dgPersonnelSelector.ItemsSource as List<PersonnelSelectorItem>;
                int displayedCount = list?.Count ?? 0;

                lblSelectionCount.Text = $"Đã chọn: {selectedCount} / {_allPersonnel.Count} cán bộ (Hiển thị: {displayedCount})";

                // Đồng bộ trạng thái CheckBox chọn tất cả của danh sách đang hiển thị
                if (list != null && displayedCount > 0 && list.All(p => p.IsSelected))
                {
                    chkSelectAll.IsChecked = true;
                }
                else
                {
                    chkSelectAll.IsChecked = false;
                }
            }
        }

        /// <summary>
        /// Xử lý thay đổi bộ lọc hoặc từ khóa tìm kiếm
        /// </summary>
        private void Filter_Changed(object sender, object e)
        {
            ApplyFilter();
        }

        /// <summary>
        /// Xử lý kéo thả cửa sổ
        /// </summary>
        private void Header_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                DragMove();
            }
        }

        /// <summary>
        /// Tích chọn/Bỏ tích chọn toàn bộ cán bộ đang hiển thị
        /// </summary>
        private void chkSelectAll_Click(object sender, RoutedEventArgs e)
        {
            if (dgPersonnelSelector.ItemsSource is List<PersonnelSelectorItem> displayedItems)
            {
                bool isChecked = chkSelectAll.IsChecked ?? false;
                foreach (var item in displayedItems)
                {
                    item.IsSelected = isChecked;
                }
                UpdateCount();
            }
        }

        /// <summary>
        /// Cập nhật lại số lượng khi CheckBox của từng hàng thay đổi
        /// </summary>
        private void chkPersonnel_CheckedChanged(object sender, RoutedEventArgs e)
        {
            UpdateCount();
        }

        /// <summary>
        /// Xác nhận thêm các cán bộ đã chọn
        /// </summary>
        private void btnConfirm_Click(object sender, RoutedEventArgs e)
        {
            if (_isSingleSelect)
            {
                if (dgPersonnelSelector.SelectedItem is PersonnelSelectorItem selectedItem)
                {
                    SelectedPersonnelIds = new List<int> { selectedItem.Id };
                }
                else
                {
                    MessageBox.Show("Vui lòng chọn một cán bộ công chức trước khi xác nhận!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
            }
            else
            {
                SelectedPersonnelIds = _allPersonnel
                                        .Where(p => p.IsSelected)
                                        .Select(p => p.Id)
                                        .ToList();

                if (!SelectedPersonnelIds.Any())
                {
                    MessageBox.Show("Vui lòng chọn ít nhất một cán bộ công chức trước khi xác nhận!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
            }

            DialogResult = true;
            Close();
        }

        /// <summary>
        /// Đóng Dialog mà không thêm
        /// </summary>
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        /// <summary>
        /// Loại bỏ dấu tiếng Việt để phục vụ tìm kiếm không dấu
        /// </summary>
        private string RemoveSign(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;

            string normalizedString = text.Normalize(System.Text.NormalizationForm.FormD);
            System.Text.StringBuilder stringBuilder = new System.Text.StringBuilder();

            foreach (char c in normalizedString)
            {
                System.Globalization.UnicodeCategory unicodeCategory = System.Globalization.CharUnicodeInfo.GetUnicodeCategory(c);
                if (unicodeCategory != System.Globalization.UnicodeCategory.NonSpacingMark)
                {
                    stringBuilder.Append(c);
                }
            }

            string result = stringBuilder.ToString().Normalize(System.Text.NormalizationForm.FormC);
            result = result.Replace('đ', 'd').Replace('Đ', 'D');

            return result;
        }
    }

    /// <summary>
    /// ViewModel đại diện cho cán bộ trong danh sách chọn
    /// </summary>
    public class PersonnelSelectorItem : INotifyPropertyChanged
    {
        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected != value)
                {
                    _isSelected = value;
                    OnPropertyChanged();
                }
            }
        }

        public int Id { get; set; }
        public int STT { get; set; }
        public string FullName { get; set; } = string.Empty;
        public string IdentityCardNumber { get; set; } = string.Empty;
        public string Department { get; set; } = string.Empty;

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string? name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
