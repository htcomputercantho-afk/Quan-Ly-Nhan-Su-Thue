using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using TaxPersonnelManagement.Data;
using TaxPersonnelManagement.Models;

namespace TaxPersonnelManagement.Views
{
    public partial class PlanningTermManagementDialog : Window
    {
        public ObservableCollection<PlanningTerm> PlanningTerms { get; set; } = new ObservableCollection<PlanningTerm>();
        public string? SelectedTermName { get; private set; }
        private PlanningTerm? _editingTerm = null;

        public PlanningTermManagementDialog()
        {
            InitializeComponent();
            LoadTerms();
            txtFromYear.Focus();
        }

        private void LoadTerms()
        {
            try
            {
                using (var context = new AppDbContext())
                {
                    var list = context.PlanningTerms
                                      .OrderByDescending(t => t.TermName)
                                      .ToList();

                    PlanningTerms = new ObservableCollection<PlanningTerm>(list);
                    lstPlanningTerms.ItemsSource = PlanningTerms;
                }
            }
            catch (Exception ex)
            {
                var warning = new WarningWindow($"Lỗi tải danh sách nhiệm kỳ: {ex.Message}", "Lỗi");
                warning.Owner = this;
                warning.ShowDialog();
            }
        }

        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            var fromStr = txtFromYear.Text.Trim();
            var toStr = txtToYear.Text.Trim();
            if (string.IsNullOrWhiteSpace(fromStr) || string.IsNullOrWhiteSpace(toStr))
            {
                new WarningWindow("Vui lòng nhập đầy đủ năm bắt đầu và năm kết thúc!", "Thông báo").ShowDialog();
                return;
            }

            if (fromStr.Length != 4 || toStr.Length != 4 || !int.TryParse(fromStr, out int fromYear) || !int.TryParse(toStr, out int toYear))
            {
                new WarningWindow("Năm nhập vào phải gồm đúng 4 chữ số!", "Thông báo").ShowDialog();
                return;
            }

            if (toYear <= fromYear)
            {
                new WarningWindow("Năm kết thúc phải lớn hơn năm bắt đầu!", "Thông báo").ShowDialog();
                return;
            }

            var newTermName = $"{fromYear}-{toYear}";

            try
            {
                using (var context = new AppDbContext())
                {
                    if (_editingTerm == null)
                    {
                        // Thêm mới
                        if (context.PlanningTerms.Any(t => t.TermName == newTermName))
                        {
                            new WarningWindow("Nhiệm kỳ này đã tồn tại!", "Thông báo").ShowDialog();
                            return;
                        }

                        var term = new PlanningTerm { TermName = newTermName };
                        context.PlanningTerms.Add(term);
                    }
                    else
                    {
                        // Cập nhật
                        if (context.PlanningTerms.Any(t => t.TermName == newTermName && t.Id != _editingTerm.Id))
                        {
                            new WarningWindow("Nhiệm kỳ này đã tồn tại!", "Thông báo").ShowDialog();
                            return;
                        }

                        var term = context.PlanningTerms.Find(_editingTerm.Id);
                        if (term != null)
                        {
                            string oldName = term.TermName;
                            term.TermName = newTermName;

                            // Cập nhật đồng bộ các bản ghi quy hoạch đang sử dụng nhiệm kỳ cũ
                            if (oldName != newTermName)
                            {
                                var matchingRecords = context.PlanningRecords.Where(r => r.PlanningTerm == oldName).ToList();
                                foreach (var record in matchingRecords)
                                {
                                    record.PlanningTerm = newTermName;
                                }
                            }
                        }
                        _editingTerm = null;
                    }
                    context.SaveChanges();
                }

                // Reset form nhập
                txtFromYear.Clear();
                txtToYear.Clear();
                txtFromYear.Focus();

                // Trả nút thêm về trạng thái bình thường
                btnAdd.Background = (System.Windows.Media.Brush)Application.Current.Resources["PrimaryHueMidBrush"];
                btnAdd.Content = new MaterialDesignThemes.Wpf.PackIcon { Kind = MaterialDesignThemes.Wpf.PackIconKind.Plus, Width = 24, Height = 24 };
                btnAdd.ToolTip = "Thêm nhiệm kỳ";

                LoadTerms();
            }
            catch (Exception ex)
            {
                var warning = new WarningWindow($"Lỗi lưu dữ liệu: {ex.Message}", "Lỗi");
                warning.Owner = this;
                warning.ShowDialog();
            }
        }

        private void btnEdit_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is PlanningTerm term)
            {
                _editingTerm = term;
                var parts = term.TermName.Split('-');
                if (parts.Length == 2)
                {
                    txtFromYear.Text = parts[0].Trim();
                    txtToYear.Text = parts[1].Trim();
                }
                else
                {
                    txtFromYear.Text = term.TermName;
                    txtToYear.Clear();
                }
                txtFromYear.Focus();

                // Đổi nút thêm thành nút lưu
                btnAdd.Content = new MaterialDesignThemes.Wpf.PackIcon { Kind = MaterialDesignThemes.Wpf.PackIconKind.ContentSave, Width = 24, Height = 24 };
                btnAdd.ToolTip = "Lưu thay đổi";
                btnAdd.Background = System.Windows.Media.Brushes.Orange;
            }
        }

        private void YearTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            // Chỉ cho phép nhập số
            e.Handled = !int.TryParse(e.Text, out _);
        }

        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is PlanningTerm term)
            {
                try
                {
                    using (var context = new AppDbContext())
                    {
                        var t = context.PlanningTerms.Find(term.Id);
                        if (t != null)
                        {
                            // Kiểm tra xem nhiệm kỳ có đang được sử dụng không
                            var count = context.PlanningRecords.Count(r => r.PlanningTerm == t.TermName);
                            if (count > 0)
                            {
                                var confirm = new ConfirmWindow(
                                    $"Nhiệm kỳ '{t.TermName}' đang được sử dụng bởi {count} bản ghi quy hoạch. Nếu xóa, các bản ghi này sẽ bị bỏ trống nhiệm kỳ. Bạn có chắc chắn muốn xóa?", 
                                    "Xác nhận xóa");

                                if (confirm.ShowDialog() == true)
                                {
                                    // Bỏ liên kết nhiệm kỳ ở các bản ghi quy hoạch liên quan
                                    var records = context.PlanningRecords.Where(r => r.PlanningTerm == t.TermName).ToList();
                                    foreach (var r in records)
                                    {
                                        r.PlanningTerm = null;
                                    }

                                    context.PlanningTerms.Remove(t);
                                    context.SaveChanges();
                                    LoadTerms();
                                }
                            }
                            else
                            {
                                var confirm = new ConfirmWindow($"Bạn có chắc chắn muốn xóa nhiệm kỳ '{t.TermName}'?", "Xác nhận xóa");
                                if (confirm.ShowDialog() == true)
                                {
                                    context.PlanningTerms.Remove(t);
                                    context.SaveChanges();
                                    LoadTerms();
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    var warning = new WarningWindow($"Lỗi khi xóa nhiệm kỳ: {ex.Message}", "Lỗi");
                    warning.Owner = this;
                    warning.ShowDialog();
                }
            }
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void ListBoxItem_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (lstPlanningTerms.SelectedItem is PlanningTerm selected)
            {
                SelectedTermName = selected.TermName;
                DialogResult = true;
                Close();
            }
        }

        private void Header_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                DragMove();
            }
        }
    }
}
