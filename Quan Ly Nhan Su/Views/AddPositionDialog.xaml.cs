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
            LoadPositions();
        }

        private void LoadPositions()
        {
            using (var context = new AppDbContext())
            {
                // Ensure table exists just in case
                try { context.Database.EnsureCreated(); } catch { }
                
                var list = context.Positions.OrderBy(p => p.Name).ToList();
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
                MessageBox.Show("Vui lòng nhập tên chức vụ!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            using (var context = new AppDbContext())
            {
                if (_editingPosition == null)
                {
                    if (context.Positions.Any(d => d.Name == newName))
                    {
                        MessageBox.Show("Chức vụ này đã tồn tại!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    var newPos = new Position { Name = newName };
                    context.Positions.Add(newPos);
                }
                else
                {
                    var p = context.Positions.Find(_editingPosition.Id);
                    if (p != null)
                    {
                        p.Name = newName;
                    }
                    _editingPosition = null;
                }
                context.SaveChanges();
            }

            txtPositionName.Clear();
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
