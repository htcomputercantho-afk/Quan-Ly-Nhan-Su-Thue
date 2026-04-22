using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using TaxPersonnelManagement.Data;
using TaxPersonnelManagement.Models;

namespace TaxPersonnelManagement.Views
{
    public partial class AddRankDialog : Window
    {
        public string? SelectedRankCode { get; private set; }
        private Rank? _editingRank;

        public AddRankDialog()
        {
            InitializeComponent();
            LoadRanks();
        }

        private void LoadRanks()
        {
            using (var context = new AppDbContext())
            {
                var ranks = context.Ranks.OrderBy(r => r.Code).ToList();
                lstRanks.ItemsSource = ranks;
            }
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true; // Return true to signal that changes might have happened
            Close();
        }

        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtCode.Text) || string.IsNullOrWhiteSpace(txtName.Text))
            {
                new WarningWindow("Vui lòng nhập đầy đủ mã và tên ngạch!", "Thông báo").ShowDialog();
                return;
            }

            using (var context = new AppDbContext())
            {
                if (_editingRank == null)
                {
                    // Check duplicate code
                    if (context.Ranks.Any(r => r.Code == txtCode.Text))
                    {
                         new WarningWindow("Mã ngạch này đã tồn tại!", "Thông báo").ShowDialog();
                         return;
                    }

                    var newRank = new Rank { Code = txtCode.Text, Name = txtName.Text };
                    context.Ranks.Add(newRank);
                    SelectedRankCode = newRank.Code;
                }
                else
                {
                    // Edit
                    var rankToEdit = context.Ranks.Find(_editingRank.Id);
                    if (rankToEdit != null)
                    {
                        rankToEdit.Code = txtCode.Text;
                        rankToEdit.Name = txtName.Text;
                        SelectedRankCode = rankToEdit.Code;
                    }
                    _editingRank = null;
                    btnAdd.Content = new StackPanel { Orientation = Orientation.Horizontal, Children = { new MaterialDesignThemes.Wpf.PackIcon { Kind = MaterialDesignThemes.Wpf.PackIconKind.Floppy, Margin = new Thickness(0,0,8,0), VerticalAlignment = VerticalAlignment.Center }, new TextBlock { Text = "Thêm mới", VerticalAlignment = VerticalAlignment.Center } } };
                }

                context.SaveChanges();
            }

            // Clear inputs
            txtCode.Text = "";
            txtName.Text = "";
            
            LoadRanks();
        }

        private void btnEdit_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is Rank rank)
            {
                _editingRank = rank;
                txtCode.Text = rank.Code;
                txtName.Text = rank.Name;
                
                // Change button text
                btnAdd.Content = new StackPanel { Orientation = Orientation.Horizontal, Children = { new MaterialDesignThemes.Wpf.PackIcon { Kind = MaterialDesignThemes.Wpf.PackIconKind.ContentSave, Margin = new Thickness(0,0,8,0), VerticalAlignment = VerticalAlignment.Center }, new TextBlock { Text = "Cập nhật", VerticalAlignment = VerticalAlignment.Center } } };
            }
        }

        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is Rank rank)
            {
                var confirm = new ConfirmWindow($"Bạn có chắc muốn xóa ngạch '{rank.Name}' ({rank.Code})?", "Xác nhận xóa");
                if (confirm.ShowDialog() == true)
                {
                    using (var context = new AppDbContext())
                    {
                        var r = context.Ranks.Find(rank.Id);
                        if (r != null)
                        {
                            context.Ranks.Remove(r);
                            context.SaveChanges();
                        }
                    }
                    LoadRanks();
                }
            }
        }
    }
}
