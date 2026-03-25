using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using TaxPersonnelManagement.Data;
using TaxPersonnelManagement.Models;
using System.Collections.ObjectModel;
using Microsoft.EntityFrameworkCore;

namespace TaxPersonnelManagement.Views
{
    public partial class SalaryDelayConfigDialog : Window
    {
        private ObservableCollection<SalaryDelayReason> _reasons = new ObservableCollection<SalaryDelayReason>();
        private SalaryDelayReason? _editingItem = null;

        public SalaryDelayConfigDialog()
        {
            InitializeComponent();
            LoadData();
            lstReasons.ItemsSource = _reasons;
        }

        private void LoadData()
        {
            _reasons.Clear();
            string dbPath = System.IO.Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "tax_personnel.db");
            TaxPersonnelManagement.App.DebugLog($"LoadData Started. DB Path: {dbPath}");
            
            try 
            {
                if (!System.IO.File.Exists(dbPath))
                {
                    TaxPersonnelManagement.App.DebugLog("Database file MISSING at path!");
                } 

                using (var context = new AppDbContext())
                {
                    try
                    {
                        var list = context.SalaryDelayReasons.OrderBy(x => x.Id).ToList();
                        foreach (var item in list)
                        {
                            _reasons.Add(item);
                        }
                    }
                    catch (Exception ex)
                    {
                         TaxPersonnelManagement.App.DebugLog($"LoadData Catch Block Hit. Error: {ex.Message}");
                         
                         // DIRECT REPAIR using Microsoft.Data.Sqlite
                         // Bypass EF Core entirely to avoid model validation errors or context state issues
                         // Use outer dbPath
                         string connectionString = $"Data Source={dbPath}";
                         
                         using (var connection = new Microsoft.Data.Sqlite.SqliteConnection(connectionString))
                         {
                             connection.Open();
                             TaxPersonnelManagement.App.DebugLog("Direct Connection Opened.");
                             
                             // 1. Create Table
                             using (var command = connection.CreateCommand())
                             {
                                 command.CommandText = @"
                                    CREATE TABLE IF NOT EXISTS SalaryDelayReasons (
                                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                                        Name TEXT NOT NULL
                                    );";
                                 command.ExecuteNonQuery();
                                 TaxPersonnelManagement.App.DebugLog("Table Create SQL Executed.");
                             }
                             
                             // 2. Check Count
                             long count = 0;
                             using (var command = connection.CreateCommand())
                             {
                                 command.CommandText = "SELECT COUNT(*) FROM SalaryDelayReasons";
                                 var res = command.ExecuteScalar();
                                 if (res != null) count = Convert.ToInt64(res);
                             }
                             
                             // 3. Seed if empty
                             if (count == 0)
                             {
                                 using (var transaction = connection.BeginTransaction())
                                 {
                                     var cmd = connection.CreateCommand();
                                     cmd.Transaction = transaction;
                                     cmd.CommandText = "INSERT INTO SalaryDelayReasons (Id, Name) VALUES (1, 'Lùi 3 tháng (Khiển trách)')";
                                     cmd.ExecuteNonQuery();
                                     
                                     cmd.CommandText = "INSERT INTO SalaryDelayReasons (Id, Name) VALUES (2, 'Lùi 6 tháng (Cảnh cáo)')";
                                     cmd.ExecuteNonQuery();
                                     
                                     cmd.CommandText = "INSERT INTO SalaryDelayReasons (Id, Name) VALUES (3, 'Lùi 12 tháng (Giáng chức/Cách chức)')";
                                     cmd.ExecuteNonQuery();
                                     
                                     cmd.CommandText = "INSERT INTO SalaryDelayReasons (Id, Name) VALUES (4, 'Nghỉ không lương')";
                                     cmd.ExecuteNonQuery();
                                     
                                     transaction.Commit();
                                 }
                             }
                         }
                        
                         // Retry load logic (again, use raw SQL or fresh context)
                         // Let's use fresh context now that table MUST exist
                         using (var retryContext = new AppDbContext())
                         {
                             var retryList = retryContext.SalaryDelayReasons.OrderBy(x => x.Id).ToList();
                             foreach (var item in retryList)
                             {
                                 _reasons.Add(item);
                             }
                         }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi tải dữ liệu (Critical): {ex.Message}\n{ex.StackTrace}", "Lỗi Hệ Thống", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtName.Text))
            {
                MessageBox.Show("Vui lòng nhập nội dung!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            using (var context = new AppDbContext())
            {
                if (_editingItem == null)
                {
                    // Add new
                    var newItem = new SalaryDelayReason { Name = txtName.Text.Trim() };
                    context.SalaryDelayReasons.Add(newItem);
                    context.SaveChanges();
                }
                else
                {
                    // Update
                    var itemToUpdate = context.SalaryDelayReasons.Find(_editingItem.Id);
                    if (itemToUpdate != null)
                    {
                        itemToUpdate.Name = txtName.Text.Trim();
                        context.SaveChanges();
                    }
                    _editingItem = null;
                    if (btnAdd.Content is StackPanel sp && sp.Children[1] is TextBlock tb)
                    {
                        tb.Text = "Thêm";
                    }
                    else
                    {
                         // Fallback if structure changes, though we defined it in XAML
                    }
                }
            }

            txtName.Clear();
            LoadData(); // reload
            txtName.Focus();
        }

        private void btnEdit_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is int id)
            {
                 var item = _reasons.FirstOrDefault(x => x.Id == id);
                 if (item != null)
                 {
                     _editingItem = item;
                     txtName.Text = item.Name;
                     
                     // Change button text to "Lưu"
                     // Quick hack to find the TextBlock inside the StackPanel inside the Button
                     if (btnAdd.Content is StackPanel sp && sp.Children.Count > 1 && sp.Children[1] is TextBlock tb)
                     {
                         tb.Text = "Lưu";
                     }
                     txtName.Focus();
                 }
            }
        }

        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is int id)
            {
                var dialog = new ConfirmDialog("Bạn có chắc chắn muốn xóa dòng này?");
                if (dialog.ShowDialog() == true)
                {
                    using (var context = new AppDbContext())
                    {
                        var item = context.SalaryDelayReasons.Find(id);
                        if (item != null)
                        {
                            context.SalaryDelayReasons.Remove(item);
                            context.SaveChanges();
                            LoadData();
                        }
                    }
                }
            }
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
