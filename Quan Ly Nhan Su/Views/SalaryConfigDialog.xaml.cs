using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Text.RegularExpressions;
using TaxPersonnelManagement.Data;
using TaxPersonnelManagement.Models;
using System.Collections.ObjectModel;

namespace TaxPersonnelManagement.Views
{
    public partial class SalaryConfigDialog : Window
    {
        private ObservableCollection<RankSalarySpec> _specs = new ObservableCollection<RankSalarySpec>();

        public SalaryConfigDialog()
        {
            InitializeComponent();
            LoadRanks();
            lstSalarySpecs.ItemsSource = _specs;
        }

        private void LoadRanks()
        {
            using (var context = new AppDbContext())
            {
                var ranks = context.Ranks.OrderBy(r => r.Code).ToList();
                cboRank.ItemsSource = ranks;
            }
        }

        private void cboRank_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cboRank.SelectedItem is Rank selectedRank)
            {
                lblEmptyState.Visibility = Visibility.Collapsed;
                LoadSpecs(selectedRank.Code);
            }
            else
            {
                lblEmptyState.Visibility = Visibility.Visible;
                _specs.Clear();
            }
        }

        private void LoadSpecs(string rankCode)
        {
            _specs.Clear();
            using (var context = new AppDbContext())
            {
                var list = context.RankSalarySpecs
                                  .Where(s => s.RankCode == rankCode)
                                  .OrderBy(s => s.Coefficient) // Order by coeff
                                  .ToList();
                
                foreach (var item in list)
                {
                    _specs.Add(item);
                }
            }
        }

        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            if (cboRank.SelectedItem is not Rank selectedRank)
            {
                MessageBox.Show("Vui lòng chọn Mã ngạch trước!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(txtSalaryStep.Text))
            {
                MessageBox.Show("Vui lòng nhập Bậc lương!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Normalize input: replace comma with dot
            string input = txtCoefficient.Text.Replace(",", ".");
            
            // Remove any characters that are not digits or dots
            string cleanInput = Regex.Replace(input, "[^0-9.]", "");

            // Count dots
            int dotCount = cleanInput.Count(c => c == '.');
            if (dotCount > 1)
            {
                MessageBox.Show("Hệ số lương không hợp lệ (có nhiều dấu chấm)!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!double.TryParse(cleanInput, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double coeff))
            {
                MessageBox.Show("Hệ số lương không hợp lệ!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            using (var context = new AppDbContext())
            {
                var newSpec = new RankSalarySpec
                {
                    RankCode = selectedRank.Code,
                    SalaryStep = txtSalaryStep.Text,
                    Coefficient = coeff
                };
                context.RankSalarySpecs.Add(newSpec);
                context.SaveChanges();
                
                // Reload to get ID
                LoadSpecs(selectedRank.Code);
                
                // Clear inputs
                txtSalaryStep.Clear();
                txtCoefficient.Clear();
                txtSalaryStep.Focus();
            }
        }

        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is int id)
            {
                if (MessageBox.Show("Bạn có chắc chắn muốn xóa dòng này?", "Xác nhận", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    using (var context = new AppDbContext())
                    {
                        var item = context.RankSalarySpecs.Find(id);
                        if (item != null)
                        {
                            context.RankSalarySpecs.Remove(item);
                            context.SaveChanges();
                            
                            // Remove from UI list directly
                            var uiItem = _specs.FirstOrDefault(s => s.Id == id);
                            if (uiItem != null) _specs.Remove(uiItem);
                        }
                    }
                }
            }
        }

        private void NumberValidationTextBox(object sender, TextCompositionEventArgs e)
        {
            Regex regex = new Regex("[^0-9.,]+"); // Allow dot and comma
            e.Handled = regex.IsMatch(e.Text);
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
