using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using TaxPersonnelManagement.Models;

namespace TaxPersonnelManagement.Views
{
    public partial class EditMonthlyIncomeDialog : Window
    {
        public ObservableCollection<IncomeItemViewModel> IncomeItems { get; set; } = new ObservableCollection<IncomeItemViewModel>();
        public string ResultNote { get; private set; } = "";

        public EditMonthlyIncomeDialog(string incomeType, int month, string currentNote)
        {
            InitializeComponent();
            txtHeaderTitle.Text = $"Chi tiết {incomeType} - Tháng {month}";

            // Parse existing note
            var parsedItems = AnnualIncomeRowViewModel.ParseBreakdown(currentNote);
            foreach (var item in parsedItems)
            {
                IncomeItems.Add(new IncomeItemViewModel
                {
                    Amount = item.Amount,
                    Reason = item.Reason
                });
            }

            // If empty, add a default empty item for user convenience
            if (IncomeItems.Count == 0)
            {
                IncomeItems.Add(new IncomeItemViewModel());
            }

            lstIncomeItems.ItemsSource = IncomeItems;
        }

        private void btnAddItem_Click(object sender, RoutedEventArgs e)
        {
            IncomeItems.Add(new IncomeItemViewModel());
        }

        private void btnDeleteItem_Click(object sender, RoutedEventArgs e)
        {
            if (sender is FrameworkElement element && element.DataContext is IncomeItemViewModel item)
            {
                IncomeItems.Remove(item);
            }
        }

        private void btnConfirm_Click(object sender, RoutedEventArgs e)
        {
            var lines = new List<string>();
            foreach (var item in IncomeItems)
            {
                decimal amount = item.Amount;
                string reason = item.Reason?.Trim() ?? "";

                // Only save items that have a positive amount or some reason text
                if (amount > 0 || !string.IsNullOrWhiteSpace(reason))
                {
                    lines.Add($"{amount:N0} đ - {reason}");
                }
            }

            ResultNote = string.Join("\n", lines);
            this.DialogResult = true;
            this.Close();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }
    }

    public class IncomeItemViewModel : INotifyPropertyChanged
    {
        private decimal _amount;
        public decimal Amount
        {
            get => _amount;
            set
            {
                if (_amount != value)
                {
                    _amount = value;
                    OnPropertyChanged();
                    // Keep DisplayAmount in sync if updated programmatically
                    string formatted = value > 0 ? value.ToString("N0") : "";
                    if (_displayAmount != formatted)
                    {
                        _displayAmount = formatted;
                        OnPropertyChanged(nameof(DisplayAmount));
                    }
                }
            }
        }

        private string _displayAmount = "";
        public string DisplayAmount
        {
            get => _displayAmount;
            set
            {
                if (_displayAmount != value)
                {
                    string clean = System.Text.RegularExpressions.Regex.Replace(value ?? "", @"[^0-9]", "");
                    if (decimal.TryParse(clean, out decimal parsed))
                    {
                        _amount = parsed;
                        _displayAmount = parsed > 0 ? parsed.ToString("N0") : "";
                    }
                    else
                    {
                        _amount = 0;
                        _displayAmount = "";
                    }
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(Amount));
                }
            }
        }

        private string _reason = "";
        public string Reason
        {
            get => _reason;
            set
            {
                if (_reason != value)
                {
                    _reason = value;
                    OnPropertyChanged();
                }
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string? name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
