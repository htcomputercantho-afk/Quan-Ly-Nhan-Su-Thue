using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
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
            txtHeaderTitle.Text = $"CHI TIẾT {incomeType.ToUpper()} - THÁNG {month}";

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

        private void AmountTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (sender is TextBox textBox && textBox.DataContext is IncomeItemViewModel item)
            {
                textBox.TextChanged -= AmountTextBox_TextChanged;
                try
                {
                    int oldCaret = textBox.SelectionStart;
                    string oldText = textBox.Text ?? "";

                    int digitsBeforeCaret = 0;
                    for (int i = 0; i < Math.Min(oldCaret, oldText.Length); i++)
                    {
                        if (char.IsDigit(oldText[i]))
                            digitsBeforeCaret++;
                    }

                    string clean = Regex.Replace(oldText, @"[^0-9]", "");
                    if (decimal.TryParse(clean, out decimal parsed))
                    {
                        var viCulture = CultureInfo.GetCultureInfo("vi-VN");
                        string formatted = parsed.ToString("N0", viCulture);

                        textBox.Text = formatted;

                        int newCaret = 0;
                        int counted = 0;
                        for (int i = 0; i < formatted.Length; i++)
                        {
                            if (char.IsDigit(formatted[i]))
                            {
                                counted++;
                            }
                            if (counted == digitsBeforeCaret)
                            {
                                newCaret = i + 1;
                                break;
                            }
                        }
                        if (digitsBeforeCaret == 0) newCaret = 0;
                        else if (counted < digitsBeforeCaret) newCaret = formatted.Length;

                        textBox.SelectionStart = Math.Max(0, Math.Min(newCaret, formatted.Length));

                        item.SetRaw(parsed, formatted);
                    }
                    else
                    {
                        textBox.Text = "";
                        item.SetRaw(0, "");
                    }
                }
                finally
                {
                    textBox.TextChanged += AmountTextBox_TextChanged;
                }
            }
        }

        private void btnConfirm_Click(object sender, RoutedEventArgs e)
        {
            var lines = new List<string>();
            var viCulture = CultureInfo.GetCultureInfo("vi-VN");
            foreach (var item in IncomeItems)
            {
                decimal amount = item.Amount;
                string reason = item.Reason?.Trim() ?? "";

                // Only save items that have a positive amount or some reason text
                if (amount > 0 || !string.IsNullOrWhiteSpace(reason))
                {
                    if (amount > 0)
                    {
                        if (!string.IsNullOrWhiteSpace(reason))
                        {
                            lines.Add($"{amount.ToString("N0", viCulture)} đ - {reason}");
                        }
                        else
                        {
                            lines.Add($"{amount.ToString("N0", viCulture)} đ");
                        }
                    }
                    else if (!string.IsNullOrWhiteSpace(reason))
                    {
                        lines.Add(reason);
                    }
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
                    var viCulture = CultureInfo.GetCultureInfo("vi-VN");
                    string formatted = value > 0 ? value.ToString("N0", viCulture) : "";
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
                    string clean = Regex.Replace(value ?? "", @"[^0-9]", "");
                    if (decimal.TryParse(clean, out decimal parsed))
                    {
                        _amount = parsed;
                        var viCulture = CultureInfo.GetCultureInfo("vi-VN");
                        _displayAmount = parsed.ToString("N0", viCulture);
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

        public void SetRaw(decimal amount, string displayAmount)
        {
            _amount = amount;
            _displayAmount = displayAmount;
            OnPropertyChanged(nameof(Amount));
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
