using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;

namespace TaxPersonnelManagement.Helpers
{
    public static class DatePickerHelper
    {
        // Attached Property to enable the fix in XAML easily
        public static readonly DependencyProperty FixCalendarLocaleProperty =
            DependencyProperty.RegisterAttached(
                "FixCalendarLocale",
                typeof(bool),
                typeof(DatePickerHelper),
                new PropertyMetadata(false, OnFixCalendarLocaleChanged));

        public static bool GetFixCalendarLocale(DependencyObject obj) => (bool)obj.GetValue(FixCalendarLocaleProperty);
        public static void SetFixCalendarLocale(DependencyObject obj, bool value) => obj.SetValue(FixCalendarLocaleProperty, value);

        // Attached Property for Smart Date Entry (typing 10102026 -> 10/10/2026)
        public static readonly DependencyProperty EnableSmartDateEntryProperty =
            DependencyProperty.RegisterAttached(
                "EnableSmartDateEntry",
                typeof(bool),
                typeof(DatePickerHelper),
                new PropertyMetadata(false, OnEnableSmartDateEntryChanged));

        public static bool GetEnableSmartDateEntry(DependencyObject obj) => (bool)obj.GetValue(EnableSmartDateEntryProperty);
        public static void SetEnableSmartDateEntry(DependencyObject obj, bool value) => obj.SetValue(EnableSmartDateEntryProperty, value);

        private static void OnFixCalendarLocaleChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is DatePicker dp)
            {
                if ((bool)e.NewValue)
                {
                    dp.CalendarOpened += DatePicker_CalendarOpened;
                }
                else
                {
                    dp.CalendarOpened -= DatePicker_CalendarOpened;
                }
            }
        }

        private static void OnEnableSmartDateEntryChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is DatePicker dp)
            {
                if ((bool)e.NewValue)
                {
                    dp.DateValidationError += DatePicker_DateValidationError;
                    dp.PreviewKeyUp += DatePicker_PreviewKeyUp;
                }
                else
                {
                    dp.DateValidationError -= DatePicker_DateValidationError;
                    dp.PreviewKeyUp -= DatePicker_PreviewKeyUp;
                }
            }
        }

        private static void DatePicker_DateValidationError(object? sender, DatePickerDateValidationErrorEventArgs e)
        {
            if (sender is DatePicker dp)
            {
                string input = e.Text;
                if (string.IsNullOrEmpty(input)) return;

                // Remove common separators to extract digits
                string digits = new string(input.Where(char.IsDigit).ToArray());

                DateTime dt;
                bool success = false;

                // Try common numeric-only formats
                if (digits.Length == 8)
                {
                    success = DateTime.TryParseExact(digits, "ddMMyyyy", null, System.Globalization.DateTimeStyles.None, out dt);
                }
                else if (digits.Length == 6)
                {
                    success = DateTime.TryParseExact(digits, "ddMMyy", null, System.Globalization.DateTimeStyles.None, out dt);
                }
                else if (digits.Length == 4)
                {
                    // Assume ddMM of current year
                    success = DateTime.TryParseExact(digits + DateTime.Now.Year.ToString(), "ddMMyyyy", null, System.Globalization.DateTimeStyles.None, out dt);
                }
                else
                {
                    // Try standard parsing as fallback
                    success = DateTime.TryParse(input, out dt);
                }

                if (success)
                {
                    dp.SelectedDate = dt;
                    e.ThrowException = false;

                    // Force update the text box to the formatted date
                    dp.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Input, new Action(() =>
                    {
                        var textBox = FindVisualChildren<DatePickerTextBox>(dp).FirstOrDefault();
                        if (textBox != null)
                        {
                            textBox.Text = dt.ToString("dd/MM/yyyy");
                        }
                    }));
                }
            }
        }

        private static void DatePicker_PreviewKeyUp(object? sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.OriginalSource is TextBox textBox)
            {
                // Only process numeric keys
                bool isNumeric = (e.Key >= System.Windows.Input.Key.D0 && e.Key <= System.Windows.Input.Key.D9) ||
                                 (e.Key >= System.Windows.Input.Key.NumPad0 && e.Key <= System.Windows.Input.Key.NumPad9);

                if (isNumeric)
                {
                    string text = textBox.Text;
                    // If length is exactly 2 (e.g. "12") or 5 (e.g. "12/04"), append "/" safely
                    if ((text.Length == 2 || text.Length == 5) && !text.EndsWith("/"))
                    {
                        textBox.Text = text + "/";
                        textBox.CaretIndex = textBox.Text.Length; // Move caret to end
                    }
                }
            }
        }

        private static void DatePicker_CalendarOpened(object? sender, RoutedEventArgs e)
        {
            if (sender is DatePicker dp)
            {
                // Wait for the popup to be fully rendered
                dp.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Loaded, new Action(() =>
                {
                    FixCalendarDayHeaders(dp);
                }));
            }
        }

        private static void FixCalendarDayHeaders(DatePicker dp)
        {
            // Find the popup through the template
            var popup = dp.Template?.FindName("PART_Popup", dp) as Popup;
            if (popup == null || !popup.IsOpen || popup.Child == null) return;

            // Vietnamese day names starting from Sunday
            string[] baseDayNames = { "CN", "T2", "T3", "T4", "T5", "T6", "T7" };
            
            // Rotate the array based on FirstDayOfWeek (Sunday=0, Monday=1, ...)
            int firstDayIndex = (int)dp.FirstDayOfWeek;
            string[] rotatedDayNames = new string[7];
            for (int i = 0; i < 7; i++)
            {
                rotatedDayNames[i] = baseDayNames[(i + firstDayIndex) % 7];
            }

            // Walk the visual tree inside the popup to find headers
            var allTextBlocks = FindVisualChildren<TextBlock>(popup.Child);
            
            // Filters only for single-character headers "T" (Mon-Sat) or "C" (Sun)
            var dayHeaders = allTextBlocks
                .Where(tb => tb.Text != null && tb.Text.Length == 1
                             && (tb.Text == "T" || tb.Text == "C"))
                .ToList();

            // Replace them with properly rotated names
            if (dayHeaders.Count >= 7)
            {
                for (int i = 0; i < 7; i++)
                    dayHeaders[i].Text = rotatedDayNames[i];
            }
        }

        private static IEnumerable<T> FindVisualChildren<T>(DependencyObject parent) where T : DependencyObject
        {
            if (parent == null) yield break;
            int count = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < count; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T t) yield return t;
                foreach (var grandChild in FindVisualChildren<T>(child))
                    yield return grandChild;
            }
        }
    }
}
