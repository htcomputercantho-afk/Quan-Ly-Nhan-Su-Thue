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

        private static void DatePicker_CalendarOpened(object sender, RoutedEventArgs e)
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
