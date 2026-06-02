using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;
using TaxPersonnelManagement.Helpers;

namespace TaxPersonnelManagement.Converters
{
    public class DateFormatterConverter : IValueConverter
    {
        public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            if (value == null)
            {
                return null;
            }

            DateTime? dt = null;
            if (value is DateTime dateTime)
            {
                dt = dateTime;
            }
            else if (value is DateTimeOffset dateTimeOffset)
            {
                dt = dateTimeOffset.DateTime;
            }
            else if (value is string str && DateTime.TryParse(str, out DateTime dtParsed))
            {
                dt = dtParsed;
            }

            if (dt.HasValue)
            {
                if (parameter != null && parameter.ToString() == "Month")
                {
                    int month = dt.Value.Month;
                    return month <= 2 ? month.ToString("00") : month.ToString();
                }
                return DatePickerHelper.FormatDateForDisplay(dt.Value);
            }

            return null;
        }

        public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            return DependencyProperty.UnsetValue;
        }
    }
}
