using System;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Windows.Data;

namespace TaxPersonnelManagement.Converters
{
    public class CurrencyFormatConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is decimal val)
            {
                return val.ToString("N0", CultureInfo.GetCultureInfo("vi-VN"));
            }
            return "0";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string str)
            {
                string clean = Regex.Replace(str, @"[^0-9]", "");
                if (decimal.TryParse(clean, out decimal parsed))
                {
                    return parsed;
                }
            }
            return 0m;
        }
    }
}
