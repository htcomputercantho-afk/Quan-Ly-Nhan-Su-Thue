using System;
using System.Globalization;
using System.Linq;
using System.Windows.Data;

namespace TaxPersonnelManagement.Converters
{
    public class PhoneFormatterConverter : IValueConverter
    {
        public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            if (value == null) return "---";
            string phone = value.ToString() ?? "";
            
            // Remove any existing dots if present (just in case they are stored differently)
            string digits = new string(phone.Where(char.IsDigit).ToArray());
            
            if (string.IsNullOrEmpty(digits)) return "---";
            
            // Format: 0XXX.XXX.XXX
            string formatted = digits.Substring(0, Math.Min(4, digits.Length));
            if (digits.Length > 4)
            {
                formatted += "." + digits.Substring(4, Math.Min(3, digits.Length - 4));
                if (digits.Length > 7)
                {
                    formatted += "." + digits.Substring(7, Math.Min(3, digits.Length - 7));
                }
            }
            
            return formatted;
        }

        public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            if (value == null) return null;
            // Strip dots when converting back (if used for two-way binding)
            return new string(value.ToString()?.Where(char.IsDigit).ToArray() ?? Array.Empty<char>());
        }
    }
}
