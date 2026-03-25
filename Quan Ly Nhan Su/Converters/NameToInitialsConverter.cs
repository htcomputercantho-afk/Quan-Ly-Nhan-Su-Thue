using System;
using System.Globalization;
using System.Linq;
using System.Windows.Data;

namespace TaxPersonnelManagement.Converters
{
    public class NameToInitialsConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string name && !string.IsNullOrWhiteSpace(name))
            {
                var parts = name.Trim().Split(' ');
                if (parts.Length > 0)
                {
                    // Return the first letter of the LAST word (Surname usually in VN logic, or Firstname in Western)
                    // For VN Name "Huỳnh Minh Triết", usually used "T".
                    // Let's take the first character of the LAST part.
                    return parts.Last()[0].ToString().ToUpper();
                }
            }
            return "?";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
