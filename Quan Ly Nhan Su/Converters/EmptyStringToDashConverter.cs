using System;
using System.Globalization;
using System.Windows.Data;

namespace TaxPersonnelManagement.Converters
{
    public class EmptyStringToDashConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null)
                return "---";

            if (value is string str && string.IsNullOrWhiteSpace(str))
                return "---";
            
            // If it's a DateTime and it's MinValue, return "---"
            if (value is DateTime dt && dt == DateTime.MinValue)
                return "---";

            return value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
