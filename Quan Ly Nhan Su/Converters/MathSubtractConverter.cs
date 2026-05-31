using System;
using System.Globalization;
using System.Windows.Data;

namespace TaxPersonnelManagement.Converters
{
    public class MathSubtractConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is double val)
            {
                double subtractValue = 15; // default subtract 15px
                if (parameter != null && double.TryParse(parameter.ToString(), out double parsed))
                {
                    subtractValue = parsed;
                }
                return Math.Max(0, val - subtractValue);
            }
            return value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
