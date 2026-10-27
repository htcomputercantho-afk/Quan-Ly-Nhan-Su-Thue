using System;
using System.Globalization;
using System.Windows.Data;

namespace TaxPersonnelManagement.Converters
{
    public class DurationConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values.Length < 1) return "";

            DateTime? startDate = values[0] as DateTime?;
            if (!startDate.HasValue) return "";

            DateTime end = DateTime.Now;
            if (values.Length >= 2 && values[1] is DateTime d1)
            {
                end = d1;
            }

            if (values.Length >= 3 && values[2] is DateTime d2)
            {
                if (d2 < end) end = d2;
            }

            if (end < startDate.Value) return "0 năm 0 tháng";

            DateTime start = startDate.Value;

            int years = end.Year - start.Year;
            if (start.Date > end.AddYears(-years)) years--;

            DateTime tmpDate = start.AddYears(years);
            int months = 0;
            while (tmpDate.AddMonths(1) <= end)
            {
                months++;
                tmpDate = tmpDate.AddMonths(1);
            }

            return $"{years} năm {months} tháng";
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
