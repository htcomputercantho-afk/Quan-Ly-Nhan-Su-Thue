using System;
using System.Globalization;
using System.Windows.Data;

namespace TaxPersonnelManagement.Converters
{
    public class MajorUniversityConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values == null || values.Length < 2)
                return "---";

            string major = values[0] as string;
            string university = values[1] as string;

            bool hasMajor = !string.IsNullOrWhiteSpace(major);
            bool hasUni = !string.IsNullOrWhiteSpace(university);

            if (hasMajor && hasUni)
                return $"{major} - {university}";
            
            if (hasMajor)
                return major;

            if (hasUni)
                return university;

            return "---";
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
