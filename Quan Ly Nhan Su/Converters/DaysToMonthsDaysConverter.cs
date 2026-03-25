using System;
using System.Globalization;
using System.Windows.Data;

namespace TaxPersonnelManagement.Converters
{
    public class DaysToMonthsDaysConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is double days || value is int || value is float || value is decimal)
            {
                double dValue = System.Convert.ToDouble(value);
                
                if (dValue >= 30)
                {
                    int m = (int)(dValue / 30);
                    double d = dValue % 30;
                    if (d > 0)
                    {
                        string dayStr = d == Math.Floor(d) ? ((int)d).ToString() : d.ToString("0.#");
                        return $"{m} tháng {dayStr} ngày";
                    }
                    else
                    {
                        return $"{m} tháng";
                    }
                }
                else
                {
                    string dayStr = dValue == Math.Floor(dValue) ? ((int)dValue).ToString() : dValue.ToString("0.#");
                    return $"{dayStr} ngày";
                }
            }
            return "";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
