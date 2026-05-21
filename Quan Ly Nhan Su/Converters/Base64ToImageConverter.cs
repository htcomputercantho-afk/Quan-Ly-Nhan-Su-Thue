using System;
using System.Globalization;
using System.IO;
using System.Windows.Data;
using System.Windows.Media.Imaging;
using TaxPersonnelManagement.Helpers;

namespace TaxPersonnelManagement.Converters
{
    public class Base64ToImageConverter : IValueConverter
    {
        public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            try
            {
                var base64 = value as string;
                if (string.IsNullOrEmpty(base64))
                    return null;

                byte[] binaryData = System.Convert.FromBase64String(base64);
                return ImageHelper.LoadAndOrientImage(binaryData);
            }
            catch
            {
                return null;
            }
        }

        public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            return Binding.DoNothing;
        }
    }
}
