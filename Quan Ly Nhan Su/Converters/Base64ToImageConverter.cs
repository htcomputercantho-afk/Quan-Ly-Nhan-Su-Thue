using System;
using System.Globalization;
using System.IO;
using System.Windows.Data;
using System.Windows.Media.Imaging;

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

                BitmapImage bi = new BitmapImage();
                bi.BeginInit();
                bi.StreamSource = new MemoryStream(binaryData);
                bi.CacheOption = BitmapCacheOption.OnLoad; // Important to load immediately
                bi.EndInit();
                bi.Freeze(); // Freezing for performance and thread safety

                return bi;
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
