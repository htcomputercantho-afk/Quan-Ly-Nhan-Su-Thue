using System;
using System.Globalization;
using System.Windows.Data;
using TaxPersonnelManagement.Models;

namespace TaxPersonnelManagement.Converters
{
    public class RoleToDisplayNameConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is UserRole role)
            {
                return role switch
                {
                    UserRole.Admin => "Quản trị viên (Admin)",
                    UserRole.Manager => "Quản lý (Manager)",
                    UserRole.Staff => "Nhân viên (Staff)",
                    _ => role.ToString()
                };
            }
            return value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
