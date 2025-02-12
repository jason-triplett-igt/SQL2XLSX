using Microsoft.Data.SqlClient;
using System;
using System.Globalization;
using System.Windows.Data;

namespace SQLScript2XLSX_2
{
    public class ConnectionStringMaskConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string connectionString)
            {
                var builder = new SqlConnectionStringBuilder(connectionString);
                if (builder.ContainsKey("Password") && builder.IntegratedSecurity != true)
                {
                    builder.Password = new string('*', builder.Password.Length);
                }
                return builder.ToString();
            }
            return value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
