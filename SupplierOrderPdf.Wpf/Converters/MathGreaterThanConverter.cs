using System;
using System.Globalization;
using System.Windows.Data;

namespace SupplierOrderPdf.Wpf.Converters
{
    /// <summary>
    /// Конвертер для сравнения числа с параметром (больше чем).
    /// Возвращает true, если значение больше параметра.
    /// </summary>
    public class MathGreaterThanConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null || parameter == null)
                return false;

            try
            {
                double val = System.Convert.ToDouble(value, CultureInfo.InvariantCulture);
                double param = System.Convert.ToDouble(parameter, CultureInfo.InvariantCulture);
                return val > param;
            }
            catch
            {
                return false;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}