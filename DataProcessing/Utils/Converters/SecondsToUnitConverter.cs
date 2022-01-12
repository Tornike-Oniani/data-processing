using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;

namespace DataProcessing.Converters
{
    public class SecondsToUnitConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null || String.IsNullOrEmpty(value.ToString())) { return ""; }
            switch ((int)value)
            {
                case 1:
                    return "sec";
                case 60:
                    return "min";
                case 3600:
                    return "hour";
                default:
                    // Without this going back to home from workgile causes error
                    return "";
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
