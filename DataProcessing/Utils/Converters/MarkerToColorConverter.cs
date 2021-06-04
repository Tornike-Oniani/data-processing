using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;
using System.Windows.Media;

namespace DataProcessing.Converters
{
    class MarkerToColorConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if ((bool)value) { return (SolidColorBrush)new BrushConverter().ConvertFrom("#fae466"); }

            SolidColorBrush brush = (SolidColorBrush)new BrushConverter().ConvertFrom("#ffffff");
            brush.Opacity = 0;
            return brush;
            //return (SolidColorBrush)new BrushConverter().ConvertFrom("#ffffff");
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
