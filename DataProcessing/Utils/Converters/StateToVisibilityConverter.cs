using DataProcessing.Constants;
using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace DataProcessing.Converters
{
    // This is used for specific criteria UI, because if we only have 2 states and there is no paradoxical sleep
    // we don't want user to see specific criteria setting for paradoxical sleep
    internal class StateToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            string state = (string)value;

            if (state == RecordingType.TwoStates) { return Visibility.Collapsed; }

            return Visibility.Visible;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
