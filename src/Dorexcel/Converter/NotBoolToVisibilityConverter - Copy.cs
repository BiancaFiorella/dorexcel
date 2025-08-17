using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Data;
using System;

namespace Dorexcel.Converter;

public class NotBoolToOpacityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, string language)
    {
        if (value is bool boolValue)
        {
            return !boolValue ? 1.0 : 0.0;
        }

        return value;
    }

    public object ConvertBack(object value, Type targetType, object parameter, string language)
    {
        if (value is Visibility visibility)
        {
            return visibility != Visibility.Visible ? 1.0 : 0.0;
        }

        return value;
    }
}
