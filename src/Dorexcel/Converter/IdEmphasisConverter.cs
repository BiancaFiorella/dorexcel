
using Microsoft.UI.Xaml.Data;
using System;

namespace Dorexcel.Converter;

public class IdEmphasisConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, string language)
    {
        if(value is string strValue)
        {
            return strValue.Trim() + " (Identidad)";
        }

        return value;
    }

    public object ConvertBack(object value, Type targetType, object parameter, string language)
    {        
        if (value is string strValue)
        {
            return strValue.Replace(" (Identidad)", "").Trim();
        }

        return value;
    }
}
