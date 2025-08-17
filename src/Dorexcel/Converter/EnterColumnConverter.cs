
using Microsoft.UI.Xaml.Data;
using System;

namespace Dorexcel.Converter;

public class EnterColumnConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, string language)
    {
        if(value is string strValue)
        {
            return "Ingrese " + strValue.Trim();
        }

        return value;
    }

    public object ConvertBack(object value, Type targetType, object parameter, string language)
    {                
        if (value is string strValue)
        {
            return strValue.Replace("Ingrese ", "").Trim();
        }

        return value;
    }
}
