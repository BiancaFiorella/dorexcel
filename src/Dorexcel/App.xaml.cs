using Dorexcel.Windows;
using Microsoft.UI.Xaml;
using OfficeOpenXml;
using System;

namespace Dorexcel;

public partial class App : Application
{
    public static Window? Window;

    static App()
    {
        Environment.SetEnvironmentVariable("MICROSOFT_WINDOWSAPPRUNTIME_BASE_DIRECTORY", AppContext.BaseDirectory);
    }

    public App()
    {        
        ExcelPackage.License.SetNonCommercialPersonal(Guid.NewGuid().ToString());
        InitializeComponent();
    }

    protected override void OnLaunched(Microsoft.UI.Xaml.LaunchActivatedEventArgs args)
    {
        Window = new MainWindow();
        Window.Activate();
    }
}
